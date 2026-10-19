# -*- coding: utf-8 -*-
"""
03SEC 全線多站點動態模擬器 (Production Line Dynamic Simulator - Re-entrant Edition)
================================================================================
角色定位：高階系統架構師兼 Python 資料科學專家展示工具
功能概述：
1. 開放設定「每日投入 Lot 數量 (預設 25 Lots，可自由設定 10~30 Lots 或任意數量)」。
   - 當日投料 10~30 Lots (進料間隔 X = 48 ~ 144 min/lot) 時，首站 D/B-IC D/B (5台機台，
     等效產出週期 34.1 min/lot，日產能達 42 Lots) 游刃有餘，不會成為假瓶頸。
2. 結合 pandas 讀取『03SEC CT (2+2)』全線 126 道工序與『黃光CT總表』設備容量：
   - 支援真實載板廠「回流製程 (Re-entrant Flow)」：同一個 Lot 在 2+2 疊構各層加工時，
     會依序反覆回到相同的設備站點群組 (如 AOI、Sputter 著膜、水平酸洗、電鍍銅、綠漆印刷等) 請求機台。
   - 修正 Excel 合併儲存格造成機台數量空白的問題，若某站點無指定機台則穩健預設為 1。
   - 內建 Warm-up (產線預熱期初 WIP) 機制，讓 Day 1, 00:00 開始時全線前中後段皆有在製品流動。
3. 運用 matplotlib.animation.FuncAnimation 建立高階主管視覺化動態橫向長條圖：
   - Y 軸：全線各核心製程/設備站點名稱 (照製程首次出現順序排列)。
   - X 軸：當下站點 WIP 數量 (固定上限為 30)。
   - 當站點 WIP > 15 時，長條圖自動轉為紅色警示色 (#E53E3E)。
4. 模擬結束後，計算 1440 分鐘內各站點平均 WIP，於終端機印出「=== 全線 Top 3 瓶頸站點分析 ===」。
================================================================================
"""

import argparse
import math
import os
import random
import sys
import warnings
from typing import Dict, List, Optional, Tuple

import matplotlib
import matplotlib.animation as animation
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import simpy

# 忽略非必要的警告，保持終端機輸出乾淨
warnings.filterwarnings("ignore")

# 設定 Windows 終端機輸出編碼為 UTF-8
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ==============================================================================
# 1. 全域模擬參數設定 (可由使用者自由調整或啟動時輸入)
# ==============================================================================
EXCEL_PATHS = [
    os.path.join("DOC", "SourceData.xlsx"),
    os.path.join("doc", "SourceData.xlsx"),
    "SourceData.xlsx",
]

DEFAULT_DAILY_LOTS = 28      # 預設每日投入 Lot 數量 (建議範圍：10 ~ 30 Lots/Day)
SIM_DURATION = 1440          # 模擬時間長度：1 天 (1440 分鐘)
WARMUP_DURATION = 1440       # 產線預熱時間：讓 Day 1, 00:00 啟動時全線各站皆有正常流動之期初 WIP
MONITOR_INTERVAL = 5         # 狀態紀錄間隔：每隔 5 分鐘紀錄一次全線 WIP
WIP_ALERT_THRESHOLD = 15     # WIP 警示門檻：超過 15 時長條圖變成紅色
X_AXIS_MAX = 30              # 視覺化 X 軸固定上限：30 Lots
RANDOM_SEED = 42             # 固定隨機種子，確保每次展示結果穩定可重現


# ==============================================================================
# 2. 載板廠設備群組對照與資料讀取模組 (pandas Data Processing)
# ==============================================================================
# 根據『黃光CT總表』與『03SEC CT (2+2)』定義核心製程設備站點與其真實機台數 (Capacity)
EQUIPMENT_STATION_MAP = {
    "UPPA010": ("L/F貼耐熱膠帶 (人工)", 2),
    "SM0053":  ("D/B-IC D/B (Die Bond)", 5),
    "TS0145":  ("Nikon半自動尺寸/首件量測 (TS0145)", 2),
    "CL0200":  ("ABF/SM 表面粗化 (CL0200)", 1),
    "OV0036":  ("粗化/酸洗烘烤 (OV0036)", 1),
    "LF0003":  ("ABF film 壓膜/壓合 (LF0003)", 1),
    "OV0750":  ("Oven 烘箱群-Cure/軟硬烤 (OV0750~1238)", 4),
    "UPRM001": ("ABF-film 撕耐熱膠帶 (人工)", 2),
    "DR0026":  ("LASER 雷射打孔 (DR0026/09)", 2),
    "DL0017":  ("CF4 Plasma 去膠渣 (DL0017)", 2),
    "CL0201":  ("水平酸洗線 (CL0201)", 1),
    "DL0027":  ("馗鼎 O2 Plasma (DL0027/26)", 2),
    "ET0023":  ("Cu 銅蝕刻線 (ET0023)", 1),
    "VC1548":  ("AOI 外觀自動光學檢查 (VC1548)", 1),
    "UPQC":    ("QC 品質檢驗站 (QC-1/2/3)", 3),
    "OV1239":  ("RDL 酸洗烘烤 (OV1239/40)", 2),
    "SP0043":  ("Sputter 濺鍍著膜 (SP0043/04)", 2),
    "LF0014":  ("RDL 乾膜壓合 (LF0014)", 1),
    "曝光新機": ("RDL 線路曝光 (曝光新機)", 1),
    "OV0038":  ("RDL 曝後烤/顯後烤 (OV0038)", 1),
    "DE0025":  ("RDL 乾膜顯影 (DE0025)", 1),
    "PL0084":  ("龍門雙面電鍍 Cu (PL0084)", 1),
    "ST0012":  ("RDL 水平去膜 (ST0012)", 1),
    "ET0026":  ("NiCr 蝕刻線 (ET0026)", 1),
    "PR0254":  ("Via/SM 樹脂綠漆印刷 (PR0254)", 1),
    "EX0044":  ("Via/SM 防焊對位曝光 (EX0044)", 1),
    "DE0005":  ("Via/SM 綠漆顯影 (DE0005)", 1),
    "DL0010":  ("SM 防焊層 O2 Plasma (DL0010)", 1),
    "EX0028":  ("Via/SM 全面曝光 (EX0028)", 1),
    "TS0014":  ("SM/Cu 人工金相尺寸量測 (TS0014)", 1),
    "PL0088":  ("雙面電鍍 NiAu (PL0088)", 1),
    "XR0008":  ("X-Ray NiAu 厚度量測 (XR0008)", 1),
    "UPSC069": ("Q3 扣帳數量比對 (人工)", 2),
}


def get_mock_routing_and_stations() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    提供內建的 03SEC 全線製程 Mock Data 備援資料（包含站點表與回流工藝路線表）。
    當外部 Excel 檔案無法讀取時自動啟用。
    """
    mock_steps = [
        ("UPPA010", "L/F貼耐熱膠帶-103", "L/F貼耐熱膠帶 (人工)", 2, 11.0),
        ("UPDB002", "D/B-IC D/B", "D/B-IC D/B (Die Bond)", 5, 170.5),
        ("UPMS001", "D/B-量測", "Nikon半自動尺寸/首件量測 (TS0145)", 2, 1.83),
        ("UPBO003", "B-ABF 粗化", "ABF/SM 表面粗化 (CL0200)", 1, 3.22),
        ("UPOV076", "ABF 粗化烘烤", "粗化/酸洗烘烤 (OV0036)", 1, 1.25),
        ("UPLM010", "ABF film-膠材對壓+壓合", "ABF film 壓膜/壓合 (LF0003)", 1, 11.67),
        ("UPOV104", "ABF-Half & Full Cure", "Oven 烘箱群-Cure/軟硬烤 (OV0750~1238)", 4, 40.0),
        ("UPLA012", "B-VIA1-LASER打孔", "LASER 雷射打孔 (DR0026/09)", 2, 16.92),
        ("UPDR004", "CF4 plasma 去膠渣", "CF4 Plasma 去膠渣 (DL0017)", 2, 15.0),
        ("UPCL022", "Desmear/RDL 水平酸洗 (累計11道)", "水平酸洗線 (CL0201)", 1, 45.57),
        ("UPSC038", "AOI 外觀檢查 (全線累計6道)", "AOI 外觀自動光學檢查 (VC1548)", 1, 73.0),
        ("UPSP006", "Sputter 著膜 (全線累計4道)", "Sputter 濺鍍著膜 (SP0043/04)", 2, 96.7),
        ("UPPL026", "雙面電鍍 Cu (全線累計3道)", "龍門雙面電鍍 Cu (PL0084)", 1, 31.25),
        ("UPPR005", "樹脂綠漆印刷 (全線累計3道)", "Via/SM 樹脂綠漆印刷 (PR0254)", 1, 51.0),
        ("UPPL023", "雙面電鍍 NiAu", "雙面電鍍 NiAu (PL0088)", 1, 30.0),
    ]
    routing_rows = []
    for idx, (code, p_name, st_name, cap, ct) in enumerate(mock_steps, start=1):
        routing_rows.append({
            "step_seq": idx * 100,
            "proc_code": code,
            "proc_name": p_name,
            "station_key": st_name,
            "capacity": cap,
            "ct": ct,
        })
    routing_df = pd.DataFrame(routing_rows)
    stations_df = build_station_summary_from_routing(routing_df)
    return stations_df, routing_df


def resolve_station_group(proc_code: str, proc_name: str, equip_list: List[str]) -> Tuple[str, int]:
    """
    將『03SEC CT (2+2)』中的每一道工序對應至真實的物理設備站點與機台數量 (Capacity)。
    若機台數量空白或未定義，則穩健預設為 1。
    """
    primary_eq = equip_list[0] if len(equip_list) > 0 else ""

    # 1. 特殊製程編號或人工站點精準歸類
    if proc_code.startswith("UPQC"):
        return EQUIPMENT_STATION_MAP["UPQC"]
    if proc_code == "UPPA010":
        return EQUIPMENT_STATION_MAP["UPPA010"]
    if proc_code == "UPRM001":
        return EQUIPMENT_STATION_MAP["UPRM001"]
    if proc_code == "UPSC069":
        return EQUIPMENT_STATION_MAP["UPSC069"]
    if proc_code == "UPEX032":
        # (R)B-Via-2-B面曝光 與 A-SM/B-SM 對位曝光共用防焊曝光站 EX0044
        return EQUIPMENT_STATION_MAP["EX0044"]

    # 2. 根據首選可用機台代碼對應核心設備站點
    if primary_eq in EQUIPMENT_STATION_MAP:
        return EQUIPMENT_STATION_MAP[primary_eq]

    # 3. 若為其他未列舉之新站點，自動以列出的有效機台數作為 Capacity（空白則預設為 1）
    valid_machines = [eq for eq in equip_list if eq not in ("", "-", "人工", "機台自動", "無效站點")]
    cap = len(valid_machines) if len(valid_machines) > 0 else 1
    fallback_name = f"{proc_name} ({primary_eq if primary_eq else '預設單機'})"
    return fallback_name, max(1, cap)


def build_station_summary_from_routing(routing_df: pd.DataFrame) -> pd.DataFrame:
    """
    從全線工藝路線表 (routing_df) 彙整出各核心設備站點的統計資訊：
    - 依各站點在製程中「首次出現順序」編號排列
    - 統計該站點被回流經過幾次 (pass_count)
    - 計算單次平均 C/T、單批全線累計總 C/T (total_ct_per_lot)，以及考慮機台數後的單批等效負荷 (effective_ct)
    """
    station_order = []
    for st_key in routing_df["station_key"]:
        if st_key not in station_order:
            station_order.append(st_key)

    station_rows = []
    for idx, st_key in enumerate(station_order, start=1):
        sub = routing_df[routing_df["station_key"] == st_key]
        cap = int(sub["capacity"].iloc[0]) if not sub.empty else 1
        cap = max(1, cap)  # 若空白或異常則預設為 1
        pass_count = len(sub)
        single_max_ct = float(sub["ct"].max())
        total_ct_per_lot = float(sub["ct"].sum())
        effective_ct = total_ct_per_lot / cap

        station_Name_with_seq = f"{idx:02d}. {st_key}"
        station_rows.append({
            "station_id": idx,
            "station_key": st_key,
            "station_name": station_Name_with_seq,
            "capacity": cap,
            "pass_count": pass_count,
            "single_ct": round(single_max_ct, 2),
            "ct": round(total_ct_per_lot, 2),           # 單批全線累計在此站點的總 C/T (min/lot)
            "effective_ct": round(effective_ct, 2),     # 除以機台數後的單批等效負荷時間 (min/lot)
        })

    stations_df = pd.DataFrame(station_rows)
    return stations_df


def load_process_data() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    讀取 DOC/SourceData.xlsx 中的 03SEC 全線工藝流程，
    回傳：
      1. stations_df: 依製程順序排列的核心站點表 (含站點名稱、機台數 Capacity、單批累計 C/T、回流次數)
      2. routing_df:  Lot 實際依序走過的每一道有效製程步驟 (共 120 道有效工序)
    """
    excel_path = None
    for path in EXCEL_PATHS:
        if os.path.exists(path):
            excel_path = path
            break

    if excel_path is None:
        print("[系統提示] 找不到 DOC/SourceData.xlsx，自動載入內建 03SEC Mock Data 備援資料。")
        return get_mock_routing_and_stations()

    try:
        xl = pd.ExcelFile(excel_path)
        if "03SEC CT (2+2)" not in xl.sheet_names:
            print("[系統提示] Excel 中找不到 '03SEC CT (2+2)'，自動切換至備援資料。")
            return get_mock_routing_and_stations()

        raw_df = xl.parse("03SEC CT (2+2)")
        if "製程序號" in raw_df.columns:
            raw_df = raw_df.sort_values(by="製程序號", ascending=True).reset_index(drop=True)

        equip_cols = [c for c in raw_df.columns if str(c).startswith("可用機台")]
        routing_steps = []

        for _, row in raw_df.iterrows():
            proc_code = str(row.get("製程編號", "")).strip()
            proc_name = str(row.get("製程名稱", "")).strip()
            step_seq = row.get("製程序號", 0)
            ct_val = row.get("Cycle Time \n(min/lot)", np.nan)

            # 過濾掉 C/T 為 '-' 或 0 的無效站點
            try:
                ct = float(ct_val)
            except (ValueError, TypeError):
                continue

            if pd.isna(ct) or ct <= 0:
                continue

            # 擷取該工序所有列出的可用機台
            equip_list = []
            for col in equip_cols:
                val = row.get(col, np.nan)
                if pd.notna(val) and str(val).strip() != "":
                    equip_list.append(str(val).strip())

            station_key, capacity = resolve_station_group(proc_code, proc_name, equip_list)

            routing_steps.append({
                "step_seq": int(step_seq) if pd.notna(step_seq) else 0,
                "proc_code": proc_code,
                "proc_name": proc_name,
                "station_key": station_key,
                "capacity": max(1, int(capacity)),
                "ct": round(ct, 4),
            })

        routing_df = pd.DataFrame(routing_steps)
        if routing_df.empty:
            return get_mock_routing_and_stations()

        stations_df = build_station_summary_from_routing(routing_df)

        # 將带序號的 station_name 映射回 routing_df，方便 SimPy 直接索引
        key_to_name = dict(zip(stations_df["station_key"], stations_df["station_name"]))
        routing_df["station_name"] = routing_df["station_key"].map(key_to_name)

        print(
            f"[資料載入成功] 已從 '{excel_path}' 載入 03SEC 全線 {len(routing_df)} 道有效工序，"
            f"歸納為 {len(stations_df)} 個核心製程設備站點（支援回流共用機台）。"
        )
        return stations_df, routing_df

    except Exception as exc:
        print(f"[資料讀取警告] 讀取 Excel 發生例外 ({exc})，自動切換至內建 Mock Data。")
        return get_mock_routing_and_stations()


# ==============================================================================
# 3. 離散事件模擬引擎 (SimPy Discrete-Event Simulation Engine)
# ==============================================================================
class ProductionLineSimulator:
    """
    基於 SimPy 的 03SEC 載板全線回流多站點動態模擬器。
    """

    def __init__(
        self,
        stations_df: pd.DataFrame,
        routing_df: pd.DataFrame,
        daily_lots: float = DEFAULT_DAILY_LOTS,
        sim_duration: float = SIM_DURATION,
        warmup_duration: float = WARMUP_DURATION,
        monitor_interval: float = MONITOR_INTERVAL,
    ):
        self.stations_df = stations_df.copy().reset_index(drop=True)
        self.routing_df = routing_df.copy().reset_index(drop=True)
        self.daily_lots = float(daily_lots)
        self.sim_duration = float(sim_duration)
        self.warmup_duration = float(warmup_duration)
        self.monitor_interval = float(monitor_interval)

        # 由使用者設定的「每日投入 Lot 數 (例如 10~30 Lots)」計算平均投料間隔 X (分鐘/Lot)
        self.interarrival_time = round(1440.0 / self.daily_lots, 2)

        # 全線理論最長單機累計等效 C/T (min/lot)
        self.max_effective_ct = float(self.stations_df["effective_ct"].max())
        self.bottleneck_station_Row = self.stations_df.loc[self.stations_df["effective_ct"].idxmax()]

        # 初始化 SimPy 環境與各站點 Resource
        random.seed(RANDOM_SEED)
        self.env = simpy.Environment()
        self.resources: Dict[str, simpy.Resource] = {}
        for _, row in self.stations_df.iterrows():
            st_name = row["station_name"]
            cap = int(row["capacity"])
            self.resources[st_name] = simpy.Resource(self.env, capacity=cap)

        # 紀錄歷史狀態
        self.history_records: List[Dict[str, float]] = []
        self.lots_created = 0
        self.lots_completed_during_sim = 0

    def lot_process(self, lot_id: int):
        """
        單一 Lot 的全線工藝路徑 (Routing)：
        嚴格依序走過『03SEC CT (2+2)』中的每一道工序 (共 120 道)。
        在工序 i 完成並釋放機台後，才能 request 進入工序 i+1。
        由於載板為 2+2 疊構回流製程，不同層別的工序會共同競爭同一站點的 simpy.Resource。
        """
        for _, step in self.routing_df.iterrows():
            st_name = step["station_name"]
            base_ct = float(step["ct"])
            resource = self.resources[st_name]

            # 向該工序對應的設備站點請求機台資源
            with resource.request() as req:
                yield req
                # 加入微幅實際產線機台作業波動 (±8%)，使多站點排隊與流動更貼近真實工廠動態
                actual_ct = base_ct * random.uniform(0.92, 1.08)
                yield self.env.timeout(actual_ct)

        # 若在正式模擬期間 (扣除預熱期後) 完工，計入單日產出批數
        if self.env.now >= self.warmup_duration:
            self.lots_completed_during_sim += 1

    def lot_generator(self):
        """
        源頭進料程序：
        根據使用者設定的「每日投入 Lot 數量」，平均每隔 X = 1440 / daily_lots 分鐘投入 1 個 Lot。
        """
        while True:
            self.lots_created += 1
            self.env.process(self.lot_process(self.lots_created))
            # 模擬產線批次派工的微幅時間間隔波動 (±10%)，平均間隔嚴格等於 self.interarrival_time
            next_interval = self.interarrival_time * random.uniform(0.90, 1.10)
            yield self.env.timeout(next_interval)

    def monitor_process(self):
        """
        背景狀態紀錄程序 (Monitor)：
        先等待預熱期 (warmup_duration) 結束，讓線上已有合理的期初在製品分布；
        接著在正式的 1440 分鐘 (Day 1, 00:00 ~ 24:00) 內，每隔 5 分鐘紀錄一次全線所有站點的 WIP。
        """
        # 等待產線預熱完成
        if self.warmup_duration > 0:
            yield self.env.timeout(self.warmup_duration)

        # 正式開始 1440 分鐘 (每 5 分鐘一次) 的歷史紀錄
        while True:
            sim_clock = round(self.env.now - self.warmup_duration, 2)
            snapshot = {"time": sim_clock}
            for st_name, res in self.resources.items():
                # 站點當下總 WIP = 正在排隊等待機台的 Lot 數 (len(res.queue)) + 正在機台內加工的 Lot 數 (res.count)
                # 這樣不僅瓶頸站會累積排隊，非瓶頸站在 Lot 經過時也會有動態伸縮反應！
                queue_len = len(res.queue)
                in_service = res.count
                snapshot[st_name] = queue_len + in_service
                snapshot[f"{st_name}__queue"] = queue_len
            self.history_records.append(snapshot)
            yield self.env.timeout(self.monitor_interval)

    def run(self) -> pd.DataFrame:
        """
        執行模擬並回傳 1440 分鐘的歷史紀錄表。
        """
        print("=" * 80)
        print("🚀 啟動【03SEC 全線多站點動態模擬器 (支援回流共用機台 & 自訂日投料量)】")
        print("=" * 80)
        print(f"  • 設定每日投入批數      : {self.daily_lots:.1f} Lots / 天 (使用者可調範圍建議 10 ~ 30 Lots)")
        print(f"  • 源頭平均進料間隔 (X)  : 每 {self.interarrival_time:.2f} 分鐘投入 1 Lot")
        print(f"  • 首站 D/B-IC D/B 產能  : 單批 170.5m / 5台並聯 = 等效 34.10 min/lot (日產能上限 42.2 Lots -> 不會塞車!)")
        print(f"  • 全線理論最重負荷站點  : {self.bottleneck_station_Row['station_name']} (等效 {self.bottleneck_station_Row['effective_ct']:.2f} min/lot)")
        print(f"  • 產線預熱時間 (Warmup) : {int(self.warmup_duration)} 分鐘 (確保 Day 1 00:00 起全線前後段皆有 WIP 流動)")
        print(f"  • 正式模擬與監控時間    : {int(self.sim_duration)} 分鐘 (每 {int(self.monitor_interval)} 分鐘紀錄一次)")
        print("-" * 80)

        self.env.process(self.lot_generator())
        self.env.process(self.monitor_process())

        total_run_time = self.warmup_duration + self.sim_duration + 0.001
        self.env.run(until=total_run_time)

        history_df = pd.DataFrame(self.history_records)
        print(f"  • 模擬完成！正式模擬期間 (1440m) 全線完工產出: {self.lots_completed_during_sim} Lots")
        print(f"  • 歷史紀錄快照筆數      : {len(history_df)} 筆 (Day 1, 00:00 ~ 24:00)")
        print("=" * 80 + "\n")
        return history_df


# ==============================================================================
# 4. 自動識別瓶頸分析模組 (Automated Bottleneck Analysis)
# ==============================================================================
def analyze_bottlenecks(stations_df: pd.DataFrame, history_df: pd.DataFrame) -> pd.DataFrame:
    """
    計算每一個站點在 1440 分鐘內的「平均 WIP (含排隊與加工)」與「平均純排隊長度 (Queue Length)」，
    並在 Console 終端機印出：
    === 全線 Top 3 瓶頸站點分析 ===
    """
    station_names = stations_df["station_name"].tolist()
    queue_cols = [f"{name}__queue" for name in station_names]

    avg_wip_series = history_df[station_names].mean(axis=0)
    max_wip_series = history_df[station_names].max(axis=0)
    avg_queue_series = history_df[queue_cols].mean(axis=0)
    avg_queue_series.index = station_names

    analysis_df = stations_df.copy()
    analysis_df["avg_wip"] = analysis_df["station_name"].map(avg_wip_series)
    analysis_df["avg_queue"] = analysis_df["station_name"].map(avg_queue_series)
    analysis_df["max_wip"] = analysis_df["station_name"].map(max_wip_series)

    # 優先依照「平均純排隊長度 (avg_queue)」與「平均總 WIP (avg_wip)」排序找出真正的排隊瓶頸！
    # 這樣 D/B-IC D/B 雖然有 5 台機台同時在加工 (正常加工中 WIP 約 2~3 Lots，但排隊 Queue 接近 0)，
    # 就不會誤排進瓶頸榜，而是精準抓出真正塞車排隊的瓶頸站！
    ranked_df = analysis_df.sort_values(
        by=["avg_queue", "avg_wip", "effective_ct"], ascending=[False, False, False]
    ).reset_index(drop=True)

    print("=== 全線 Top 3 瓶頸站點分析 ===")
    for idx in range(min(3, len(ranked_df))):
        row = ranked_df.iloc[idx]
        print(
            f"{idx + 1}. {row['station_name']} - "
            f"平均 WIP: {row['avg_wip']:.2f} (平均排隊: {row['avg_queue']:.2f}, 峰值 WIP: {int(row['max_wip'])}) - "
            f"C/T: {row['ct']:.2f} min/lot (回流 {int(row['pass_count'])} 次, 機台數: {int(row['capacity'])}, 等效 C/T: {row['effective_ct']:.2f})"
        )
    print("=" * 65 + "\n")

    return ranked_df


# ==============================================================================
# 5. 動態視覺化模組 (Dynamic Visualization with matplotlib.animation)
# ==============================================================================
def setup_chinese_font():
    """設定 Matplotlib 正體中文字型。"""
    plt.rcParams["font.sans-serif"] = [
        "Microsoft JhengHei",
        "PMingLiU",
        "SimHei",
        "Arial Unicode MS",
        "DejaVu Sans",
    ]
    plt.rcParams["axes.unicode_minus"] = False


def format_sim_time(minutes: float) -> str:
    """將模擬分鐘數轉為 (Day X, HH:MM) 格式。"""
    total_mins = int(round(minutes))
    day = (total_mins // 1440) + 1
    rem_mins = total_mins % 1440
    hours = rem_mins // 60
    mins = rem_mins % 60
    if total_mins == 1440:
        return "Day 1, 24:00"
    return f"Day {day}, {hours:02d}:{mins:02d}"


def create_and_show_animation(
    stations_df: pd.DataFrame,
    history_df: pd.DataFrame,
    ranked_df: pd.DataFrame,
    daily_lots: float,
    interarrival_time: float,
):
    """
    利用歷史紀錄表與 matplotlib.animation.FuncAnimation 製作高階主管動態簡報圖表。
    - 橫向長條圖 (Horizontal Bar Chart)
    - Y 軸：全線 33 個核心製程/設備站點名稱 (照製程順序排列)
    - X 軸：當下 WIP 數量 (固定上限為 30)
    - 警示機制：當 WIP > 15 時，長條圖自動變為紅色 (#E53E3E)；當有排隊 (Queue > 0) 時顯示橘色提醒；正常加工為科技藍。
    """
    setup_chinese_font()

    station_names = stations_df["station_name"].tolist()
    queue_cols = [f"{name}__queue" for name in station_names]
    num_stations = len(station_names)

    fig, ax = plt.subplots(figsize=(14.5, 9.5), dpi=100)
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_facecolor("#FFFFFF")

    plt.subplots_adjust(left=0.27, right=0.95, top=0.88, bottom=0.08)

    y_pos = np.arange(num_stations)
    initial_wip = history_df.iloc[0][station_names].values.astype(float)

    COLOR_NORMAL = "#2B6CB0"     # 正常流動 (科技藍)
    COLOR_QUEUE = "#DD6B20"      # 輕微排隊中 (琥珀橘, 5 < WIP <= 15)
    COLOR_WARNING = "#E53E3E"    # 嚴重塞車警示 (警戒紅, WIP > 15)

    bars = ax.barh(
        y_pos,
        initial_wip,
        height=0.72,
        color=COLOR_NORMAL,
        edgecolor="#1A365D",
        linewidth=0.6,
        alpha=0.92,
    )

    ax.set_yticks(y_pos)
    ax.set_yticklabels(station_names, fontsize=8.5, fontweight="bold", color="#2D3748")
    ax.invert_yaxis()

    # 固定 X 軸上限為 30
    ax.set_xlim(0, X_AXIS_MAX)
    ax.set_xticks(np.arange(0, X_AXIS_MAX + 1, 5))
    ax.set_xlabel(
        "站點當下 WIP 批數 (含排隊等待 Queue + 機台加工中 Lots)",
        fontsize=11,
        fontweight="bold",
        color="#1A202C",
        labelpad=8,
    )

    # 繪製 WIP = 15 的紅色警戒虛線
    ax.axvline(
        x=WIP_ALERT_THRESHOLD,
        color=COLOR_WARNING,
        linestyle="--",
        linewidth=1.6,
        alpha=0.85,
        label=f"WIP 嚴重塞車警戒線 (> {WIP_ALERT_THRESHOLD} Lots 變紅色)",
    )

    ax.grid(axis="x", linestyle=":", alpha=0.6, color="#A0AEC0")
    ax.set_axisbelow(True)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CBD5E0")
    ax.spines["bottom"].set_color("#CBD5E0")

    value_texts = [
        ax.text(0.3, y, "", va="center", ha="left", fontsize=8, fontweight="bold", color="#4A5568")
        for y in y_pos
    ]

    fig.suptitle(
        f"【03SEC 載板產線】全線多站點 WIP 動態堆積與瓶頸監控 (日投料設定: {int(daily_lots)} Lots/天)",
        fontsize=14.5,
        fontweight="bold",
        color="#1A202C",
        y=0.96,
    )
    sub_title = ax.set_title(
        f"當前模擬時間: Day 1, 00:00  |  平均投料間隔 X = {interarrival_time} min/lot  |  全線總 WIP: 0 Lots",
        fontsize=11,
        fontweight="bold",
        color="#2C5282",
        pad=12,
    )

    # 右下角顯示 Top 3 瓶頸分析摘要盒
    top3_lines = [f"【全線 Top 3 瓶頸站點 (日投 {int(daily_lots)} Lots 模擬結果)】"]
    for idx in range(min(3, len(ranked_df))):
        r = ranked_df.iloc[idx]
        top3_lines.append(
            f"TOP {idx+1}: {r['station_name']} "
            f"(平均WIP: {r['avg_wip']:.1f} [排隊:{r['avg_queue']:.1f}] | 累計C/T: {r['ct']:.1f}m / {int(r['capacity'])}機)"
        )
    ax.text(
        0.985,
        0.025,
        "\n".join(top3_lines),
        transform=ax.transAxes,
        fontsize=8.5,
        verticalalignment="bottom",
        horizontalalignment="right",
        bbox=dict(boxstyle="round,pad=0.55", facecolor="#EDF2F7", edgecolor="#CBD5E0", alpha=0.94),
        color="#1A202C",
    )

    ax.legend(loc="upper right", frameon=True, facecolor="#FFFFFF", edgecolor="#CBD5E0", fontsize=9.5)

    def update(frame_idx: int):
        row = history_df.iloc[frame_idx]
        current_time = float(row["time"])
        wips = row[station_names].values.astype(float)
        queues = row[queue_cols].values.astype(float)
        total_wip = int(np.sum(wips))
        total_queue = int(np.sum(queues))
        alert_count = int(np.sum(wips > WIP_ALERT_THRESHOLD))

        for bar, wip_val, q_val, txt in zip(bars, wips, queues, value_texts):
            display_width = min(wip_val, X_AXIS_MAX - 0.2)
            bar.set_width(display_width)

            if wip_val > WIP_ALERT_THRESHOLD:
                bar.set_facecolor(COLOR_WARNING)
                bar.set_edgecolor("#9B2C2C")
                txt.set_color(COLOR_WARNING)
                txt.set_text(f" {int(wip_val)} (排隊:{int(q_val)}) ⚠️瓶頸!")
            elif q_val >= 3:
                bar.set_facecolor(COLOR_QUEUE)
                bar.set_edgecolor("#9C4221")
                txt.set_color("#C05621")
                txt.set_text(f" {int(wip_val)} (排隊:{int(q_val)})")
            else:
                bar.set_facecolor(COLOR_NORMAL)
                bar.set_edgecolor("#1A365D")
                txt.set_color("#2D3748")
                if wip_val > 0:
                    txt.set_text(f" {int(wip_val)}" + (f" (排:{int(q_val)})" if q_val > 0 else ""))
                else:
                    txt.set_text("")

            txt.set_x(
                min(display_width + 0.25, X_AXIS_MAX - 4.8)
                if wip_val > WIP_ALERT_THRESHOLD
                else display_width + 0.2
            )

        time_str = format_sim_time(current_time)
        sub_title.set_text(
            f"當前模擬時間: {time_str} ({int(current_time)} min)  |  "
            f"全線線上總 WIP: {total_wip} Lots (含排隊等待: {total_queue} Lots)  |  "
            f"紅色超載站點 (WIP > {WIP_ALERT_THRESHOLD}): {alert_count} 站"
        )
        sub_title.set_color(COLOR_WARNING if alert_count > 0 else "#2C5282")

        return list(bars) + value_texts + [sub_title]

    ani = animation.FuncAnimation(
        fig=fig,
        func=update,
        frames=len(history_df),
        interval=45,
        blit=False,
        repeat=True,
    )
    fig._animation_ref = ani

    print("[視覺化啟動] 正在開啟 Matplotlib 動態視覺化視窗 (關閉視窗即可結束程式)...")
    plt.show()
    return ani


# ==============================================================================
# 6. 使用者互動參數讀取與主程式進入點 (Main Entry Point)
# ==============================================================================
def get_user_daily_lots() -> float:
    """
    開放讓使用者透過命令列參數 (--lots) 或互動式輸入設定「一天投入的 Lot 數量 (建議 10~30)」。
    """
    parser = argparse.ArgumentParser(description="03SEC 全線多站點動態模擬器")
    parser.add_argument(
        "--lots",
        type=float,
        default=None,
        help=f"設定每日 (1440分鐘) 投入的 Lot 數量 (建議 10~30，預設: {DEFAULT_DAILY_LOTS})",
    )
    args, _ = parser.parse_known_args()

    if args.lots is not None and args.lots > 0:
        return float(args.lots)

    # 若在自動化測試環境或非互動式終端機，直接採用預設值
    if os.environ.get("HEADLESS_TEST", "0") == "1" or not sys.stdin.isatty():
        return float(DEFAULT_DAILY_LOTS)

    try:
        user_input = input(
            f"\n👉 請輸入每日預計投入的 Lot 數量 (實際生產建議 10 ~ 30，直接按 Enter 預設為 {DEFAULT_DAILY_LOTS} Lots): "
        ).strip()
        if user_input == "":
            return float(DEFAULT_DAILY_LOTS)
        val = float(user_input)
        if val <= 0:
            print(f"[提示] 輸入數值需大於 0，已自動使用預設值 {DEFAULT_DAILY_LOTS} Lots。")
            return float(DEFAULT_DAILY_LOTS)
        return val
    except Exception:
        return float(DEFAULT_DAILY_LOTS)


def main():
    # Step 1: 讀取 Excel 全線回流工藝路線與各核心設備站點
    stations_df, routing_df = load_process_data()

    print("\n【03SEC 全線核心設備站點與回流負荷總覽 (依製程順序排列)】")
    preview_df = stations_df[
        ["station_id", "station_name", "capacity", "pass_count", "single_ct", "ct", "effective_ct"]
    ].copy()
    preview_df.columns = [
        "站點序號",
        "站點名稱",
        "機台數(Cap)",
        "回流次數",
        "單次最大C/T",
        "單Lot累計C/T",
        "單Lot等效負荷(min)",
    ]
    print(preview_df.to_string(index=False))

    # Step 2: 取得使用者設定的每日投入 Lot 數 (10 ~ 30 Lots)
    daily_lots = get_user_daily_lots()

    # Step 3: 啟動 SimPy 模擬器 (含預熱期與 1440 分鐘正式紀錄)
    simulator = ProductionLineSimulator(
        stations_df=stations_df,
        routing_df=routing_df,
        daily_lots=daily_lots,
        sim_duration=SIM_DURATION,
        warmup_duration=WARMUP_DURATION,
        monitor_interval=MONITOR_INTERVAL,
    )
    history_df = simulator.run()

    # Step 4: 自動識別並印出全線 Top 3 瓶頸站點分析
    ranked_df = analyze_bottlenecks(stations_df, history_df)

    # Step 5: 啟動 Matplotlib 橫向長條圖動態視覺化
    if os.environ.get("HEADLESS_TEST", "0") != "1":
        create_and_show_animation(
            stations_df=stations_df,
            history_df=history_df,
            ranked_df=ranked_df,
            daily_lots=daily_lots,
            interarrival_time=simulator.interarrival_time,
        )


if __name__ == "__main__":
    main()
