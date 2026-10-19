# -*- coding: utf-8 -*-
"""
03SEC & 01SG 全線多站點動態混線模擬器 (Production Line Dynamic Simulator)
================================================================================
角色定位：高階系統架構師兼 Python 資料科學專家展示工具
功能概述：
1. 雙產品混線模擬 (03SEC 與 01SG)：
   - 透過 pandas 同時讀取『03SEC CT (2+2)』與『01SG CT (1+1)』兩條產品的全線工藝流程。
   - 預設採用 03SEC : 01SG = 7 : 3 的比例混合投料（開放使用者於啟動時或命令列自由設定比例）。
2. 開放核心生產參數設定：
   - 每日投入 Lot 數量：預設 25 Lots/天（建議範圍 10 ~ 30 Lots，可自由設定）。
   - 產品混合比例 (03SEC:01SG)：預設 7:3（可設定如 7:3、5:5、8:2、1:0 等）。
   - 機台 RUN% (設備正常運作時間百分比/稼動率)：預設 80%（代表一天有 80% 時間可正常生產，20% 為停機/保養/換線損失）。
3. 完整納入 D/B-IC D/B (Die Bond) 與全線回流共用機台 (Re-entrant Flow) 計算：
   - 將 D/B 站點與全線 33 個核心製程設備站點一併納入總 WIP 與 Top 3 瓶頸分析計算。
4. 動態視覺化儀表板 (matplotlib.animation.FuncAnimation)：
   - 橫向長條圖，Y 軸依製程順序排列，X 軸固定上限為 30。
   - 標題動態顯示當前模擬時間 (Day X, HH:MM)、混線比例、機台 RUN% 與即時 WIP。
   - 當某站點當下 WIP > 15 時，長條圖自動轉為紅色警示色 (#E53E3E)。
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
# 1. 全域模擬參數設定 (使用者可直接修改此處常數，或於執行時互動輸入)
# ==============================================================================
EXCEL_PATHS = [
    os.path.join("DOC", "SourceData.xlsx"),
    os.path.join("doc", "SourceData.xlsx"),
    "SourceData.xlsx",
]

DEFAULT_DAILY_LOTS = 25          # 預設每日投入 Lot 總數 (實際生產建議範圍：10 ~ 30 Lots/天)
DEFAULT_MIX_RATIO_STR = "7:3"    # 預設 03SEC 與 01SG 混線比例 (7:3 代表 70% 03SEC + 30% 01SG)
DEFAULT_WORK_HOURS = 20.0        # 預設每天生產時間 (小時)，預設 20 小時 (即 60 * 20 = 1200 分鐘/天)
DEFAULT_RUN_PCT = 80.0           # 預設機台 RUN% (預設 80%，故每天每台設備實際生產時間為 60 * 20 * 0.8 = 960 分鐘)
DEFAULT_TRANSFER_MIN = 2.0       # 預設站與站之間人員搬送時間 (分鐘/次，設為 0 代表不計搬送時間)
DEFAULT_TOP_N = 3                # 預設顯示 Top N 瓶頸站點數量 (可自由設定要顯示 Top 多少)

SIM_DURATION = 4320              # 正式模擬時間長度 (可依天數或分鐘設定)
WARMUP_DURATION = 1440           # 產線預熱時間：確保 Day 1, 00:00 啟動時全線皆有正常期初 WIP
MONITOR_INTERVAL = 15            # 狀態紀錄間隔 (分鐘)
WIP_ALERT_THRESHOLD = 20         # WIP 紅色警示門檻：超過此數值時長條圖變成紅色
X_AXIS_MAX = 30                  # 視覺化 X 軸固定上限：30 Lots
RANDOM_SEED = 42                 # 固定隨機種子，確保簡報展示結果穩定可重現


# ==============================================================================
# 2. 載板廠設備群組對照與雙產品資料讀取模組 (pandas Data Processing)
# ==============================================================================
# 根據『黃光CT總表』、『設備加工上限』、『03SEC CT (2+2)』與『01SG CT (1+1)』定義全線核心站點與機台數量 (Capacity)
EQUIPMENT_STATION_MAP = {
    "UPPA010": ("L/F貼耐熱膠帶 (TE0109/人工)", 2),
    "TE0109":  ("L/F貼耐熱膠帶 (TE0109/人工)", 2),
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


def resolve_station_group(proc_code: str, proc_name: str, equip_list: List[str]) -> Tuple[str, int]:
    """
    將『03SEC CT (2+2)』與『01SG CT (1+1)』中的每一道工序對應至真實的物理設備站點與機台數量 (Capacity)。
    若機台數量空白或未定義，則穩健預設為 1。
    """
    primary_eq = equip_list[0] if len(equip_list) > 0 else ""

    # 1. 特殊製程編號或人工站點精準歸類
    if proc_code.startswith("UPQC"):
        return EQUIPMENT_STATION_MAP["UPQC"]
    if proc_code in ("UPPA010", "UPPA008"):
        return EQUIPMENT_STATION_MAP["UPPA010"]
    if proc_code == "UPRM001":
        return EQUIPMENT_STATION_MAP["UPRM001"]
    if proc_code == "UPSC069":
        return EQUIPMENT_STATION_MAP["UPSC069"]
    if proc_code == "UPEX032":
        return EQUIPMENT_STATION_MAP["EX0044"]

    # 2. 根據首選可用機台代碼對應核心設備站點
    if primary_eq in EQUIPMENT_STATION_MAP:
        return EQUIPMENT_STATION_MAP[primary_eq]

    # 3. 若某站點機台數量空白或不在預設字典中，統計有效機台數，空白則預設為 1
    valid_machines = [eq for eq in equip_list if eq not in ("", "-", "人工", "機台自動", "無效站點")]
    cap = len(valid_machines) if len(valid_machines) > 0 else 1
    fallback_name = f"{proc_name} ({primary_eq if primary_eq else '預設單機'})"
    return fallback_name, max(1, cap)


def parse_product_routing_sheet(xl: pd.ExcelFile, sheet_name: str, product_code: str) -> pd.DataFrame:
    """
    解析單一產品工作表（如『03SEC CT (2+2)』或『01SG CT (1+1)』）的完整工藝流程。
    """
    raw_df = xl.parse(sheet_name)
    if "製程序號" in raw_df.columns:
        raw_df = raw_df.sort_values(by="製程序號", ascending=True).reset_index(drop=True)

    equip_cols = [c for c in raw_df.columns if str(c).startswith("可用機台")]
    routing_steps = []

    for _, row in raw_df.iterrows():
        proc_code = str(row.get("製程編號", "")).strip()
        proc_name = str(row.get("製程名稱", "")).strip()
        step_seq = row.get("製程序號", 0)
        ct_val = row.get("Cycle Time \n(min/lot)", np.nan)

        # 過濾掉 C/T 為 '-'、'無效站點' 或 0 的步驟
        try:
            ct = float(ct_val)
        except (ValueError, TypeError):
            continue

        if pd.isna(ct) or ct <= 0:
            continue

        equip_list = []
        for col in equip_cols:
            val = row.get(col, np.nan)
            if pd.notna(val) and str(val).strip() != "":
                equip_list.append(str(val).strip())

        station_key, capacity = resolve_station_group(proc_code, proc_name, equip_list)

        routing_steps.append({
            "product": product_code,
            "step_seq": int(step_seq) if pd.notna(step_seq) else 0,
            "proc_code": proc_code,
            "proc_name": proc_name,
            "station_key": station_key,
            "capacity": max(1, int(capacity)),
            "ct": round(ct, 4),
        })

    return pd.DataFrame(routing_steps)


def build_combined_station_summary(
    routing_03sec: pd.DataFrame,
    routing_01sg: pd.DataFrame,
    ratio_03sec: float,
    ratio_01sg: float,
    run_pct: float,
    work_hours: float = DEFAULT_WORK_HOURS,
) -> pd.DataFrame:
    """
    根據 03SEC 與 01SG 的工藝路線、混合比例 (例如 7:3)、每天生產時間 (例如 20 小時) 以及機台 RUN% (例如 80%)，
    建立全線各站點的彙整資料表：
    - 每天設備實際可生產時間 = 60 * work_hours * (run_pct / 100.0) (預設 60 * 20 * 0.8 = 960 分鐘/天/台)
    - 折算至 24 小時日曆模擬時間軸的綜合可用率因子 = (work_hours / 24.0) * (run_pct / 100.0)
    - 計算該站點「每日最大產能批數 (daily_cap_lots = (daily_equip_prod_mins * capacity) / weighted_ct)」
    """
    station_order = []
    for st_key in pd.concat([routing_03sec["station_key"], routing_01sg["station_key"]]):
        if st_key not in station_order:
            station_order.append(st_key)

    run_rate = max(0.05, min(1.0, run_pct / 100.0))
    work_hours_clamped = max(1.0, min(24.0, float(work_hours)))
    # 每台設備每天實際可生產分鐘數 (預設 60 * 20 * 0.8 = 960 分鐘)
    daily_equip_prod_mins = 60.0 * work_hours_clamped * run_rate
    # 對應於一天 1440 分鐘日曆時間的綜合有效運轉率
    overall_availability = daily_equip_prod_mins / 1440.0

    station_rows = []

    for idx, st_key in enumerate(station_order, start=1):
        sub_03 = routing_03sec[routing_03sec["station_key"] == st_key]
        sub_01 = routing_01sg[routing_01sg["station_key"] == st_key]

        cap = int(sub_03["capacity"].iloc[0]) if not sub_03.empty else int(sub_01["capacity"].iloc[0])
        cap = max(1, cap)

        pass_03 = len(sub_03)
        pass_01 = len(sub_01)
        ct_03 = float(sub_03["ct"].sum()) if not sub_03.empty else 0.0
        ct_01 = float(sub_01["ct"].sum()) if not sub_01.empty else 0.0

        # 依 03SEC : 01SG 混合比例計算加權單批累計 C/T 與回流次數
        weighted_ct = (ratio_03sec * ct_03) + (ratio_01sg * ct_01)
        weighted_pass = (ratio_03sec * pass_03) + (ratio_01sg * pass_01)

        # 考慮機台數 (Capacity)、每天生產時數 (work_hours) 與 機台 RUN% 後：
        # 1. 在設備實際生產時間 (960分鐘) 下的單批等效週期 = weighted_ct / (cap * run_rate)
        # 2. 該站每天 (在 60 * work_hours * run_rate 分鐘內) 最多可產出幾批 Lot
        daily_max_lots = (daily_equip_prod_mins * cap) / weighted_ct if weighted_ct > 0 else 999.0
        # 3. 折算至 1440 分鐘日曆時間軸的等效週期 = 1440 / daily_max_lots = weighted_ct / (cap * overall_availability)
        effective_ct = weighted_ct / (cap * overall_availability)

        station_rows.append({
            "station_id": idx,
            "station_key": st_key,
            "station_name": f"{idx:02d}. {st_key}",
            "capacity": cap,
            "work_hours": round(work_hours_clamped, 1),
            "run_pct": round(run_pct, 1),
            "daily_equip_mins": round(daily_equip_prod_mins, 1),
            "pass_03sec": pass_03,
            "pass_01sg": pass_01,
            "weighted_pass": round(weighted_pass, 1),
            "ct_03sec": round(ct_03, 2),
            "ct_01sg": round(ct_01, 2),
            "ct": round(weighted_ct, 2),                    # 混線加權單批累計 C/T (min/lot)
            "effective_ct": round(effective_ct, 2),         # 折算工時與 RUN% 後之等效週期 (min/lot)
            "daily_max_lots": round(daily_max_lots, 1),     # 該站每天最大可消化批數 (Lots/天)
        })

    return pd.DataFrame(station_rows)


def get_mock_dual_product_data(
    ratio_03sec: float,
    ratio_01sg: float,
    run_pct: float,
    work_hours: float = DEFAULT_WORK_HOURS,
) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """
    提供內建的 03SEC 與 01SG 雙產品 Mock Data 備援機制。
    """
    mock_03 = pd.DataFrame([
        {"product": "03SEC", "step_seq": 100, "proc_code": "UPPA010", "proc_name": "貼膠帶", "station_key": "L/F貼耐熱膠帶 (TE0109/人工)", "capacity": 2, "ct": 11.0},
        {"product": "03SEC", "step_seq": 200, "proc_code": "UPDB002", "proc_name": "D/B", "station_key": "D/B-IC D/B (Die Bond)", "capacity": 5, "ct": 170.5},
        {"product": "03SEC", "step_seq": 300, "proc_code": "UPSC038", "proc_name": "AOI", "station_key": "AOI 外觀自動光學檢查 (VC1548)", "capacity": 1, "ct": 73.0},
        {"product": "03SEC", "step_seq": 400, "proc_code": "UPSP006", "proc_name": "著膜", "station_key": "Sputter 濺鍍著膜 (SP0043/04)", "capacity": 2, "ct": 96.7},
    ])
    mock_01 = pd.DataFrame([
        {"product": "01SG", "step_seq": 100, "proc_code": "UPPA008", "proc_name": "貼膠帶", "station_key": "L/F貼耐熱膠帶 (TE0109/人工)", "capacity": 2, "ct": 11.15},
        {"product": "01SG", "step_seq": 200, "proc_code": "UPDB002", "proc_name": "D/B", "station_key": "D/B-IC D/B (Die Bond)", "capacity": 5, "ct": 118.98},
        {"product": "01SG", "step_seq": 300, "proc_code": "UPSC038", "proc_name": "AOI", "station_key": "AOI 外觀自動光學檢查 (VC1548)", "capacity": 1, "ct": 36.5},
        {"product": "01SG", "step_seq": 400, "proc_code": "UPSP003", "proc_name": "著膜", "station_key": "Sputter 濺鍍著膜 (SP0043/04)", "capacity": 2, "ct": 48.35},
    ])
    stations_df = build_combined_station_summary(
        mock_03, mock_01, ratio_03sec, ratio_01sg, run_pct, work_hours
    )
    key_to_name = dict(zip(stations_df["station_key"], stations_df["station_name"]))
    mock_03["station_name"] = mock_03["station_key"].map(key_to_name)
    mock_01["station_name"] = mock_01["station_key"].map(key_to_name)
    return stations_df, {"03SEC": mock_03, "01SG": mock_01}


def load_process_data(
    ratio_03sec: float = 0.7,
    ratio_01sg: float = 0.3,
    run_pct: float = DEFAULT_RUN_PCT,
    work_hours: float = DEFAULT_WORK_HOURS,
) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    """
    讀取 DOC/SourceData.xlsx 中的『03SEC CT (2+2)』與『01SG CT (1+1)』雙產品工藝流程。
    """
    excel_path = None
    for path in EXCEL_PATHS:
        if os.path.exists(path):
            excel_path = path
            break

    if excel_path is None:
        print("[系統提示] 找不到 DOC/SourceData.xlsx，自動載入內建雙產品 Mock Data。")
        return get_mock_dual_product_data(ratio_03sec, ratio_01sg, run_pct, work_hours)

    try:
        xl = pd.ExcelFile(excel_path)
        routing_03sec = parse_product_routing_sheet(xl, "03SEC CT (2+2)", "03SEC")
        routing_01sg = parse_product_routing_sheet(xl, "01SG CT (1+1)", "01SG")

        if routing_03sec.empty or routing_01sg.empty:
            return get_mock_dual_product_data(ratio_03sec, ratio_01sg, run_pct, work_hours)

        stations_df = build_combined_station_summary(
            routing_03sec, routing_01sg, ratio_03sec, ratio_01sg, run_pct, work_hours
        )

        key_to_name = dict(zip(stations_df["station_key"], stations_df["station_name"]))
        routing_03sec["station_name"] = routing_03sec["station_key"].map(key_to_name)
        routing_01sg["station_name"] = routing_01sg["station_key"].map(key_to_name)

        print(
            f"[資料載入成功] 已從 '{excel_path}' 載入雙產品全線工藝：\n"
            f"  • 03SEC (2+2 疊構) : 共 {len(routing_03sec)} 道有效工序 (D/B C/T = 170.50 min/lot)\n"
            f"  • 01SG  (1+1 疊構) : 共 {len(routing_01sg)} 道有效工序 (D/B C/T = 118.98 min/lot)\n"
            f"  • 整合共用設備站點 : 共 {len(stations_df)} 個核心站點 (含 02. D/B-IC D/B 完整納入計算)"
        )
        return stations_df, {"03SEC": routing_03sec, "01SG": routing_01sg}

    except Exception as exc:
        print(f"[資料讀取警告] 讀取 Excel 發生例外 ({exc})，自動切換至內建 Mock Data。")
        return get_mock_dual_product_data(ratio_03sec, ratio_01sg, run_pct, work_hours)


# ==============================================================================
# 3. 離散事件模擬引擎 (SimPy Discrete-Event Simulation Engine)
# ==============================================================================
class ProductionLineSimulator:
    """
    支援 03SEC & 01SG 比例混線、每天生產時間 (預設20小時)、機台 RUN% (預設80%)、
    以及站間人員搬送時間 (Transfer Time) 的 SimPy 動態模擬器。
    """

    def __init__(
        self,
        stations_df: pd.DataFrame,
        routings: Dict[str, pd.DataFrame],
        daily_lots: float = DEFAULT_DAILY_LOTS,
        ratio_03sec: float = 0.7,
        ratio_01sg: float = 0.3,
        work_hours: float = DEFAULT_WORK_HOURS,
        run_pct: float = DEFAULT_RUN_PCT,
        transfer_min: float = DEFAULT_TRANSFER_MIN,
        sim_duration: float = SIM_DURATION,
        warmup_duration: float = WARMUP_DURATION,
        monitor_interval: float = MONITOR_INTERVAL,
    ):
        self.stations_df = stations_df.copy().reset_index(drop=True)
        self.routings = routings
        self.daily_lots = float(daily_lots)
        self.ratio_03sec = float(ratio_03sec)
        self.ratio_01sg = float(ratio_01sg)
        self.work_hours = max(1.0, min(24.0, float(work_hours)))
        self.run_pct = float(run_pct)
        self.run_rate = max(0.05, min(1.0, self.run_pct / 100.0))
        self.transfer_min = max(0.0, float(transfer_min))

        # 每天生產分鐘數 (預設 60 * 20 = 1200 分鐘) 與 每天每台設備實際生產時間 (預設 60 * 20 * 0.8 = 960 分鐘)
        self.daily_work_mins = 60.0 * self.work_hours
        self.daily_equip_prod_mins = self.daily_work_mins * self.run_rate
        # 對應至 1440 分鐘日曆時間軸的機台總有效可用率 (預設 960 / 1440 = 66.67%)
        self.overall_availability = self.daily_equip_prod_mins / 1440.0

        self.sim_duration = float(sim_duration)
        self.warmup_duration = float(warmup_duration)
        self.monitor_interval = float(monitor_interval)

        # 根據每日投入總 Lot 數換算源頭平均進料間隔 X (分鐘/Lot)
        self.interarrival_time = round(1440.0 / self.daily_lots, 2)

        # 找出考慮混線比例、每天生產時數與 RUN% 後，全線理論等效負荷最重的站點
        self.bottleneck_row = self.stations_df.loc[self.stations_df["effective_ct"].idxmax()]

        # 初始化 SimPy 環境與各站點 Resource
        random.seed(RANDOM_SEED)
        self.env = simpy.Environment()
        self.resources: Dict[str, simpy.Resource] = {}
        for _, row in self.stations_df.iterrows():
            st_name = row["station_name"]
            cap = int(row["capacity"])
            self.resources[st_name] = simpy.Resource(self.env, capacity=cap)

        # 紀錄歷史狀態、在途搬送批數與產出統計
        self.history_records: List[Dict[str, float]] = []
        self.lots_in_transfer = 0
        self.lots_created = {"03SEC": 0, "01SG": 0}
        self.lots_completed = {"03SEC": 0, "01SG": 0}

    def select_next_product(self, seq_index: int) -> str:
        """
        依據設定的 03SEC : 01SG 比例 (例如 7:3)，採用平準化交錯投料 (Heijunka / Bresenham 配比法)，
        確保實際投入的 03SEC 與 01SG 比例精準符合設定值。
        """
        if self.ratio_03sec <= 0:
            return "01SG"
        if self.ratio_01sg <= 0:
            return "03SEC"

        expected_03sec = seq_index * self.ratio_03sec
        actual_03sec = self.lots_created["03SEC"]
        if actual_03sec < expected_03sec:
            return "03SEC"
        return "01SG"

    def lot_process(self, lot_id: int, product_type: str):
        """
        單一 Lot (03SEC 或 01SG) 的全線工藝路徑 (Routing)：
        1. 依序走過該產品專屬的 Excel 工藝步驟表。
        2. 每天生產時間 (work_hours) 與機台 RUN% 效應：
           每天每台設備實際生產時間為 `60 * work_hours * run_rate` (預設 60*20*0.8 = 960 分鐘)，
           折算至全天模擬時間軸，每道工序實際耗用時間為 `base_ct / overall_availability`。
        3. 站與站之間人員搬送時間 (transfer_min)：
           當工序 k 完成並釋放機台後，若下一道工序 k+1 是「不同站點」且設定 transfer_min > 0，
           則觸發人員推車跨站搬送時間 `yield self.env.timeout(actual_transfer)`（期間不佔用機台）。
        """
        routing_df = self.routings[product_type]
        num_steps = len(routing_df)

        for step_idx in range(num_steps):
            step = routing_df.iloc[step_idx]
            st_name = step["station_name"]
            base_ct = float(step["ct"])
            resource = self.resources[st_name]

            # 1. 請求站點機台並進行加工
            with resource.request() as req:
                yield req
                effective_process_time = (base_ct / self.overall_availability) * random.uniform(0.92, 1.08)
                yield self.env.timeout(effective_process_time)

            # 2. 釋放當前站點機台後，檢查是否需要跨站人員搬送至下一站
            if self.transfer_min > 0 and step_idx < num_steps - 1:
                next_st_name = routing_df.iloc[step_idx + 1]["station_name"]
                # 若下一道工序位於不同設備站點，計入一次人員跨站搬送時間
                if next_st_name != st_name:
                    self.lots_in_transfer += 1
                    actual_transfer = self.transfer_min * random.uniform(0.85, 1.15)
                    yield self.env.timeout(actual_transfer)
                    self.lots_in_transfer -= 1

        # 若在正式模擬期間完工，計入完工統計
        if self.env.now >= self.warmup_duration:
            self.lots_completed[product_type] += 1

    def lot_generator(self):
        """
        源頭混線進料程序：
        每隔 X (1440 / daily_lots) 分鐘依 03SEC:01SG 設定比例投入 1 個 Lot。
        """
        seq = 1
        while True:
            prod_type = self.select_next_product(seq)
            self.lots_created[prod_type] += 1
            self.env.process(self.lot_process(seq, prod_type))
            seq += 1

            next_interval = self.interarrival_time * random.uniform(0.92, 1.08)
            yield self.env.timeout(next_interval)

    def monitor_process(self):
        """
        背景狀態紀錄程序 (Monitor)：
        等待預熱完成後，在正式模擬期間內，每隔 MONITOR_INTERVAL 分鐘紀錄一次全線所有站點的當下 WIP。
        """
        if self.warmup_duration > 0:
            yield self.env.timeout(self.warmup_duration)

        while True:
            sim_clock = round(self.env.now - self.warmup_duration, 2)
            snapshot = {"time": sim_clock, "in_transfer": self.lots_in_transfer}
            for st_name, res in self.resources.items():
                queue_len = len(res.queue)
                in_service = res.count
                snapshot[st_name] = queue_len + in_service
                snapshot[f"{st_name}__queue"] = queue_len
            self.history_records.append(snapshot)
            yield self.env.timeout(self.monitor_interval)

    def run(self) -> pd.DataFrame:
        """
        啟動模擬並回傳歷史紀錄表。
        """
        db_row = self.stations_df[self.stations_df["station_key"].str.contains("D/B-IC")].iloc[0]

        print("=" * 88)
        print("🚀 啟動【03SEC & 01SG 全線多站點混線動態模擬器 (含每日工時、機台 RUN% & 人員搬送)】")
        print("=" * 88)
        print(f"  • 每日投入總批數        : {self.daily_lots:.1f} Lots / 天")
        print(f"  • 產品混線比例設定      : 03SEC : 01SG = {self.ratio_03sec*100:.0f}% : {self.ratio_01sg*100:.0f}%")
        print(
            f"  • 每日設備實際生產時間  : 60m × {self.work_hours:.1f}hr × {self.run_pct:.0f}% (RUN%) "
            f"= {self.daily_equip_prod_mins:.0f} 分鐘 / 天 / 台"
        )
        print(
            f"  • 站間人員搬送時間設定  : 每次跨站搬送 {self.transfer_min:.1f} 分鐘 "
            f"{'(不計搬送時間)' if self.transfer_min == 0 else '(已啟用跨站人員推車搬送延遲)'}"
        )
        print(
            f"  • 02. D/B-IC D/B 站點   : 混線平均 C/T = {db_row['ct']:.2f}m | 5台並聯 -> "
            f"每日最大產能 {db_row['daily_max_lots']:.1f} Lots/天"
        )
        print(
            f"  • 全線最高負荷瓶頸預估  : {self.bottleneck_row['station_name']} "
            f"(每日最大產能僅 {self.bottleneck_row['daily_max_lots']:.1f} Lots/天)"
        )
        print("-" * 88)

        self.env.process(self.lot_generator())
        self.env.process(self.monitor_process())

        total_run_time = self.warmup_duration + self.sim_duration + 0.001
        self.env.run(until=total_run_time)

        history_df = pd.DataFrame(self.history_records)
        total_out = self.lots_completed["03SEC"] + self.lots_completed["01SG"]
        sim_days = self.sim_duration / 1440.0
        print(
            f"  • 模擬完成！{int(self.sim_duration)} 分鐘 ({sim_days:.1f} 天) 內全線完工總產出: {total_out} Lots "
            f"(03SEC: {self.lots_completed['03SEC']} Lots, 01SG: {self.lots_completed['01SG']} Lots)"
        )
        print(f"  • 歷史紀錄快照筆數      : {len(history_df)} 筆")
        print("=" * 88 + "\n")
        return history_df


# ==============================================================================
# 4. 自動識別瓶頸分析模組 (Automated Bottleneck Analysis - 支援自訂 Top N 顯示)
# ==============================================================================
def analyze_bottlenecks(
    stations_df: pd.DataFrame,
    history_df: pd.DataFrame,
    top_n: int = DEFAULT_TOP_N,
) -> pd.DataFrame:
    """
    計算全線每一個站點（包含 02. D/B-IC D/B）在模擬期間內的「平均 WIP (Average WIP)」，
    嚴格依平均 WIP 由高至低排序，並在終端機印出使用者指定數量的：
    === 全線 Top N 瓶頸站點分析 ===
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

    # 嚴格依照「平均 WIP (avg_wip)」由高至低排序（把 D/B 與全線所有站點一視同仁納入排名）
    ranked_df = analysis_df.sort_values(
        by=["avg_wip", "effective_ct"], ascending=[False, False]
    ).reset_index(drop=True)

    display_n = max(1, min(int(top_n), len(ranked_df)))
    print(f"=== 全線 Top {display_n} 瓶頸站點分析 ===")
    for idx in range(display_n):
        row = ranked_df.iloc[idx]
        print(
            f"{idx + 1}. {row['station_name']} - "
            f"平均 WIP: {row['avg_wip']:.2f} (排隊: {row['avg_queue']:.2f}, 峰值 WIP: {int(row['max_wip'])}) - "
            f"C/T: {row['ct']:.2f} min/lot (機台數: {int(row['capacity'])}, 每日上限: {row['daily_max_lots']:.1f} Lots/天)"
        )
    print("=" * 80 + "\n")

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
    if total_mins > 0 and rem_mins == 0:
        return f"Day {total_mins // 1440}, 24:00"
    return f"Day {day}, {hours:02d}:{mins:02d}"


def create_and_show_animation(
    stations_df: pd.DataFrame,
    history_df: pd.DataFrame,
    ranked_df: pd.DataFrame,
    daily_lots: float,
    ratio_03sec: float,
    ratio_01sg: float,
    work_hours: float,
    run_pct: float,
    transfer_min: float,
    interarrival_time: float,
    top_n: int = DEFAULT_TOP_N,
):
    """
    利用歷史紀錄表與 matplotlib.animation.FuncAnimation 製作高階主管動態簡報圖表。
    """
    setup_chinese_font()

    station_names = stations_df["station_name"].tolist()
    queue_cols = [f"{name}__queue" for name in station_names]
    num_stations = len(station_names)
    daily_equip_mins = 60.0 * work_hours * (run_pct / 100.0)

    fig, ax = plt.subplots(figsize=(15.2, 9.6), dpi=100)
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_facecolor("#FFFFFF")

    plt.subplots_adjust(left=0.27, right=0.95, top=0.88, bottom=0.08)

    y_pos = np.arange(num_stations)
    initial_wip = history_df.iloc[0][station_names].values.astype(float)

    COLOR_NORMAL = "#2B6CB0"     # 正常流動 (科技藍)
    COLOR_QUEUE = "#DD6B20"      # 排隊累積提醒 (琥珀橘)
    COLOR_WARNING = "#E53E3E"    # 嚴重塞車警示 (警戒紅)

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

    ax.set_xlim(0, X_AXIS_MAX)
    ax.set_xticks(np.arange(0, X_AXIS_MAX + 1, 5))
    ax.set_xlabel(
        "站點當下 WIP 批數 (含排隊等待 Queue + 機台加工中 Lots)",
        fontsize=11,
        fontweight="bold",
        color="#1A202C",
        labelpad=8,
    )

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
        f"【03SEC & 01SG 混線模擬】日投: {int(daily_lots)} Lots | 比例 {ratio_03sec*100:.0f}:{ratio_01sg*100:.0f} | "
        f"每日生產: 60×{work_hours:.0f}h×{run_pct:.0f}% = {daily_equip_mins:.0f} min/天 | 跨站搬送: {transfer_min:.1f}m",
        fontsize=13,
        fontweight="bold",
        color="#1A202C",
        y=0.96,
    )
    sub_title = ax.set_title(
        f"當前模擬時間: Day 1, 00:00  |  全線站點 WIP: 0 Lots  |  在途搬送中: 0 Lots",
        fontsize=11,
        fontweight="bold",
        color="#2C5282",
        pad=12,
    )

    # 右下角顯示 Top N 瓶頸分析摘要盒
    display_n = max(1, min(int(top_n), len(ranked_df)))
    box_display_n = min(display_n, 10)
    top_lines = [
        f"【全線 Top {display_n} 瓶頸站點 (日投 {int(daily_lots)} Lots | 每日實際生產 {daily_equip_mins:.0f}m | 搬送 {transfer_min:.1f}m)】"
    ]
    for idx in range(box_display_n):
        r = ranked_df.iloc[idx]
        top_lines.append(
            f"TOP {idx+1}: {r['station_name']} "
            f"(平均WIP: {r['avg_wip']:.2f} [排隊:{r['avg_queue']:.2f}] | 日產能上限: {r['daily_max_lots']:.1f} Lots)"
        )
    if display_n > box_display_n:
        top_lines.append(f"... (其餘 Top {box_display_n+1}~{display_n} 請見 Console 終端機輸出)")

    box_fontsize = 8.5 if box_display_n <= 5 else 7.5
    ax.text(
        0.985,
        0.025,
        "\n".join(top_lines),
        transform=ax.transAxes,
        fontsize=box_fontsize,
        verticalalignment="bottom",
        horizontalalignment="right",
        bbox=dict(boxstyle="round,pad=0.55", facecolor="#EDF2F7", edgecolor="#CBD5E0", alpha=0.94),
        color="#1A202C",
    )

    ax.legend(loc="upper right", frameon=True, facecolor="#FFFFFF", edgecolor="#CBD5E0", fontsize=9.5)

    def update(frame_idx: int):
        row = history_df.iloc[frame_idx]
        current_time = float(row["time"])
        in_trans = int(row.get("in_transfer", 0))
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
            elif wip_val >= 5 or q_val >= 2:
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
            f"站點總 WIP: {total_wip} Lots (排隊: {total_queue})  |  "
            f"人員搬送中: {in_trans} Lots  |  紅色超載站: {alert_count} 站"
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
# 6. 使用者互動參數解析與主程式進入點 (Main Entry Point)
# ==============================================================================
def parse_mix_ratio(mix_str: str) -> Tuple[float, float]:
    """
    解析使用者輸入的 03SEC:01SG 混線比例字串（例如 '7:3', '70:30', '5:5', '1:0'），
    正規化為總和 1.0 的兩個浮點數 (ratio_03sec, ratio_01sg)。
    """
    try:
        cleaned = mix_str.strip().replace("：", ":").replace("/", ":")
        if ":" in cleaned:
            parts = cleaned.split(":")
            r1 = max(0.0, float(parts[0]))
            r2 = max(0.0, float(parts[1]))
            total = r1 + r2
            if total > 0:
                return round(r1 / total, 4), round(r2 / total, 4)
        else:
            val = float(cleaned)
            if val > 1.0:
                val = val / 100.0
            val = max(0.0, min(1.0, val))
            return round(val, 4), round(1.0 - val, 4)
    except Exception:
        pass
    return 0.7, 0.3


def get_simulation_parameters() -> Tuple[float, float, float, float, float, float, int]:
    """
    開放讓使用者透過命令列參數或互動式輸入設定六大核心參數：
    1. 每日投入 Lot 數量 (--lots，預設 25，建議 10~30)
    2. 03SEC 與 01SG 混合比例 (--mix，預設 7:3)
    3. 每天生產時間 (小時) (--work-hours，預設 20 小時)
    4. 機台 RUN% 稼動率 (--run-pct，預設 80%，故每天實際生產時間 = 60 * 20 * 0.8 = 960 分鐘)
    5. 站與站之間人員搬送時間 (分鐘) (--transfer-min，預設 2.0 分鐘，0 代表不計)
    6. 顯示 Top N 瓶頸站點數量 (--top-n，預設 3)
    """
    parser = argparse.ArgumentParser(description="03SEC & 01SG 全線多站點動態混線模擬器")
    parser.add_argument(
        "--lots",
        type=float,
        default=None,
        help=f"設定每日投入總 Lot 數 (建議 10~30，預設: {DEFAULT_DAILY_LOTS})",
    )
    parser.add_argument(
        "--mix",
        type=str,
        default=None,
        help=f"設定 03SEC 與 01SG 混線比例 (例如 7:3，預設: {DEFAULT_MIX_RATIO_STR})",
    )
    parser.add_argument(
        "--work-hours",
        type=float,
        default=None,
        help=f"設定每天生產時間(小時) (1~24，預設: {DEFAULT_WORK_HOURS} 小時)",
    )
    parser.add_argument(
        "--run-pct",
        type=float,
        default=None,
        help=f"設定機台 RUN%% 正常運作時間百分比 (1~100，預設: {DEFAULT_RUN_PCT}%%)",
    )
    parser.add_argument(
        "--transfer-min",
        type=float,
        default=None,
        help=f"設定站與站之間人員搬送時間(分鐘/次) (預設: {DEFAULT_TRANSFER_MIN} 分鐘，設為 0 則不計)",
    )
    parser.add_argument(
        "--top-n",
        type=int,
        default=None,
        help=f"設定要顯示的 Top N 瓶頸站點數量 (預設: {DEFAULT_TOP_N})",
    )
    args, _ = parser.parse_known_args()

    daily_lots = args.lots
    mix_str = args.mix
    work_hours = args.work_hours
    run_pct = args.run_pct
    transfer_min = args.transfer_min
    top_n = args.top_n

    is_interactive = (os.environ.get("HEADLESS_TEST", "0") != "1") and sys.stdin.isatty()

    if is_interactive:
        print("\n" + "=" * 78)
        print("⚙️  【模擬參數設定】(直接按 Enter 可快速套用預設值)")
        print("=" * 78)

        if daily_lots is None:
            try:
                inp = input(f"  1. 請輸入每日投入 Lot 總數 (建議 10~30) [預設 {DEFAULT_DAILY_LOTS}]: ").strip()
                daily_lots = float(inp) if inp != "" else float(DEFAULT_DAILY_LOTS)
            except Exception:
                daily_lots = float(DEFAULT_DAILY_LOTS)

        if mix_str is None:
            try:
                inp = input(f"  2. 請輸入 03SEC : 01SG 混線比例 (如 7:3, 5:5) [預設 {DEFAULT_MIX_RATIO_STR}]: ").strip()
                mix_str = inp if inp != "" else DEFAULT_MIX_RATIO_STR
            except Exception:
                mix_str = DEFAULT_MIX_RATIO_STR

        if work_hours is None:
            try:
                inp = input(f"  3. 請輸入每天生產時間 (小時，如 20) [預設 {DEFAULT_WORK_HOURS} 小時]: ").strip()
                work_hours = float(inp) if inp != "" else float(DEFAULT_WORK_HOURS)
            except Exception:
                work_hours = float(DEFAULT_WORK_HOURS)

        if run_pct is None:
            try:
                inp = input(
                    f"  4. 請輸入機台 RUN% 正常運作百分比 (預設 {DEFAULT_RUN_PCT}% -> 每天實際生產 "
                    f"60*{work_hours:.0f}*0.8={60*work_hours*0.8:.0f}m): "
                ).strip()
                inp = inp.replace("%", "").strip()
                run_pct = float(inp) if inp != "" else float(DEFAULT_RUN_PCT)
            except Exception:
                run_pct = float(DEFAULT_RUN_PCT)

        if transfer_min is None:
            try:
                inp = input(
                    f"  5. 請輸入站與站之間人員搬送時間 (分鐘/次，輸入 0 代表不計) [預設 {DEFAULT_TRANSFER_MIN} 分鐘]: "
                ).strip()
                transfer_min = float(inp) if inp != "" else float(DEFAULT_TRANSFER_MIN)
            except Exception:
                transfer_min = float(DEFAULT_TRANSFER_MIN)

        if top_n is None:
            try:
                inp = input(f"  6. 請輸入要顯示的 Top N 瓶頸站點數量 (如 3, 5, 10) [預設 {DEFAULT_TOP_N}]: ").strip()
                top_n = int(inp) if inp != "" else int(DEFAULT_TOP_N)
            except Exception:
                top_n = int(DEFAULT_TOP_N)

    if daily_lots is None or daily_lots <= 0:
        daily_lots = float(DEFAULT_DAILY_LOTS)
    if mix_str is None or mix_str.strip() == "":
        mix_str = DEFAULT_MIX_RATIO_STR
    if work_hours is None or work_hours <= 0 or work_hours > 24:
        work_hours = float(DEFAULT_WORK_HOURS)
    if run_pct is None or run_pct <= 0 or run_pct > 100:
        run_pct = float(DEFAULT_RUN_PCT)
    if transfer_min is None or transfer_min < 0:
        transfer_min = float(DEFAULT_TRANSFER_MIN)
    if top_n is None or top_n <= 0:
        top_n = int(DEFAULT_TOP_N)

    ratio_03sec, ratio_01sg = parse_mix_ratio(mix_str)
    return daily_lots, ratio_03sec, ratio_01sg, work_hours, run_pct, transfer_min, top_n


def main():
    # Step 1: 取得使用者設定的模擬參數
    daily_lots, ratio_03sec, ratio_01sg, work_hours, run_pct, transfer_min, top_n = get_simulation_parameters()
    daily_equip_mins = 60.0 * work_hours * (run_pct / 100.0)

    # Step 2: 讀取 03SEC 與 01SG 雙產品 Excel 工藝流程並計算混線加權 C/T 與每日實際產能
    stations_df, routings = load_process_data(
        ratio_03sec=ratio_03sec,
        ratio_01sg=ratio_01sg,
        run_pct=run_pct,
        work_hours=work_hours,
    )

    print(
        f"\n【03SEC ({ratio_03sec*100:.0f}%) & 01SG ({ratio_01sg*100:.0f}%) 全線站點負荷總覽 "
        f"(每日設備實際生產時間 = 60×{work_hours:.0f}×{run_pct/100:.2f} = {daily_equip_mins:.0f} 分鐘/台)】"
    )
    preview_df = stations_df[
        [
            "station_id",
            "station_name",
            "capacity",
            "ct_03sec",
            "ct_01sg",
            "ct",
            "daily_max_lots",
        ]
    ].copy()
    preview_df.columns = [
        "序號",
        "站點名稱",
        "機台數",
        "03SEC累計C/T",
        "01SG累計C/T",
        "混線加權C/T",
        f"每日最大產能(Lots/天)",
    ]
    print(preview_df.to_string(index=False))
    print()

    # Step 3: 啟動 SimPy 雙產品混線模擬器
    simulator = ProductionLineSimulator(
        stations_df=stations_df,
        routings=routings,
        daily_lots=daily_lots,
        ratio_03sec=ratio_03sec,
        ratio_01sg=ratio_01sg,
        work_hours=work_hours,
        run_pct=run_pct,
        transfer_min=transfer_min,
        sim_duration=SIM_DURATION,
        warmup_duration=WARMUP_DURATION,
        monitor_interval=MONITOR_INTERVAL,
    )
    history_df = simulator.run()

    # Step 4: 自動識別並印出全線 Top N 瓶頸站點分析
    ranked_df = analyze_bottlenecks(stations_df, history_df, top_n=top_n)

    # Step 5: 啟動 Matplotlib 橫向長條圖動態視覺化
    if os.environ.get("HEADLESS_TEST", "0") != "1":
        create_and_show_animation(
            stations_df=stations_df,
            history_df=history_df,
            ranked_df=ranked_df,
            daily_lots=daily_lots,
            ratio_03sec=ratio_03sec,
            ratio_01sg=ratio_01sg,
            work_hours=work_hours,
            run_pct=run_pct,
            transfer_min=transfer_min,
            interarrival_time=simulator.interarrival_time,
            top_n=top_n,
        )


if __name__ == "__main__":
    main()
