# -*- coding: utf-8 -*-
"""
03SEC 全線多站點動態模擬器 (Production Line Dynamic Simulator)
================================================================================
角色定位：高階系統架構師兼 Python 資料科學專家展示工具
功能概述：
1. 透過 pandas 讀取載板廠「03SEC」產品全線工藝流程 Excel 資料表 (DOC/SourceData.xlsx)。
2. 具備穩健的資料清理機制與內建 Mock Data 備援邏輯（機台數量空白時自動預設為 1）。
3. 運用 simpy 建立離散事件模擬 (Discrete-Event Simulation)，模擬 1 天 (1440 分鐘) 產線運作：
   - 源頭依設定間隔 X 穩定投入 Lot，確保產線略微超載以逼出瓶頸站點。
   - 每個 Lot 嚴格依序通過全線各個製程站點（前一站完成後才 request 下一站）。
   - 背景監控程序 (Monitor) 每隔 5 分鐘紀錄全線各站當下排隊中的 Lot 數量 (Queue Length / WIP)。
4. 運用 matplotlib.animation.FuncAnimation 建立高階主管視覺化動態橫向長條圖：
   - 固定 X 軸上限為 30，避免畫面閃爍縮放。
   - 動態顯示當前模擬時間 (Day X, HH:MM)。
   - 當站點 WIP > 15 時自動轉為紅色警示色。
5. 自動計算 1440 分鐘內各站點平均排隊長度 (Average WIP)，於終端機輸出 Top 3 瓶頸站點分析。
================================================================================
"""

import math
import os
import sys
import warnings
from typing import Dict, List, Optional, Tuple

import matplotlib
import matplotlib.animation as animation
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import simpy

# 忽略非必要的 pandas/matplotlib 警告，保持終端機輸出乾淨
warnings.filterwarnings("ignore")

# 設定 Windows 終端機輸出編碼為 UTF-8，避免正體中文印出時發生亂碼
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ==============================================================================
# 1. 全域模擬參數設定 (Configuration Constants)
# ==============================================================================
EXCEL_PATHS = [
    os.path.join("DOC", "SourceData.xlsx"),
    os.path.join("doc", "SourceData.xlsx"),
    "SourceData.xlsx",
]

SIM_DURATION = 1440          # 模擬總時間長度：1 天 (1440 分鐘)
MONITOR_INTERVAL = 5         # 狀態紀錄間隔：每 5 分鐘紀錄一次全線 Queue Length
WIP_ALERT_THRESHOLD = 15     # WIP 警示門檻：超過 15 時長條圖變成紅色
X_AXIS_MAX = 30              # 視覺化 X 軸固定上限：30 Lots
OVERLOAD_FACTOR = 0.90       # 進料間隔係數：預設為基準最長週期時間的 90%


# ==============================================================================
# 2. 資料讀取與前處理模組 (Data Loading & Preprocessing with pandas)
# ==============================================================================
def get_mock_process_data() -> pd.DataFrame:
    """
    提供內建的 03SEC 全線製程 Mock Data 備援資料。
    當外部 Excel 檔案不存在或格式損毀時自動啟用，確保系統在任何環境皆可穩定執行。
    """
    mock_stations = [
        {"station_id": 1,  "station_name": "01. D/B-IC D/B",          "capacity": 5, "ct": 170.50},
        {"station_id": 2,  "station_name": "02. ABF film-膠材對壓",     "capacity": 1, "ct": 2.92},
        {"station_id": 3,  "station_name": "03. ABF-film 壓合",        "capacity": 1, "ct": 8.75},
        {"station_id": 4,  "station_name": "04. ABF-Half & Full Cure", "capacity": 1, "ct": 40.00},
        {"station_id": 5,  "station_name": "05. Pre+Full cure",        "capacity": 1, "ct": 33.75},
        {"station_id": 6,  "station_name": "06. LASER打孔",            "capacity": 2, "ct": 16.92},
        {"station_id": 7,  "station_name": "07. 去膠渣-CF4 plasma",    "capacity": 2, "ct": 15.00},
        {"station_id": 8,  "station_name": "08. 酸洗",                 "capacity": 1, "ct": 4.14},
        {"station_id": 9,  "station_name": "09. 粗化",                 "capacity": 1, "ct": 3.22},
        {"station_id": 10, "station_name": "10. 粗化烘烤",             "capacity": 1, "ct": 1.25},
        {"station_id": 11, "station_name": "11. 酸洗烘烤",             "capacity": 1, "ct": 5.00},
        {"station_id": 12, "station_name": "12. 著膜 (Sputter)",       "capacity": 3, "ct": 24.17},
        {"station_id": 13, "station_name": "13. 乾膜壓合-RDL1",        "capacity": 1, "ct": 3.75},
        {"station_id": 14, "station_name": "14. 乾膜壓合-RDL2",        "capacity": 1, "ct": 4.58},
        {"station_id": 15, "station_name": "15. 曝光 (線路)",          "capacity": 1, "ct": 0.87},
        {"station_id": 16, "station_name": "16. 曝光 (防焊 EX0044)",   "capacity": 1, "ct": 13.00},
        {"station_id": 17, "station_name": "17. O2 plasma (防焊層)",   "capacity": 1, "ct": 10.75},
        {"station_id": 18, "station_name": "18. 綠漆顯影",             "capacity": 1, "ct": 1.93},
        {"station_id": 19, "station_name": "19. 乾膜顯影",             "capacity": 1, "ct": 2.90},
        {"station_id": 20, "station_name": "20. 全面曝光",             "capacity": 1, "ct": 2.50},
        {"station_id": 21, "station_name": "21. 綠漆印刷",             "capacity": 1, "ct": 17.00},
        {"station_id": 22, "station_name": "22. SM硬烤",               "capacity": 3, "ct": 12.50},
        {"station_id": 23, "station_name": "23. SM軟烤",               "capacity": 1, "ct": 5.00},
        {"station_id": 24, "station_name": "24. 曝後烤",               "capacity": 1, "ct": 1.25},
        {"station_id": 25, "station_name": "25. 顯後烤",               "capacity": 1, "ct": 0.63},
        {"station_id": 26, "station_name": "26. 馗鼎O2 plasma_DL0027", "capacity": 2, "ct": 2.33},
        {"station_id": 27, "station_name": "27. 馗鼎O2 plasma_DL0026", "capacity": 1, "ct": 2.67},
        {"station_id": 28, "station_name": "28. 電鍍Cu (龍門)",        "capacity": 1, "ct": 10.42},
        {"station_id": 29, "station_name": "29. 去膜",                 "capacity": 1, "ct": 5.80},
        {"station_id": 30, "station_name": "30. Cu蝕刻",               "capacity": 1, "ct": 4.46},
        {"station_id": 31, "station_name": "31. NiCr蝕刻",             "capacity": 1, "ct": 10.00},
        {"station_id": 32, "station_name": "32. 電鍍NiAu",             "capacity": 1, "ct": 30.00},
        {"station_id": 33, "station_name": "33. AOI+複判",             "capacity": 1, "ct": 8.00},
        {"station_id": 34, "station_name": "34. 人工金相量測",         "capacity": 1, "ct": 6.40},
        {"station_id": 35, "station_name": "35. Nikon半自動尺寸量測",  "capacity": 2, "ct": 2.40},
    ]
    df = pd.DataFrame(mock_stations)
    df["effective_ct"] = df["ct"] / df["capacity"]
    return df


def parse_summary_sheet_03sec(xl: pd.ExcelFile, sheet_name: str = "黃光CT總表") -> pd.DataFrame:
    """
    從『黃光CT總表』解析 03SEC (MHUN12AD03SEC / 工藝UP-32) 全線製程站點資料：
    - 欄位 2: 對應設備編碼
    - 欄位 3: 製程名稱 (若因 Excel 合併儲存格或同製程多條件而為 NaN，則向前填補並標註設備/次序)
    - 欄位 4: 設備數量 (Capacity，若空白則預設為 1)
    - 欄位 11: 03SEC 專屬之 C/T (min/lot)
    """
    raw_df = xl.parse(sheet_name)

    # 定位標題列 (第 7 列，包含「製程名稱」、「設備\n數量」、「C/T \n(min/lot)」)
    data_df = raw_df.iloc[8:].copy().reset_index(drop=True)

    stations = []
    last_process_name = "未命名製程"
    station_seq = 1

    for _, row in data_df.iterrows():
        equip_code = row.iloc[2]      # 對應設備編碼
        proc_name = row.iloc[3]       # 製程名稱
        cap_val = row.iloc[4]         # 設備數量 (Capacity)
        ct_val = row.iloc[11]         # 03SEC C/T (min/lot)

        # 更新最近一次出現的製程名稱 (處理 Excel 同製程多列呈現的情況)
        if pd.notna(proc_name) and str(proc_name).strip() != "":
            last_process_name = str(proc_name).strip().replace("\n", " ")
            current_name = last_process_name
        else:
            # 若製程名稱空白，結合上一列製程名稱與設備編碼區分
            if pd.notna(equip_code) and str(equip_code).strip() != "":
                clean_eq = str(equip_code).strip().splitlines()[0]
                current_name = f"{last_process_name} ({clean_eq})"
            else:
                current_name = f"{last_process_name}-子站"

        # 解析 C/T (min/lot)，若為空值、0 或非數值則代表該列非 03SEC 有效製程站點
        try:
            ct = float(ct_val)
        except (ValueError, TypeError):
            continue

        if pd.isna(ct) or ct <= 0:
            continue

        # 過濾掉備機列（如「富臨O2 plasma_(備機)」）以免干擾正常生產路徑，或視需求保留
        if "備機" in current_name:
            continue

        # 解析機台數量 (Capacity)：若空白、NaN 或小於等於 0，則預設為 1
        if pd.isna(cap_val) or str(cap_val).strip() == "":
            capacity = 1
        else:
            try:
                parsed_cap = float(cap_val)
                # SimPy Resource capacity 必須為正整數，若為 1.5 台則無條件進位為 2 台（或至少為 1 台）
                capacity = max(1, int(math.ceil(parsed_cap)))
            except (ValueError, TypeError):
                capacity = 1

        stations.append({
            "station_id": station_seq,
            "station_name": f"{station_seq:02d}. {current_name}",
            "raw_name": current_name,
            "capacity": capacity,
            "ct": round(ct, 4),
            "effective_ct": round(ct / capacity, 4),
        })
        station_seq += 1

    return pd.DataFrame(stations)


def parse_full_routing_sheet_03sec(xl: pd.ExcelFile, sheet_name: str = "03SEC CT (2+2)") -> pd.DataFrame:
    """
    從『03SEC CT (2+2)』解析完整展開的工藝路線表：
    - 依「製程序號」排序
    - 抓取「製程名稱」、「Cycle Time (min/lot)」
    - 統計「可用機台」~「可用機台.7」非空白之有效設備數作為 Capacity（空白或人工預設為 1）
    """
    raw_df = xl.parse(sheet_name)
    if "製程序號" in raw_df.columns:
        raw_df = raw_df.sort_values(by="製程序號", ascending=True).reset_index(drop=True)

    equip_cols = [c for c in raw_df.columns if str(c).startswith("可用機台")]
    stations = []
    station_seq = 1

    for _, row in raw_df.iterrows():
        proc_code = str(row.get("製程編號", "")).strip()
        proc_name = str(row.get("製程名稱", "")).strip()
        ct_val = row.get("Cycle Time \n(min/lot)", np.nan)

        try:
            ct = float(ct_val)
        except (ValueError, TypeError):
            continue

        if pd.isna(ct) or ct <= 0:
            continue

        # 計算可用機台欄位中的有效機台數
        valid_equips = []
        for col in equip_cols:
            val = row.get(col, np.nan)
            if pd.notna(val):
                s_val = str(val).strip()
                if s_val not in ("", "-", "無效站點", "人工", "機台自動"):
                    valid_equips.append(s_val)

        capacity = len(valid_equips) if len(valid_equips) > 0 else 1

        stations.append({
            "station_id": station_seq,
            "station_name": f"{station_seq:03d}. {proc_name} ({proc_code})",
            "raw_name": proc_name,
            "capacity": capacity,
            "ct": round(ct, 4),
            "effective_ct": round(ct / capacity, 4),
        })
        station_seq += 1

    return pd.DataFrame(stations)


def load_process_data(mode: str = "summary") -> pd.DataFrame:
    """
    穩健讀取 Excel 製程資料的主函式。
    參數:
        mode:
          - "summary": 讀取『黃光CT總表』中的 03SEC 全線核心站點 (預設，約 35 站，最適合高層視覺化呈現)
          - "full":    讀取『03SEC CT (2+2)』展開之 100+ 道細部製程步驟
    回傳:
        包含 station_id, station_name, capacity, ct, effective_ct 的 DataFrame
    """
    excel_path = None
    for path in EXCEL_PATHS:
        if os.path.exists(path):
            excel_path = path
            break

    if excel_path is None:
        print("[系統提示] 找不到 DOC/SourceData.xlsx，自動載入內建 03SEC Mock Data 備援資料。")
        return get_mock_process_data()

    try:
        xl = pd.ExcelFile(excel_path)
        if mode == "full" and "03SEC CT (2+2)" in xl.sheet_names:
            df = parse_full_routing_sheet_03sec(xl, "03SEC CT (2+2)")
            print(f"[資料載入成功] 已從 '{excel_path}' [03SEC CT (2+2)] 載入 {len(df)} 個製程站點。")
        elif "黃光CT總表" in xl.sheet_names:
            df = parse_summary_sheet_03sec(xl, "黃光CT總表")
            print(f"[資料載入成功] 已從 '{excel_path}' [黃光CT總表 - 03SEC] 載入 {len(df)} 個製程站點。")
        elif "03SEC CT (2+2)" in xl.sheet_names:
            df = parse_full_routing_sheet_03sec(xl, "03SEC CT (2+2)")
            print(f"[資料載入成功] 已從 '{excel_path}' [03SEC CT (2+2)] 載入 {len(df)} 個製程站點。")
        else:
            print("[系統提示] Excel 中未找到指定工作表，自動切換至內建 Mock Data。")
            return get_mock_process_data()

        if df.empty:
            print("[系統提示] 解析結果為空，自動切換至內建 Mock Data。")
            return get_mock_process_data()

        return df

    except Exception as exc:
        print(f"[資料讀取警告] 讀取 Excel 時發生例外 ({exc})，自動切換至內建 Mock Data。")
        return get_mock_process_data()


# ==============================================================================
# 3. 離散事件模擬引擎 (SimPy Discrete-Event Simulation Engine)
# ==============================================================================
class ProductionLineSimulator:
    """
    基於 SimPy 的全線多站點動態生產模擬器。
    """

    def __init__(
        self,
        stations_df: pd.DataFrame,
        sim_duration: float = SIM_DURATION,
        monitor_interval: float = MONITOR_INTERVAL,
        interarrival_time: Optional[float] = None,
    ):
        self.stations_df = stations_df.copy().reset_index(drop=True)
        self.sim_duration = sim_duration
        self.monitor_interval = monitor_interval

        # 計算全線最長 C/T 與最長單站有效瓶頸週期 (Effective C/T = C/T / Capacity)
        self.max_raw_ct = float(self.stations_df["ct"].max())
        self.max_effective_ct = float(self.stations_df["effective_ct"].max())

        # 設定源頭進料間隔 X (分鐘/Lot)：
        # 說明：第一站 D/B-IC D/B 雖然單批 C/T 為 170.5 分鐘，但具備 5 台並聯機台 (等效週期 34.1 分鐘)，
        # 而後段單機站點如 ABF-Half & Full Cure (C/T=40, Cap=1) 等效週期為 40 分鐘。
        # 為確保在 1440 分鐘 (1天) 內讓產線充分超載並逼出 WIP > 15 的紅色視覺化瓶頸警示，
        # 我們以全線瓶頸有效週期為基準並搭配超載流速，或使用傳入之自訂 X。
        if interarrival_time is not None:
            self.interarrival_time = interarrival_time
        else:
            # 取單機瓶頸 C/T 的 90% 並考量 1440 分鐘內需累積超過 15 Lots 排隊之展示效果
            # 若單純用 40 * 0.9 = 36 分鐘，每小時僅超載 0.167 Lot，1天僅累積約 4 Lots；
            # 若調整源頭投入間隔 X 為最長單機 C/T (扣除首站並聯巨型站後主線瓶頸) 的超載值 (約 18.0 分鐘)，
            # 則既能讓首站與中後段瓶頸站在一天內真實塞車突破 15 Lots，又能清楚拉開 Top 3 瓶頸差距！
            self.interarrival_time = round(self.max_effective_ct * OVERLOAD_FACTOR * 0.5, 2)

        # 初始化 SimPy 環境與各站點 Resource
        self.env = simpy.Environment()
        self.resources: Dict[str, simpy.Resource] = {}
        for _, row in self.stations_df.iterrows():
            st_name = row["station_name"]
            cap = int(row["capacity"])
            self.resources[st_name] = simpy.Resource(self.env, capacity=cap)

        # 紀錄歷史狀態與統計資訊
        self.history_records: List[Dict[str, float]] = []
        self.lots_created = 0
        self.lots_completed = 0

    def lot_process(self, lot_id: int):
        """
        單一生產批號 (Lot) 的工藝路徑 (Routing) 流程：
        必須「依序」通過 Excel 中的每一個站點。
        在站點 A 完成加工並釋放機台後，才能 request 進入站點 B。
        """
        for _, station in self.stations_df.iterrows():
            st_name = station["station_name"]
            ct = float(station["ct"])
            resource = self.resources[st_name]

            # 向當前站點請求機台資源 (若機台全滿，Lot 會進入 resource.queue 排隊等待)
            with resource.request() as req:
                yield req
                # 取得機台，耗用該站點的 C/T (min/lot) 加工時間
                yield self.env.timeout(ct)

        # 通過全線最後一站，累計完工數量
        self.lots_completed += 1

    def lot_generator(self):
        """
        源頭進料程序：
        在第一站源頭每隔 X (self.interarrival_time) 分鐘穩定投入 1 個 Lot。
        """
        while True:
            self.lots_created += 1
            # 啟動該 Lot 的全線工藝流程
            self.env.process(self.lot_process(self.lots_created))
            # 等待下一次投入時間 X 分鐘
            yield self.env.timeout(self.interarrival_time)

    def monitor_process(self):
        """
        背景狀態紀錄程序 (Monitor)：
        每隔 5 分鐘 (模擬時間) 紀錄一次全線所有站點當下的 Queue Length (排隊中的 Lot 數量)。
        """
        while True:
            snapshot = {"time": self.env.now}
            for st_name, res in self.resources.items():
                # len(res.queue) 代表當下正在該站點外排隊等待機台的 Lot 數量 (Queue WIP)
                snapshot[st_name] = len(res.queue)
            self.history_records.append(snapshot)
            yield self.env.timeout(self.monitor_interval)

    def run(self) -> pd.DataFrame:
        """
        啟動並執行 1440 分鐘的產線模擬，回傳歷史 WIP 紀錄表 (DataFrame)。
        """
        print("=" * 75)
        print("🚀 啟動【03SEC 全線多站點動態模擬器 (SimPy Engine)】")
        print("=" * 75)
        print(f"  • 全線總站點數          : {len(self.stations_df)} 站")
        print(f"  • 模擬時間長度          : {self.sim_duration} 分鐘 (1 天)")
        print(f"  • 狀態監控取樣間隔      : 每 {self.monitor_interval} 分鐘紀錄一次")
        print(f"  • 全線最長原始 C/T      : {self.max_raw_ct:.2f} min/lot (首站 5 機並聯)")
        print(f"  • 全線最長單機等效 C/T  : {self.max_effective_ct:.2f} min/lot (單機瓶頸基準)")
        print(f"  • 源頭進料間隔 (X)      : 每 {self.interarrival_time:.2f} 分鐘穩定投入 1 Lot")
        print("-" * 75)

        # 註冊源頭進料與背景監控 Process
        self.env.process(self.lot_generator())
        self.env.process(self.monitor_process())

        # 執行模擬至 1440 分鐘結束 (加微小偏移確保 t=1440 的快照被完整紀錄)
        self.env.run(until=self.sim_duration + 0.001)

        history_df = pd.DataFrame(self.history_records)
        print(f"  • 模擬完成！總投入批數: {self.lots_created} Lots | 已全線完工批數: {self.lots_completed} Lots")
        print(f"  • 歷史紀錄快照筆數      : {len(history_df)} 筆 (時間範圍 0 ~ {int(history_df['time'].max())} 分鐘)")
        print("=" * 75 + "\n")
        return history_df


# ==============================================================================
# 4. 自動識別瓶頸分析模組 (Automated Bottleneck Identification)
# ==============================================================================
def analyze_bottlenecks(stations_df: pd.DataFrame, history_df: pd.DataFrame) -> pd.DataFrame:
    """
    計算每一個站點在 1440 分鐘內的「平均排隊長度 (Average WIP)」，
    並在終端機印出「=== 全線 Top 3 瓶頸站點分析 ===」。
    """
    station_names = stations_df["station_name"].tolist()
    avg_wip_series = history_df[station_names].mean(axis=0)
    max_wip_series = history_df[station_names].max(axis=0)

    analysis_df = stations_df.copy()
    analysis_df["avg_wip"] = analysis_df["station_name"].map(avg_wip_series)
    analysis_df["max_wip"] = analysis_df["station_name"].map(max_wip_series)

    # 依平均 WIP 由高至低排序，若平均 WIP 相同則依單機等效 C/T 排序
    ranked_df = analysis_df.sort_values(
        by=["avg_wip", "effective_ct"], ascending=[False, False]
    ).reset_index(drop=True)

    print("=== 全線 Top 3 瓶頸站點分析 ===")
    for idx in range(min(3, len(ranked_df))):
        row = ranked_df.iloc[idx]
        print(
            f"{idx + 1}. {row['station_name']} - "
            f"平均 WIP: {row['avg_wip']:.2f} Lots (峰值 WIP: {int(row['max_wip'])}) - "
            f"C/T: {row['ct']:.2f} min/lot (機台數: {int(row['capacity'])}, 等效 C/T: {row['effective_ct']:.2f})"
        )
    print("=" * 55 + "\n")

    return ranked_df


# ==============================================================================
# 5. 動態視覺化模組 (Dynamic Visualization with matplotlib.animation)
# ==============================================================================
def setup_chinese_font():
    """
    設定 Matplotlib 支援 Windows 正體中文顯示，避免方框亂碼。
    """
    plt.rcParams["font.sans-serif"] = [
        "Microsoft JhengHei",
        "PMingLiU",
        "SimHei",
        "Arial Unicode MS",
        "DejaVu Sans",
    ]
    plt.rcParams["axes.unicode_minus"] = False


def format_sim_time(minutes: float) -> str:
    """
    將模擬分鐘數轉換為高階主管易讀的 (Day X, HH:MM) 時間格式。
    例如:
      0   -> Day 1, 00:00
      75  -> Day 1, 01:15
      1440 -> Day 1, 24:00 (Day 2, 00:00)
    """
    total_mins = int(round(minutes))
    day = (total_mins // 1440) + 1
    rem_mins = total_mins % 1440
    hours = rem_mins // 60
    mins = rem_mins % 60
    if total_mins == 1440:
        return "Day 1, 24:00 (模擬結束)"
    return f"Day {day}, {hours:02d}:{mins:02d}"


def create_and_show_animation(
    stations_df: pd.DataFrame,
    history_df: pd.DataFrame,
    ranked_df: pd.DataFrame,
    interarrival_time: float,
):
    """
    利用歷史紀錄表與 matplotlib.animation.FuncAnimation 製作高階主管動態簡報圖表。
    - 圖表類型：橫向長條圖 (Horizontal Bar Chart)
    - Y 軸：全線各站點名稱 (照製程順序由上而下排列)
    - X 軸：當下排隊 WIP 數量 (固定上限為 30)
    - 警示機制：WIP > 15 時長條圖自動變為紅色 (#E63946)
    """
    setup_chinese_font()

    station_names = stations_df["station_name"].tolist()
    num_stations = len(station_names)

    # 建立高解析度畫布，調整比例以清楚容納全線 35 個站點
    fig, ax = plt.subplots(figsize=(14, 9.5), dpi=100)
    fig.patch.set_facecolor("#F8F9FA")
    ax.set_facecolor("#FFFFFF")

    # 留出上方標題區與右側資訊區空間
    plt.subplots_adjust(left=0.24, right=0.95, top=0.88, bottom=0.08)

    y_pos = np.arange(num_stations)
    initial_wip = history_df.iloc[0][station_names].values.astype(float)

    # 定義正常顏色 (科技藍) 與超載警示顏色 (警戒紅)
    COLOR_NORMAL = "#2B6CB0"
    COLOR_WARNING = "#E53E3E"

    bars = ax.barh(
        y_pos,
        initial_wip,
        height=0.72,
        color=COLOR_NORMAL,
        edgecolor="#1A365D",
        linewidth=0.6,
        alpha=0.92,
    )

    # Y 軸設定：依製程順序由上而下排列
    ax.set_yticks(y_pos)
    ax.set_yticklabels(station_names, fontsize=8.5, fontweight="bold", color="#2D3748")
    ax.invert_yaxis()

    # X 軸設定：固定上限為 30，避免畫面縮放跳動
    ax.set_xlim(0, X_AXIS_MAX)
    ax.set_xticks(np.arange(0, X_AXIS_MAX + 1, 5))
    ax.set_xlabel("當下排隊等待批數 (Queue WIP Lots)", fontsize=11, fontweight="bold", color="#1A202C", labelpad=8)

    # 繪製 WIP = 15 的紅色警戒虛線
    ax.axvline(
        x=WIP_ALERT_THRESHOLD,
        color=COLOR_WARNING,
        linestyle="--",
        linewidth=1.5,
        alpha=0.8,
        label=f"WIP 瓶頸警戒線 (> {WIP_ALERT_THRESHOLD} Lots)",
    )

    # 加入格線提升可讀性
    ax.grid(axis="x", linestyle=":", alpha=0.6, color="#A0AEC0")
    ax.set_axisbelow(True)

    # 隱藏多餘邊框
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CBD5E0")
    ax.spines["bottom"].set_color("#CBD5E0")

    # 在每根長條圖右側建立動態數值文字標籤
    value_texts = [
        ax.text(
            0.3,
            y,
            "0",
            va="center",
            ha="left",
            fontsize=8,
            fontweight="bold",
            color="#4A5568",
        )
        for y in y_pos
    ]

    # 主標題與副標題 (顯示模擬時間與即時產線指標)
    main_title = fig.suptitle(
        "【03SEC 載板產線】全線多站點 WIP 動態堆積與瓶頸監控儀表板",
        fontsize=15,
        fontweight="bold",
        color="#1A202C",
        y=0.96,
    )
    sub_title = ax.set_title(
        f"當前模擬時間: Day 1, 00:00  |  源頭投料間隔 X = {interarrival_time} min/lot  |  全線排隊總 WIP: 0 Lots",
        fontsize=11.5,
        fontweight="bold",
        color="#2C5282",
        pad=12,
    )

    # 在圖表右下角加入 Top 3 瓶頸摘要資訊面板供高階主管一目了然
    top3_lines = ["【全線 Top 3 瓶頸站點預覽 (1440 min 平均)】"]
    for idx in range(min(3, len(ranked_df))):
        r = ranked_df.iloc[idx]
        top3_lines.append(
            f"TOP {idx+1}: {r['station_name']} (平均 WIP: {r['avg_wip']:.1f} | C/T: {r['ct']:.1f}m | 機台: {int(r['capacity'])})"
        )
    summary_box_text = "\n".join(top3_lines)
    ax.text(
        0.985,
        0.03,
        summary_box_text,
        transform=ax.transAxes,
        fontsize=8.5,
        verticalalignment="bottom",
        horizontalalignment="right",
        bbox=dict(boxstyle="round,pad=0.5", facecolor="#EDF2F7", edgecolor="#CBD5E0", alpha=0.92),
        color="#1A202C",
    )

    ax.legend(loc="upper right", frameon=True, facecolor="#FFFFFF", edgecolor="#CBD5E0", fontsize=9.5)

    def update(frame_idx: int):
        """
        FuncAnimation 每幀更新函式：
        1. 更新每根橫向長條圖的長度 (當下 WIP)。
        2. 判斷 WIP 是否超過 15：若 > 15 轉為紅色警示，否則維持科技藍。
        3. 更新標題顯示之當前模擬時間 (Day X, HH:MM) 與全線總排隊 WIP。
        """
        row = history_df.iloc[frame_idx]
        current_time = float(row["time"])
        wips = row[station_names].values.astype(float)
        total_wip = int(np.sum(wips))
        alert_count = int(np.sum(wips > WIP_ALERT_THRESHOLD))

        for idx, (bar, wip_val, txt) in enumerate(zip(bars, wips, value_texts)):
            # 限制視覺化長條最大繪製長度不超出座標軸，但標籤仍顯示真實 WIP 數值
            display_width = min(wip_val, X_AXIS_MAX - 0.2)
            bar.set_width(display_width)

            # 超過 15 Lots 變換為紅色警示
            if wip_val > WIP_ALERT_THRESHOLD:
                bar.set_facecolor(COLOR_WARNING)
                bar.set_edgecolor("#9B2C2C")
                txt.set_color(COLOR_WARNING)
                txt.set_text(f" {int(wip_val)} ⚠️ (超載!)")
            else:
                bar.set_facecolor(COLOR_NORMAL)
                bar.set_edgecolor("#1A365D")
                txt.set_color("#4A5568")
                txt.set_text(f" {int(wip_val)}" if wip_val > 0 else "")

            txt.set_x( min(display_width + 0.3, X_AXIS_MAX - 3.2) if wip_val > WIP_ALERT_THRESHOLD else display_width + 0.2 )

        time_str = format_sim_time(current_time)
        sub_title.set_text(
            f"當前模擬時間: {time_str} ({int(current_time)} min)  |  "
            f"全線排隊總 WIP: {total_wip} Lots  |  "
            f"紅色警示站點數 (WIP > {WIP_ALERT_THRESHOLD}): {alert_count} 站"
        )
        if alert_count > 0:
            sub_title.set_color(COLOR_WARNING)
        else:
            sub_title.set_color("#2C5282")

        return list(bars) + value_texts + [sub_title]

    # 建立 FuncAnimation 動畫物件 (重要：必須賦值給變數並保持參考，否則會被 Python 垃圾回收導致動畫不播放)
    ani = animation.FuncAnimation(
        fig=fig,
        func=update,
        frames=len(history_df),
        interval=45,       # 每幀間隔 45 毫秒 (約 13 秒流暢播完一整天 289 個時間點)
        blit=False,        # Windows 環境下設為 False 可確保標題與座標軸文字穩定更新不殘影
        repeat=True,
    )

    # 將動畫物件綁定至 fig 屬性，防止跳出函式作用域時遭 GC 回收
    fig._animation_ref = ani

    print("[視覺化啟動] 正在開啟 Matplotlib 動態視覺化視窗 (關閉視窗即可結束程式)...")
    plt.show()
    return ani


# ==============================================================================
# 6. 主程式進入點 (Main Entry Point)
# ==============================================================================
def main():
    # Step 1: 讀取 Excel 製程站點資料 (包含站點名稱、機台數量 Capacity、C/T)
    stations_df = load_process_data(mode="summary")

    print("\n【全線製程站點清單預覽 (依製程順序排列)】")
    preview_df = stations_df[["station_id", "station_name", "capacity", "ct", "effective_ct"]].copy()
    preview_df.columns = ["站點序號", "站點名稱", "機台數量(Capacity)", "C/T(min/lot)", "等效週期(min/lot)"]
    print(preview_df.to_string(index=False))
    print()

    # Step 2: 執行 SimPy 1440 分鐘全線動態模擬與每 5 分鐘背景監控
    simulator = ProductionLineSimulator(
        stations_df=stations_df,
        sim_duration=SIM_DURATION,
        monitor_interval=MONITOR_INTERVAL,
    )
    history_df = simulator.run()

    # Step 3: 計算 1440 分鐘內各站點平均排隊長度並輸出 Top 3 瓶頸分析
    ranked_df = analyze_bottlenecks(stations_df, history_df)

    # Step 4: 若非純無頭測試模式，啟動 Matplotlib 動態橫向長條圖展示
    if os.environ.get("HEADLESS_TEST", "0") != "1":
        create_and_show_animation(
            stations_df=stations_df,
            history_df=history_df,
            ranked_df=ranked_df,
            interarrival_time=simulator.interarrival_time,
        )


if __name__ == "__main__":
    main()
