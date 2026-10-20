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

DEFAULT_SIM_MODE = "dynamic"     # 預設模擬模式："dynamic" (動態模擬：播放動畫+總結圖+分析圖) 或 "static" (靜態模擬：直接產生總結圖+分析圖)
DEFAULT_DAILY_LOTS = 22          # 預設每日投入 Lot 總數 (實際生產建議範圍：10 ~ 30 Lots/天)
DEFAULT_MIX_RATIO_STR = "7:3"    # 預設 03SEC 與 01SG 混線比例 (7:3 代表 70% 03SEC + 30% 01SG)
DEFAULT_WORK_HOURS = 20.0        # 預設每天生產時間 (小時)，預設 20 小時 (即 60 * 20 = 1200 分鐘/天)
DEFAULT_RUN_PCT = 80.0           # 預設機台 RUN% (預設 80%，故每天每台設備實際生產時間為 60 * 20 * 0.8 = 960 分鐘)
DEFAULT_TRANSFER_MIN = 8.0       # 預設站與站之間人員搬送時間 (分鐘/次，設為 0 代表不計搬送時間)
DEFAULT_TOP_N = 3                # 預設顯示 Top N 瓶頸站點數量 (可自由設定要顯示 Top 多少)

SIM_DURATION = 10080              # 正式模擬時間長度 (可依天數或分鐘設定)
WARMUP_DURATION = 10080          # 產線預熱時間：確保 Day 1, 00:00 啟動時全線皆有正常期初 WIP
MONITOR_INTERVAL = 15            # 狀態紀錄間隔 (分鐘)
WIP_ALERT_THRESHOLD = 20         # WIP 紅色警示門檻：超過此數值時長條圖變成紅色，且分析圖在此處畫紅色虛線
X_AXIS_MAX = 30                  # 視覺化 X 軸固定上限：30 Lots
RANDOM_SEED = 42                 # 固定隨機種子，確保簡報展示結果穩定可重現
OUTPUT_CHART_DIR = "output_charts"  # 自動儲存總結圖與 Top N 分析圖的資料夾路徑


# ==============================================================================
# 2. 載板廠設備群組對照與雙產品資料讀取模組 (pandas Data Processing)
# ==============================================================================
# 站點標準顯示名稱對照字典（機台數量 Capacity 不在此寫死，改由 Excel 可用機台欄位動態判定）
EQUIPMENT_NAME_MAP = {
    "UPPA010": "L/F貼耐熱膠帶 (TE0109/人工)",
    "TE0109":  "L/F貼耐熱膠帶 (TE0109/人工)",
    "SM0053":  "D/B-IC D/B (Die Bond)",
    "TS0145":  "Nikon半自動尺寸/首件量測 (TS0145)",
    "CL0200":  "ABF/SM 表面粗化 (CL0200)",
    "OV0036":  "粗化/酸洗烘烤 (OV0036)",
    "LF0003":  "ABF film 壓膜/壓合 (LF0003)",
    "OV0750":  "Oven 烘箱群-Cure/軟硬烤 (OV0750~1238)",
    "UPRM001": "ABF-film 撕耐熱膠帶 (人工)",
    "DR0026":  "LASER 雷射打孔 (DR0026/09)",
    "DL0017":  "CF4 Plasma 去膠渣 (DL0017)",
    "CL0201":  "水平酸洗線 (CL0201)",
    "DL0027":  "馗鼎 O2 Plasma (DL0027/26)",
    "ET0023":  "Cu 銅蝕刻線 (ET0023)",
    "VC1548":  "AOI 外觀自動光學檢查 (VC1548)",
    "UPQC":    "QC 品質檢驗站 (QC-1/2/3)",
    "OV1239":  "RDL 酸洗烘烤 (OV1239/40)",
    "SP0043":  "Sputter 濺鍍著膜 (SP0043/04)",
    "LF0014":  "RDL 乾膜壓合 (LF0014)",
    "曝光新機": "RDL 線路曝光 (曝光新機)",
    "OV0038":  "RDL 曝後烤/顯後烤 (OV0038)",
    "DE0025":  "RDL 乾膜顯影 (DE0025)",
    "PL0084":  "龍門雙面電鍍 Cu (PL0084)",
    "ST0012":  "RDL 水平去膜 (ST0012)",
    "ET0026":  "NiCr 蝕刻線 (ET0026)",
    "PR0254":  "Via/SM 樹脂綠漆印刷 (PR0254)",
    "EX0044":  "Via/SM 防焊對位曝光 (EX0044)",
    "DE0005":  "Via/SM 綠漆顯影 (DE0005)",
    "DL0010":  "SM 防焊層 O2 Plasma (DL0010)",
    "EX0028":  "Via/SM 全面曝光 (EX0028)",
    "TS0014":  "SM/Cu 人工金相尺寸量測 (TS0014)",
    "PL0088":  "雙面電鍍 NiAu (PL0088)",
    "XR0008":  "X-Ray NiAu 厚度量測 (XR0008)",
    "UPSC069": "Q3 扣帳數量比對 (人工)",
}

# 人工作業站點標準配置人數 (非機台)
MANUAL_STATION_CAPACITY = {
    "UPQC": 3,     # QC-1, QC-2, QC-3 三班/三站
    "UPPA010": 2,  # 貼耐熱膠帶作業人數
    "UPRM001": 2,  # 撕耐熱膠帶作業人數
    "UPSC069": 2,  # 扣帳比對作業人數
    "TE0109": 2,   # 貼膠機作業人數
}


def resolve_station_group(proc_code: str, proc_name: str, equip_list: List[str]) -> Tuple[str, int, List[str]]:
    """
    將工藝流程中的工序對應至真實的物理站點與可用設備清單，並依 Excel 可用機台欄位動態判定機台數量 (Capacity)。
    """
    primary_eq = equip_list[0] if len(equip_list) > 0 else ""

    # 1. 決定標準站點名稱
    if proc_code.startswith("UPQC"):
        st_name = EQUIPMENT_NAME_MAP["UPQC"]
    elif proc_code in ("UPPA010", "UPPA008"):
        st_name = EQUIPMENT_NAME_MAP["UPPA010"]
    elif proc_code == "UPRM001":
        st_name = EQUIPMENT_NAME_MAP["UPRM001"]
    elif proc_code == "UPSC069":
        st_name = EQUIPMENT_NAME_MAP["UPSC069"]
    elif proc_code == "UPEX032":
        st_name = EQUIPMENT_NAME_MAP["EX0044"]
    elif primary_eq in EQUIPMENT_NAME_MAP:
        st_name = EQUIPMENT_NAME_MAP[primary_eq]
    else:
        st_name = f"{proc_name} ({primary_eq if primary_eq else '預設單機'})"

    # 2. 統計真實有效機台清單 (只要出現在可用機台內就算可生產設備)
    valid_machines = [eq for eq in equip_list if eq and eq not in ("-", "機台自動", "無效站點")]
    non_manual_machines = [eq for eq in valid_machines if eq != "人工"]

    # 3. 動態判定機台數量 (Capacity)
    if len(non_manual_machines) > 0:
        # 有填寫具體設備 (例如 PR0254, 舊機改機, OV0750~1238, SM0053~新機等)
        cap = len(non_manual_machines)
    elif proc_code.startswith("UPQC"):
        cap = MANUAL_STATION_CAPACITY["UPQC"]
    elif proc_code in MANUAL_STATION_CAPACITY:
        cap = MANUAL_STATION_CAPACITY[proc_code]
    elif primary_eq in MANUAL_STATION_CAPACITY:
        cap = MANUAL_STATION_CAPACITY[primary_eq]
    else:
        cap = max(1, len(valid_machines))

    return st_name, max(1, cap), non_manual_machines


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

        station_key, capacity, valid_eqs = resolve_station_group(proc_code, proc_name, equip_list)

        routing_steps.append({
            "product": product_code,
            "step_seq": int(step_seq) if pd.notna(step_seq) else 0,
            "proc_code": proc_code,
            "proc_name": proc_name,
            "station_key": station_key,
            "capacity": max(1, int(capacity)),
            "valid_eqs": valid_eqs,
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
    - 依據 Excel 中該站點出現之相異設備集合，動態聚合計算真實設備數量 (Capacity)
    - 每天設備實際可生產時間 = 60 * work_hours * (run_pct / 100.0) (預設 60 * 20 * 0.8 = 960 分鐘/天/台)
    - 折算至 24 小時日曆模擬時間軸的綜合可用率因子 = (work_hours / 24.0) * (run_pct / 100.0)
    - 計算該站點「每日最大產能批數 (daily_cap_lots = (daily_equip_prod_mins * capacity) / weighted_ct)」
    """
    all_routings = pd.concat([routing_03sec, routing_01sg], ignore_index=True)
    station_order = []
    for st_key in all_routings["station_key"]:
        if st_key not in station_order:
            station_order.append(st_key)

    # 聚合每個站點在所有工序中出現過的所有相異有效設備，動態計算站點總容量 (Capacity)
    station_capacity_map: Dict[str, int] = {}
    for st_key, grp in all_routings.groupby("station_key"):
        all_unique_eqs = set()
        for eqs in grp["valid_eqs"]:
            all_unique_eqs.update(eqs)
        if len(all_unique_eqs) > 0:
            station_capacity_map[st_key] = max(1, len(all_unique_eqs))
        else:
            station_capacity_map[st_key] = max(1, int(grp["capacity"].max()))

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

        cap = station_capacity_map.get(st_key, 1)

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
        等待預熱完成後，在正式模擬期間內，每隔 MONITOR_INTERVAL 分鐘紀錄一次全線所有站點的當下 WIP
        以及「產線總 WIP = 站點 WIP (製程中 + 等待中) + 人員搬送 Lot」。
        """
        if self.warmup_duration > 0:
            yield self.env.timeout(self.warmup_duration)

        while True:
            sim_clock = round(self.env.now - self.warmup_duration, 2)
            in_trans = int(self.lots_in_transfer)
            snapshot = {"time": sim_clock, "in_transfer": in_trans}
            station_wip_sum = 0
            in_service_sum = 0
            queue_sum = 0

            for st_name, res in self.resources.items():
                queue_len = len(res.queue)
                in_service = res.count
                st_wip = queue_len + in_service
                snapshot[st_name] = st_wip
                snapshot[f"{st_name}__queue"] = queue_len

                station_wip_sum += st_wip
                in_service_sum += in_service
                queue_sum += queue_len

            snapshot["total_station_wip"] = station_wip_sum
            snapshot["total_in_service"] = in_service_sum
            snapshot["total_queue"] = queue_sum
            snapshot["total_line_wip"] = station_wip_sum + in_trans

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

        avg_station_wip = float(history_df["total_station_wip"].mean())
        avg_in_service = float(history_df["total_in_service"].mean())
        avg_queue = float(history_df["total_queue"].mean())
        avg_transfer = float(history_df["in_transfer"].mean())
        avg_line_wip = float(history_df["total_line_wip"].mean())

        end_row = history_df.iloc[-1]
        end_station_wip = int(end_row["total_station_wip"])
        end_in_service = int(end_row["total_in_service"])
        end_queue = int(end_row["total_queue"])
        end_transfer = int(end_row["in_transfer"])
        end_line_wip = int(end_row["total_line_wip"])

        print(
            f"  • 模擬完成！{int(self.sim_duration)} 分鐘 ({sim_days:.1f} 天) 內全線完工總產出: {total_out} Lots "
            f"(03SEC: {self.lots_completed['03SEC']} Lots, 01SG: {self.lots_completed['01SG']} Lots)"
        )
        print(
            f"  • 平均【產線總 WIP】    : {avg_line_wip:.2f} Lots "
            f"= 站點 WIP {avg_station_wip:.2f} Lots (製程中: {avg_in_service:.2f} + 等待中: {avg_queue:.2f}) "
            f"+ 人員搬送 {avg_transfer:.2f} Lots"
        )
        print(
            f"  • 期末【產線總 WIP】    : {end_line_wip} Lots "
            f"= 站點 WIP {end_station_wip} Lots (製程中: {end_in_service} + 等待中: {end_queue}) "
            f"+ 人員搬送 {end_transfer} Lots"
        )
        print(f"  • 歷史紀錄快照筆數      : {len(history_df)} 筆")
        print("=" * 88 + "\n")
        return history_df


# ==============================================================================
# 4. 自動識別瓶頸分析與最高峰統計模組 (Automated Bottleneck & Peak Analysis)
# ==============================================================================
def analyze_bottlenecks(
    stations_df: pd.DataFrame,
    history_df: pd.DataFrame,
    top_n: int = DEFAULT_TOP_N,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    計算全線每一個站點在模擬期間內的：
    1. 平均 WIP 排名 (ranked_avg_df)
    2. WIP 最高峰排名 (ranked_peak_df)
    以及產線總 WIP、站點 WIP、人員搬送 Lot 的【最高 / 最低 / 平均】統計。
    """
    station_names = stations_df["station_name"].tolist()
    queue_cols = [f"{name}__queue" for name in station_names]

    avg_wip_series = history_df[station_names].mean(axis=0)
    max_wip_series = history_df[station_names].max(axis=0)
    min_wip_series = history_df[station_names].min(axis=0)
    avg_queue_series = history_df[queue_cols].mean(axis=0)
    avg_queue_series.index = station_names
    max_queue_series = history_df[queue_cols].max(axis=0)
    max_queue_series.index = station_names

    # 找出各站點發生 WIP 最高峰的首次模擬時間點
    peak_time_map = {}
    for st in station_names:
        peak_idx = history_df[st].idxmax()
        peak_time_map[st] = float(history_df.loc[peak_idx, "time"])

    analysis_df = stations_df.copy()
    analysis_df["avg_wip"] = analysis_df["station_name"].map(avg_wip_series)
    analysis_df["avg_queue"] = analysis_df["station_name"].map(avg_queue_series)
    analysis_df["max_wip"] = analysis_df["station_name"].map(max_wip_series)
    analysis_df["min_wip"] = analysis_df["station_name"].map(min_wip_series)
    analysis_df["max_queue"] = analysis_df["station_name"].map(max_queue_series)
    analysis_df["peak_time"] = analysis_df["station_name"].map(peak_time_map)

    # 1. 依「平均 WIP (avg_wip)」由高至低排序
    ranked_avg_df = analysis_df.sort_values(
        by=["avg_wip", "max_wip", "effective_ct"], ascending=[False, False, False]
    ).reset_index(drop=True)

    # 2. 依「WIP 最高峰 (max_wip)」由高至低排序
    ranked_peak_df = analysis_df.sort_values(
        by=["max_wip", "avg_wip", "effective_ct"], ascending=[False, False, False]
    ).reset_index(drop=True)

    # 印出全線 WIP 與人員搬送 Lot 的【最高 / 最低 / 平均】總結統計
    line_max = int(history_df["total_line_wip"].max())
    line_min = int(history_df["total_line_wip"].min())
    line_avg = float(history_df["total_line_wip"].mean())

    st_max = int(history_df["total_station_wip"].max())
    st_min = int(history_df["total_station_wip"].min())
    st_avg = float(history_df["total_station_wip"].mean())

    tr_max = int(history_df["in_transfer"].max())
    tr_min = int(history_df["in_transfer"].min())
    tr_avg = float(history_df["in_transfer"].mean())

    print("=== 全線產線 WIP 與人員搬送 Lot 統計總結 (最高 / 最低 / 平均) ===")
    print(f"  • 產線總 WIP (站點+搬送) : 最高 {line_max:3d} Lots  |  最低 {line_min:3d} Lots  |  平均 {line_avg:6.2f} Lots")
    print(f"  • 站點總 WIP (製程+等待) : 最高 {st_max:3d} Lots  |  最低 {st_min:3d} Lots  |  平均 {st_avg:6.2f} Lots")
    print(f"  • 人員搬送 Lot 數        : 最高 {tr_max:3d} Lots  |  最低 {tr_min:3d} Lots  |  平均 {tr_avg:6.2f} Lots")
    print("=" * 80 + "\n")

    display_n = max(1, min(int(top_n), len(ranked_avg_df)))
    print(f"=== 全線【平均 WIP】Top {display_n} 瓶頸站點名單 ===")
    for idx in range(display_n):
        row = ranked_avg_df.iloc[idx]
        print(
            f"{idx + 1}. {row['station_name']} - "
            f"平均 WIP: {row['avg_wip']:.2f} (排隊: {row['avg_queue']:.2f}, 最高峰 WIP: {int(row['max_wip'])}) - "
            f"C/T: {row['ct']:.2f} min/lot (機台數: {int(row['capacity'])}, 每日上限: {row['daily_max_lots']:.1f} Lots/天)"
        )
    print("=" * 80 + "\n")

    print(f"=== 全線【WIP 最高峰 (Peak WIP)】Top {display_n} 站點名單 ===")
    for idx in range(display_n):
        row = ranked_peak_df.iloc[idx]
        peak_t_str = format_sim_time(row["peak_time"])
        print(
            f"{idx + 1}. {row['station_name']} - "
            f"最高峰 WIP: {int(row['max_wip'])} Lots (發生於 {peak_t_str}, 最大排隊: {int(row['max_queue'])}, 平均 WIP: {row['avg_wip']:.2f}) - "
            f"每日上限: {row['daily_max_lots']:.1f} Lots/天"
        )
    print("=" * 80 + "\n")

    return ranked_avg_df, ranked_peak_df


# ==============================================================================
# 5. 視覺化模組 (動態長條圖動畫 + 總結儀表板圖 + Top N 站點折線分析圖)
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


def build_time_axis_ticks(times: np.ndarray, max_ticks: int = 8) -> Tuple[List[float], List[str]]:
    """為時間折線圖建立清晰的 X 軸時間刻度與標籤，支援限制最大刻度數避免文字重疊。"""
    max_t = float(np.max(times)) if len(times) > 0 else 1440.0
    if max_t <= 1440:
        step = 240  # 每 4 小時一個刻度
    elif max_t <= 4320:
        step = 720  # 每 12 小時一個刻度
    else:
        step = 1440  # 每 24 小時一個刻度

    tick_vals = list(np.arange(0, max_t + 1, step))
    while len(tick_vals) > max_ticks and len(tick_vals) > 2:
        step *= 2
        tick_vals = list(np.arange(0, max_t + 1, step))
        if tick_vals[-1] < max_t:
            tick_vals.append(max_t)

    tick_labels = [f"{format_sim_time(t)}\n({int(t)}m)" for t in tick_vals]
    return tick_vals, tick_labels


def generate_summary_and_station_charts(
    history_df: pd.DataFrame,
    ranked_avg_df: pd.DataFrame,
    ranked_peak_df: pd.DataFrame,
    daily_lots: float,
    ratio_03sec: float,
    ratio_01sg: float,
    work_hours: float,
    run_pct: float,
    transfer_min: float,
    top_n: int = DEFAULT_TOP_N,
    show_windows: bool = True,
) -> List[str]:
    """
    產生並儲存：
    1. 模擬總結圖 (Summary Dashboard)：
       - 產線總 WIP、站點 WIP、人員搬送 Lot 的【最高 / 最低 / 平均】統計表與全期趨勢圖
       - 【WIP 最高峰 Top N 名單】與【平均 WIP Top N 名單】橫向對照圖
    2. Top N 站點個別折線分析圖 (一個站點一張圖)：
       - Y 軸：WIP 數 (含該站總 WIP 實線與等待中 Queue 虛線)
       - X 軸：模擬時間
       - 紅色虛線：在警戒 WIP 數 (WIP_ALERT_THRESHOLD) 處畫出醒目紅色虛線
    """
    setup_chinese_font()
    os.makedirs(OUTPUT_CHART_DIR, exist_ok=True)
    saved_files: List[str] = []

    times = history_df["time"].values.astype(float)
    tick_vals_full, tick_labels_full = build_time_axis_ticks(times, max_ticks=8)
    tick_vals_sub, tick_labels_sub = build_time_axis_ticks(times, max_ticks=5)
    daily_equip_mins = 60.0 * work_hours * (run_pct / 100.0)
    display_n = max(1, min(int(top_n), len(ranked_peak_df)))

    # --------------------------------------------------------------------------
    # 圖表 A：產線模擬總結圖 (Summary Dashboard)
    # --------------------------------------------------------------------------
    # 根據 display_n 動態調整畫布高度，並將上下兩列分別設定適合的 wspace，徹底消除左右重疊
    bottom_height_ratio = max(1.15, 0.25 * display_n)
    fig_height = max(10.0, 5.2 + 0.95 * display_n)
    fig_sum = plt.figure(figsize=(18.5, fig_height), dpi=100)
    fig_sum.patch.set_facecolor("#F8F9FA")

    outer_gs = fig_sum.add_gridspec(
        2, 1,
        height_ratios=[1.05, bottom_height_ratio],
        hspace=0.38,
        left=0.18,
        right=0.97,
        top=0.89,
        bottom=0.08,
    )
    gs_top = outer_gs[0].subgridspec(1, 2, width_ratios=[1.05, 1.05], wspace=0.22)
    # 下方兩個橫向柱狀圖的 Y 軸有長站點名稱，因此將 wspace 設為 0.62 預留充足空間，絕不重疊
    gs_bot = outer_gs[1].subgridspec(1, 2, width_ratios=[1.0, 1.0], wspace=0.62)

    fig_sum.suptitle(
        f"【03SEC & 01SG 產線模擬總結報告】日投: {int(daily_lots)} Lots | 混線 {ratio_03sec*100:.0f}:{ratio_01sg*100:.0f} | "
        f"每日生產: 60×{work_hours:.0f}h×{run_pct:.0f}% = {daily_equip_mins:.0f}m | 跨站搬送: {transfer_min:.1f}m",
        fontsize=14.5,
        fontweight="bold",
        color="#1A202C",
        y=0.96,
    )

    # (A-1) 左上：全線產線總 WIP、站點 WIP 與人員搬送 Lot 時間趨勢圖
    ax_trend = fig_sum.add_subplot(gs_top[0, 0])
    ax_trend.set_facecolor("#FFFFFF")
    ax_trend.plot(times, history_df["total_line_wip"], color="#1A365D", linewidth=2.2, label="產線總 WIP (站點+搬送)")
    ax_trend.plot(times, history_df["total_station_wip"], color="#2B6CB0", linewidth=1.8, linestyle="--", label="站點總 WIP (製程+等待)")
    ax_trend.plot(times, history_df["in_transfer"], color="#D69E2E", linewidth=1.6, label="人員搬送中 Lot 數")
    ax_trend.fill_between(times, history_df["total_station_wip"], history_df["total_line_wip"], color="#FEFCBF", alpha=0.5)
    ax_trend.set_title("全線產線總 WIP vs 站點 WIP vs 人員搬送 Lot 趨勢", fontsize=11.5, fontweight="bold", color="#2C5282", pad=10)
    ax_trend.set_ylabel("Lot 數量", fontsize=10, fontweight="bold")
    ax_trend.set_xticks(tick_vals_sub)
    ax_trend.set_xticklabels(tick_labels_sub, fontsize=8.5)
    ax_trend.grid(True, linestyle=":", alpha=0.6)
    ax_trend.legend(loc="upper left", fontsize=8.5, frameon=True, facecolor="#FFFFFF", framealpha=0.9)

    # (A-2) 右上：產線 WIP 與人員搬送 Lot 的【最高 / 最低 / 平均】統計表格
    ax_table = fig_sum.add_subplot(gs_top[0, 1])
    ax_table.axis("off")

    line_max, line_min, line_avg = int(history_df["total_line_wip"].max()), int(history_df["total_line_wip"].min()), float(history_df["total_line_wip"].mean())
    st_max, st_min, st_avg = int(history_df["total_station_wip"].max()), int(history_df["total_station_wip"].min()), float(history_df["total_station_wip"].mean())
    srv_max, srv_min, srv_avg = int(history_df["total_in_service"].max()), int(history_df["total_in_service"].min()), float(history_df["total_in_service"].mean())
    q_max, q_min, q_avg = int(history_df["total_queue"].max()), int(history_df["total_queue"].min()), float(history_df["total_queue"].mean())
    tr_max, tr_min, tr_avg = int(history_df["in_transfer"].max()), int(history_df["in_transfer"].min()), float(history_df["in_transfer"].mean())

    ax_table.set_title("產線 WIP 與人員搬送 Lot 統計總表 (最高 / 最低 / 平均)", fontsize=11.5, fontweight="bold", color="#2C5282", pad=10)

    table_data = [
        ["產線總 WIP (站點+搬送)", f"{line_max} Lots", f"{line_min} Lots", f"{line_avg:.2f} Lots"],
        ["站點總 WIP (製程+等待)", f"{st_max} Lots", f"{st_min} Lots", f"{st_avg:.2f} Lots"],
        ["  • 機台製程中 WIP", f"{srv_max} Lots", f"{srv_min} Lots", f"{srv_avg:.2f} Lots"],
        ["  • 站點等待中 Queue", f"{q_max} Lots", f"{q_min} Lots", f"{q_avg:.2f} Lots"],
        ["人員搬送中 Lot 數", f"{tr_max} Lots", f"{tr_min} Lots", f"{tr_avg:.2f} Lots"],
    ]
    col_labels = ["統計項目", "最高 (Max)", "最低 (Min)", "平均 (Avg)"]

    # 明確配置第一欄寬度為 43%，後三欄各 19%，確保長中文字串絕不超出表格格線
    tbl = ax_table.table(
        cellText=table_data,
        colLabels=col_labels,
        colWidths=[0.43, 0.19, 0.19, 0.19],
        loc="center",
        cellLoc="center",
        bbox=[0.0, 0.08, 1.0, 0.84],
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(9.5)
    for (r_idx, c_idx), cell in tbl.get_celld().items():
        cell.set_edgecolor("#CBD5E0")
        cell.set_linewidth(1.0)
        if r_idx == 0:
            cell.set_facecolor("#2B6CB0")
            cell.get_text().set_color("#FFFFFF")
            cell.get_text().set_fontweight("bold")
        elif r_idx == 1:
            cell.set_facecolor("#EBF8FF")
            cell.get_text().set_fontweight("bold")
        elif r_idx == 5:
            cell.set_facecolor("#FEFCBF")
            cell.get_text().set_fontweight("bold")
        else:
            cell.set_facecolor("#FFFFFF")

    # (A-3) 左下：【WIP 最高峰 (Peak WIP)】Top N 站點排行圖
    ax_peak = fig_sum.add_subplot(gs_bot[0, 0])
    ax_peak.set_facecolor("#FFFFFF")
    sub_peak = ranked_peak_df.iloc[:display_n].copy()
    y_p = np.arange(len(sub_peak))
    peak_vals = sub_peak["max_wip"].values.astype(float)
    peak_colors = ["#E53E3E" if v >= WIP_ALERT_THRESHOLD else "#DD6B20" for v in peak_vals]

    ax_peak.barh(y_p, peak_vals, height=0.58, color=peak_colors, edgecolor="#742A2A", alpha=0.9)
    ax_peak.axvline(WIP_ALERT_THRESHOLD, color="#E53E3E", linestyle="--", linewidth=1.5, label=f"警戒線 ({WIP_ALERT_THRESHOLD} Lots)")
    ax_peak.set_yticks(y_p)
    ax_peak.set_yticklabels([f"Top {i+1}: {n}" for i, n in enumerate(sub_peak["station_name"])], fontsize=9, fontweight="bold")
    ax_peak.invert_yaxis()
    # 動態放大 X 軸上限 (預留 62% 空間放右側峰值與時間文字)，確保文字絕不超出右邊框線
    max_p_val = float(np.max(peak_vals)) if len(peak_vals) > 0 else 10.0
    x_limit_peak = max(X_AXIS_MAX * 1.35, WIP_ALERT_THRESHOLD * 1.45, max_p_val * 1.65 + 4.0)
    ax_peak.set_xlim(0, x_limit_peak)
    ax_peak.set_xlabel("最高峰 WIP 數 (Peak Lots)", fontsize=9.5, fontweight="bold")
    ax_peak.set_title(f"【WIP 最高峰 (Peak WIP)】Top {display_n} 站點名單", fontsize=11.5, fontweight="bold", color="#9B2C2C", pad=10)
    ax_peak.grid(axis="x", linestyle=":", alpha=0.6)
    for i, (val, avg_v, p_t) in enumerate(zip(peak_vals, sub_peak["avg_wip"], sub_peak["peak_time"])):
        ax_peak.text(
            val + x_limit_peak * 0.015,
            i,
            f"峰值: {int(val)} Lots (平均: {avg_v:.1f} | {format_sim_time(p_t)})",
            va="center",
            ha="left",
            fontsize=8.5,
            fontweight="bold",
            color="#1A202C",
            clip_on=True,
        )
    ax_peak.legend(loc="lower right", fontsize=8.5, frameon=True, facecolor="#FFFFFF", framealpha=0.9)

    # (A-4) 右下：【平均 WIP (Average WIP)】Top N 瓶頸站點排行圖
    ax_avg = fig_sum.add_subplot(gs_bot[0, 1])
    ax_avg.set_facecolor("#FFFFFF")
    sub_avg = ranked_avg_df.iloc[:display_n].copy()
    y_a = np.arange(len(sub_avg))
    avg_vals = sub_avg["avg_wip"].values.astype(float)
    avg_colors = ["#E53E3E" if v >= WIP_ALERT_THRESHOLD else "#2B6CB0" for v in avg_vals]

    ax_avg.barh(y_a, avg_vals, height=0.58, color=avg_colors, edgecolor="#1A365D", alpha=0.9)
    ax_avg.axvline(WIP_ALERT_THRESHOLD, color="#E53E3E", linestyle="--", linewidth=1.5, label=f"警戒線 ({WIP_ALERT_THRESHOLD} Lots)")
    ax_avg.set_yticks(y_a)
    ax_avg.set_yticklabels([f"Top {i+1}: {n}" for i, n in enumerate(sub_avg["station_name"])], fontsize=9, fontweight="bold")
    ax_avg.invert_yaxis()
    max_a_val = float(np.max(avg_vals)) if len(avg_vals) > 0 else 10.0
    x_limit_avg = max(X_AXIS_MAX * 1.35, WIP_ALERT_THRESHOLD * 1.45, max_a_val * 1.65 + 4.0)
    ax_avg.set_xlim(0, x_limit_avg)
    ax_avg.set_xlabel("平均 WIP 數 (Average Lots)", fontsize=9.5, fontweight="bold")
    ax_avg.set_title(f"【平均 WIP (Average WIP)】Top {display_n} 瓶頸站點名單", fontsize=11.5, fontweight="bold", color="#2C5282", pad=10)
    ax_avg.grid(axis="x", linestyle=":", alpha=0.6)
    for i, (val, q_v, cap_l) in enumerate(zip(avg_vals, sub_avg["avg_queue"], sub_avg["daily_max_lots"])):
        ax_avg.text(
            val + x_limit_avg * 0.015,
            i,
            f"平均: {val:.2f} (排隊: {q_v:.2f} | 日上限: {cap_l:.1f})",
            va="center",
            ha="left",
            fontsize=8.5,
            fontweight="bold",
            color="#1A202C",
            clip_on=True,
        )
    ax_avg.legend(loc="lower right", fontsize=8.5, frameon=True, facecolor="#FFFFFF", framealpha=0.9)

    summary_path = os.path.join(OUTPUT_CHART_DIR, "00_Summary_Dashboard.png")
    fig_sum.savefig(summary_path, dpi=130, bbox_inches="tight", pad_inches=0.25)
    saved_files.append(summary_path)

    # --------------------------------------------------------------------------
    # 圖表 B：Top N 瓶頸站點個別折線分析圖 (一個站點一張圖)
    # --------------------------------------------------------------------------
    # 針對 WIP 最高峰 Top N (或綜合瓶頸 Top N) 生成獨立折線圖
    # 我們以 ranked_peak_df 前 display_n 名生成單站分析圖（包含最高峰與平均統計資訊）
    max_popup_figs = 10  # 避免使用者設 Top 30 時一次彈出 30 個視窗卡住系統

    for idx in range(display_n):
        row = ranked_peak_df.iloc[idx]
        st_name = row["station_name"]
        q_col = f"{st_name}__queue"

        st_wips = history_df[st_name].values.astype(float)
        st_queues = history_df[q_col].values.astype(float)
        max_w = int(row["max_wip"])
        min_w = int(row["min_wip"])
        avg_w = float(row["avg_wip"])
        avg_q = float(row["avg_queue"])
        peak_t = float(row["peak_time"])

        fig_st, ax_st = plt.subplots(figsize=(11.8, 6.4), dpi=100)
        fig_st.patch.set_facecolor("#F8F9FA")
        ax_st.set_facecolor("#FFFFFF")

        # 繪製該站總 WIP (製程中+等待中) 折線與等待中 Queue 折線
        ax_st.plot(
            times,
            st_wips,
            color="#2B6CB0",
            linewidth=2.3,
            label=f"站點總 WIP (製程+等待) [最高: {max_w}, 最低: {min_w}, 平均: {avg_w:.2f}]",
        )
        ax_st.plot(
            times,
            st_queues,
            color="#DD6B20",
            linewidth=1.5,
            linestyle=":",
            label=f"等待中 Queue WIP [平均排隊: {avg_q:.2f} Lots]",
        )

        # 在警戒 WIP 數 (WIP_ALERT_THRESHOLD) 的地方畫一條醒目的紅色虛線
        ax_st.axhline(
            y=WIP_ALERT_THRESHOLD,
            color="#E53E3E",
            linestyle="--",
            linewidth=2.0,
            label=f"警戒 WIP 紅色虛線 ({WIP_ALERT_THRESHOLD} Lots)",
        )

        # 若有超過警戒線的區段，填上淡紅色警示區域 (使用 >= 避免字型缺字亂碼)
        ax_st.fill_between(
            times,
            WIP_ALERT_THRESHOLD,
            st_wips,
            where=(st_wips >= WIP_ALERT_THRESHOLD),
            color="#FEB2B2",
            alpha=0.45,
            interpolate=True,
            label="超載警戒區 (WIP >= 警戒線)",
        )
        ax_st.fill_between(times, 0, st_wips, color="#BEE3F8", alpha=0.25)

        # 標註最高峰點位 (若靠近左右邊界則自動內縮文字框對齊方向，防止超出圖表外)
        max_sim_t = float(np.max(times)) if len(times) > 0 else 1440.0
        if peak_t > max_sim_t * 0.85:
            annot_ha = "right"
            text_x = max(0.0, peak_t - max_sim_t * 0.02)
        elif peak_t < max_sim_t * 0.15:
            annot_ha = "left"
            text_x = min(max_sim_t, peak_t + max_sim_t * 0.02)
        else:
            annot_ha = "center"
            text_x = peak_t

        ax_st.scatter([peak_t], [max_w], color="#E53E3E", s=70, zorder=5)
        ax_st.annotate(
            f"最高峰 WIP: {max_w} Lots\n({format_sim_time(peak_t)})",
            xy=(peak_t, max_w),
            xytext=(text_x, max_w + max(2.2, max_w * 0.09)),
            ha=annot_ha,
            fontsize=9.5,
            fontweight="bold",
            color="#9B2C2C",
            arrowprops=dict(arrowstyle="->", color="#E53E3E", lw=1.2),
            bbox=dict(boxstyle="round,pad=0.3", facecolor="#FFF5F5", edgecolor="#FEB2B2"),
        )

        y_upper = max(X_AXIS_MAX, WIP_ALERT_THRESHOLD + 6, max_w + 6)
        ax_st.set_ylim(0, y_upper)
        ax_st.set_xlim(0, max_sim_t)
        ax_st.set_xticks(tick_vals_full)
        ax_st.set_xticklabels(tick_labels_full, fontsize=9)

        ax_st.set_xlabel("模擬時間 (Time)", fontsize=11, fontweight="bold", color="#1A202C", labelpad=8)
        ax_st.set_ylabel("WIP 數量 (Lots)", fontsize=11, fontweight="bold", color="#1A202C", labelpad=8)
        ax_st.set_title(
            f"【Top {idx+1} 瓶頸站點 WIP 趨勢分析圖】{st_name}\n"
            f"(機台數: {int(row['capacity'])} 台 | 混線加權 C/T: {row['ct']:.1f}m | 每日最大產能: {row['daily_max_lots']:.1f} Lots/天 | "
            f"最高 WIP: {max_w} / 最低: {min_w} / 平均: {avg_w:.2f})",
            fontsize=12,
            fontweight="bold",
            color="#1A202C",
            pad=12,
        )

        ax_st.grid(True, linestyle=":", alpha=0.65, color="#A0AEC0")
        ax_st.spines["top"].set_visible(False)
        ax_st.spines["right"].set_visible(False)
        ax_st.legend(loc="upper left", frameon=True, facecolor="#FFFFFF", edgecolor="#CBD5E0", fontsize=9.5)

        safe_name = (
            st_name.replace("/", "_")
            .replace("\\", "_")
            .replace(" ", "_")
            .replace(":", "")
            .replace(".", "")
        )
        chart_path = os.path.join(OUTPUT_CHART_DIR, f"{idx+1:02d}_Top{idx+1}_{safe_name}.png")
        fig_st.savefig(chart_path, dpi=130, bbox_inches="tight", pad_inches=0.2)
        saved_files.append(chart_path)

        # 若超過視窗彈出上限則關閉 Figure，僅保留檔案與前 10 張視窗
        if not show_windows or idx >= max_popup_figs:
            plt.close(fig_st)

    if not show_windows:
        plt.close(fig_sum)

    print(f"📊 [圖表產生完成] 已自動生成 1 張總結圖與 {display_n} 張單站折線分析圖，並儲存至 '{OUTPUT_CHART_DIR}/' 資料夾：")
    for fpath in saved_files:
        print(f"   • {fpath}")
    print()

    if show_windows:
        print("[視覺化啟動] 正在開啟【總結圖】與【Top N 瓶頸站點分析折線圖】視窗...")
        plt.show()

    return saved_files


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

    fig, ax = plt.subplots(figsize=(15.4, 9.6), dpi=100)
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
        "站點當下 WIP 批數 (含製程中 In-Process + 等待中 Queue Lots)",
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
        f"【03SEC & 01SG 動態模擬】日投: {int(daily_lots)} Lots | 比例 {ratio_03sec*100:.0f}:{ratio_01sg*100:.0f} | "
        f"每日生產: 60×{work_hours:.0f}h×{run_pct:.0f}% = {daily_equip_mins:.0f} min/天 | 跨站搬送: {transfer_min:.1f}m",
        fontsize=13,
        fontweight="bold",
        color="#1A202C",
        y=0.96,
    )
    sub_title = ax.set_title(
        f"當前模擬時間: Day 1, 00:00  |  【產線總 WIP: 0 Lots】 = 站點 WIP: 0 (製程:0+等待:0) + 人員搬送: 0 Lots",
        fontsize=10.5,
        fontweight="bold",
        color="#2C5282",
        pad=12,
    )

    avg_station_wip = float(history_df["total_station_wip"].mean())
    avg_in_service = float(history_df["total_in_service"].mean())
    avg_queue = float(history_df["total_queue"].mean())
    avg_transfer = float(history_df["in_transfer"].mean())
    avg_line_wip = float(history_df["total_line_wip"].mean())

    display_n = max(1, min(int(top_n), len(ranked_df)))
    box_display_n = min(display_n, 10)
    top_lines = [
        f"【全線平均產線總 WIP: {avg_line_wip:.1f} Lots = 站點 WIP {avg_station_wip:.1f} (製程:{avg_in_service:.1f}+等待:{avg_queue:.1f}) + 人員搬送 {avg_transfer:.1f} Lots】",
        f"【全線 Top {display_n} 瓶頸站點 (關閉此動畫視窗後將自動顯示總結圖與單站分析折線圖)】",
    ]
    for idx in range(box_display_n):
        r = ranked_df.iloc[idx]
        top_lines.append(
            f"TOP {idx+1}: {r['station_name']} "
            f"(平均WIP: {r['avg_wip']:.2f} [最高峰:{int(r['max_wip'])}] | 日產能上限: {r['daily_max_lots']:.1f} Lots)"
        )
    if display_n > box_display_n:
        top_lines.append(f"... (其餘 Top {box_display_n+1}~{display_n} 請見 Console 終端機輸出)")

    box_fontsize = 8.2 if box_display_n <= 5 else 7.4
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
        total_station_wip = int(np.sum(wips))
        total_queue = int(np.sum(queues))
        total_in_service = total_station_wip - total_queue
        total_line_wip = total_station_wip + in_trans
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
            f"模擬時間: {time_str} ({int(current_time)}m)  |  "
            f"【產線總 WIP: {total_line_wip} Lots】 = 站點 WIP: {total_station_wip} (製程:{total_in_service}+等待:{total_queue}) "
            f"+ 人員搬送: {in_trans} Lots  |  紅色超載站: {alert_count} 站"
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

    print("[動態模擬啟動] 正在開啟 Matplotlib 動態長條圖視窗 (關閉此視窗後，將自動顯示總結圖與各站折線分析圖)...")
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


def parse_sim_mode(mode_str: Optional[str]) -> str:
    """解析模擬模式開關 ('dynamic' 或 'static')。"""
    if mode_str is None:
        return DEFAULT_SIM_MODE
    s = str(mode_str).strip().lower()
    if s in ("2", "static", "s", "靜態", "靜態模擬"):
        return "static"
    return "dynamic"


def get_simulation_parameters() -> Tuple[str, float, float, float, float, float, float, int]:
    """
    開放讓使用者透過命令列參數或互動式輸入設定七大核心參數：
    0. 模擬模式開關 (--mode，'dynamic' 動態模擬 或 'static' 靜態模擬，預設 dynamic)
    1. 每日投入 Lot 數量 (--lots，預設 23，建議 10~30)
    2. 03SEC 與 01SG 混合比例 (--mix，預設 7:3)
    3. 每天生產時間 (小時) (--work-hours，預設 20 小時)
    4. 機台 RUN% 稼動率 (--run-pct，預設 80%，故每天實際生產時間 = 60 * 20 * 0.8 = 960 分鐘)
    5. 站與站之間人員搬送時間 (分鐘) (--transfer-min，預設 5.0 分鐘，0 代表不計)
    6. 顯示 Top N 瓶頸站點數量 (--top-n，預設 3)
    """
    parser = argparse.ArgumentParser(description="03SEC & 01SG 全線多站點動態/靜態混線模擬器")
    parser.add_argument(
        "--mode",
        type=str,
        default=None,
        help="設定模擬模式：'dynamic' (1: 動態模擬) 或 'static' (2: 靜態模擬)",
    )
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

    mode_raw = args.mode
    daily_lots = args.lots
    mix_str = args.mix
    work_hours = args.work_hours
    run_pct = args.run_pct
    transfer_min = args.transfer_min
    top_n = args.top_n

    is_interactive = (os.environ.get("HEADLESS_TEST", "0") != "1") and sys.stdin.isatty()

    if is_interactive:
        print("\n" + "=" * 80)
        print("⚙️  【模擬參數設定】(直接按 Enter 可快速套用預設值)")
        print("=" * 80)

        if mode_raw is None:
            default_mode_label = "1: 動態模擬" if DEFAULT_SIM_MODE == "dynamic" else "2: 靜態模擬"
            inp = input(
                f"  0. 請選擇模擬模式 (1: 動態模擬[動畫+總結+分析圖], 2: 靜態模擬[直接出總結+分析圖]) [預設 {default_mode_label}]: "
            ).strip()
            mode_raw = inp if inp != "" else DEFAULT_SIM_MODE

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

    sim_mode = parse_sim_mode(mode_raw)
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
    return sim_mode, daily_lots, ratio_03sec, ratio_01sg, work_hours, run_pct, transfer_min, top_n


def main():
    # Step 1: 取得使用者設定的模擬參數 (含動態/靜態模擬開關)
    sim_mode, daily_lots, ratio_03sec, ratio_01sg, work_hours, run_pct, transfer_min, top_n = (
        get_simulation_parameters()
    )
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
        f"(模式: {'動態模擬' if sim_mode == 'dynamic' else '靜態模擬'} | "
        f"每日設備實際生產時間 = 60×{work_hours:.0f}×{run_pct/100:.2f} = {daily_equip_mins:.0f} 分鐘/台)】"
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
        "每日最大產能(Lots/天)",
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

    # Step 4: 自動識別並印出全線 WIP 最高/最低/平均總結、平均 WIP Top N 與最高峰 WIP Top N 名單
    ranked_avg_df, ranked_peak_df = analyze_bottlenecks(stations_df, history_df, top_n=top_n)

    is_headless = os.environ.get("HEADLESS_TEST", "0") == "1"

    # Step 5: 根據「動態模擬 (dynamic)」或「靜態模擬 (static)」開關執行視覺化
    if sim_mode == "dynamic" and not is_headless:
        # 5A. 動態模擬：先播放全線多站點橫向長條圖動態動畫
        create_and_show_animation(
            stations_df=stations_df,
            history_df=history_df,
            ranked_df=ranked_avg_df,
            daily_lots=daily_lots,
            ratio_03sec=ratio_03sec,
            ratio_01sg=ratio_01sg,
            work_hours=work_hours,
            run_pct=run_pct,
            transfer_min=transfer_min,
            interarrival_time=simulator.interarrival_time,
            top_n=top_n,
        )

    # 5B. 產生「模擬總結圖」與「Top N 瓶頸站點折線分析圖 (一個一張 + 警戒紅虛線)」
    # (無論動態模擬結束後或靜態模擬，皆會自動生成圖檔並彈出圖表視窗)
    generate_summary_and_station_charts(
        history_df=history_df,
        ranked_avg_df=ranked_avg_df,
        ranked_peak_df=ranked_peak_df,
        daily_lots=daily_lots,
        ratio_03sec=ratio_03sec,
        ratio_01sg=ratio_01sg,
        work_hours=work_hours,
        run_pct=run_pct,
        transfer_min=transfer_min,
        top_n=top_n,
        show_windows=(not is_headless),
    )


if __name__ == "__main__":
    main()
