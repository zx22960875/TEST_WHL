# -*- coding: utf-8 -*-
"""
Created on Sun Nov  9 13:57:22 2025

@author: PC
"""

from rust_stdf_helper import stdf_to_log_sheet_stats_v6,analyzeSTDF, generate_database, stdf_to_xlsx, TestIDType, norm_cdf, norm_ppf, empirical_cdf
class Sig:
    def emit(self, *args): print(*args)
class Stop:
    stop = False
    

class DummyProgress:
    def emit(self, percent: int):
        print(f"進度: {percent / 100:.2f}%")

class DummyStop:
    stop = False  # 若要中途停止可改成 True

# 輸入 STDF 路徑與輸出 CSV 路徑
input_stdf = "C:/data/sample.stdf"
output_csv = "C:/data/sample_log.xlsx"

# 呼叫 Rust 端函式
stdf_to_log_sheet_stats_v6(
    r"C:\Users\PC\Downloads\ROOS_20140728_131230.stdf",
    "example.xlsx",
    TestIDType.TestNumberAndName,  # 或 TestIDType.TestNumberOnly
    DummyProgress(),               # 用於顯示進度
    DummyStop(),                   # 用於中途停止
)
