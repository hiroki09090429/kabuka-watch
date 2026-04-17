#!/usr/bin/env python3
"""GitHub Actions用 株価データ更新スクリプト（Excel不要・クラウド実行）"""
import yfinance as yf
import json, math, os
from datetime import datetime, timezone, timedelta
import warnings
warnings.filterwarnings('ignore')

JST = timezone(timedelta(hours=9))
now_jst = datetime.now(JST)
now_str = now_jst.strftime("%Y-%m-%d %H:%M")

json_path = "data.json"

US_TICKERS = {"VYM","HDV","SPYD"}
stocks = [
    ("9986","蔵王産業","卸売"),("3076","あいHD","卸売"),("8130","サンゲツ","卸売"),
    ("2659","サンエー","小売"),("3333","あさひ","小売"),
    ("4008","住友精化","化学"),("4042","東ソー","化学"),("4097","高圧ガス工業","化学"),
    ("8309","三井住友トラストG","銀行"),("8725","MS&ADインシュアランスHD","保険"),
    ("8593","三菱HCキャピタル","その他金融"),("8584","ジャックス","その他金融"),
    ("6785","鈴木","電気機器"),("7723","愛知時計電機","精密機器"),
    ("3231","野村不動産HD","不動産"),("3003","ヒューリック","不動産"),
    ("2169","CDS","サービス"),("9757","船井総研HD","サービス"),
    ("9769","学究社","サービス"),("4641","アルプス技研","サービス"),
    ("3817","SRAホールディングス","情報・通信"),("3901","マークラインズ","情報・通信"),
    ("4674","クレスコ","情報・通信"),("2003","日東富士製粉","食料品"),
    ("1414","ショーボンドHD","建設"),("1928","積水ハウス","建設"),
    ("6345","アイチコーポレーション","機械"),("9364","上組","倉庫・運輸関連"),
    ("9381","エーアイティー","倉庫・運輸関連"),("5388","クニミネ工業","ガラス・土石製品"),
    ("7989","立川ブラインド工業","金属製品"),("7820","ニホンフラッシュ","その他製品"),
    ("7994","オカムラ","その他製品"),("4540","ツムラ","医薬品"),
    ("1343","NF・J-REIT ETF","J-REIT"),
    ("VYM","VYM（米国高配当）","米国ETF"),("HDV","HDV（米国高配当）","米国ETF"),
    ("SPYD","SPYD（米国高配当）","米国ETF"),
]

tickers_yf = [c+".T" if c not in US_TICKERS else c for c,_,_ in stocks]

print(f"{now_str} データ取得中...")
all_data = yf.download(tickers_yf, period="20d", interval="1d", auto_adjust=True, progress=False)
closes = all_data["Close"] if "Close" in all_data.columns else all_data.xs("Close", axis=1, level=0)

jp_tickers = [t for t in closes.columns if t not in US_TICKERS]
valid_dates = closes[jp_tickers].dropna(how='all').index[-5:]
date_labels = [d.strftime("%m/%d") for d in valid_dates]

# 既存のlongtermデータを保持
try:
    with open(json_path) as f:
        existing = json.load(f)
    lt_stocks = existing.get("longterm",{}).get("stocks",[])
    longterm_updated_at = existing.get("longterm_updated_at","")
except:
    lt_stocks = []
    longterm_updated_at = ""

weekly_stocks = []
for code, name, sector in stocks:
    key = code+".T" if code not in US_TICKERS else code
    prices = []
    for d in valid_dates:
        try:
            v = closes.loc[d, key]
            prices.append(round(float(v),2) if (v==v and not math.isnan(float(v))) else None)
        except:
            prices.append(None)
    valid = [p for p in prices if p is not None]
    week_pct = round((valid[-1]/valid[0]-1)*100, 2) if len(valid)>=2 else None
    weekly_stocks.append({"code":code,"name":name,"sector":sector,
                          "is_us":code in US_TICKERS,"week_pct":week_pct,"prices":prices})

json_data = {
    "updated_at": now_str,
    "weekly_updated_at": now_str,
    "longterm_updated_at": longterm_updated_at,
    "weekly": {"dates": date_labels, "stocks": weekly_stocks},
    "longterm": {"stocks": lt_stocks}
}
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(json_data, f, ensure_ascii=False, indent=2)

print(f"  更新完了: {date_labels[0]}〜{date_labels[-1]}")
print(f"  銘柄数: {len(weekly_stocks)}")
