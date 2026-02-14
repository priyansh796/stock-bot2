import yfinance as yf
import pandas as pd
import numpy as np
from ta.momentum import RSIIndicator
from openpyxl import load_workbook
import os

MARKET_CAP_LIMIT = 5000 * 10**7
MONTHLY_HISTORY = "15y"
WEEKLY_HISTORY = "5y"
PORTFOLIO_FILE = "portfolio.xlsx"


def super_smoother(price, period):
    a1 = np.exp(-1.414 * np.pi / period)
    b1 = 2 * a1 * np.cos(1.414 * np.pi / period)
    c2 = b1
    c3 = -a1 * a1
    c1 = 1 - c2 - c3

    filt = np.zeros(len(price))

    for i in range(2, len(price)):
        filt[i] = (
            c1 * (price[i] + price[i - 1]) / 2
            + c2 * filt[i - 1]
            + c3 * filt[i - 2]
        )

    return filt


# ================= LOAD STOCK LIST =================
stocks_df = pd.read_csv("nse_stocks.csv")
symbols = stocks_df['SYMBOL'].dropna().tolist()
stocks = [symbol + ".NS" for symbol in symbols]

weekly_buy = []
monthly_buy = []
sell_signals = []


# ================= SIGNAL GENERATION =================
for stock in stocks:

    print(f"Processing {stock} ...")

    try:
        ticker = yf.Ticker(stock)
        info = ticker.info

        market_cap = info.get("marketCap", None)
        if market_cap is None or market_cap < MARKET_CAP_LIMIT:
            continue

        monthly_df = ticker.history(period=MONTHLY_HISTORY, interval="1mo")
        if len(monthly_df) < 260:
            continue

        m_close = monthly_df['Close'].values

        monthly_df['SSF_50'] = super_smoother(m_close, 50)
        monthly_df['SSF_100'] = super_smoother(m_close, 100)
        monthly_df['SSF_250'] = super_smoother(m_close, 250)

        m_latest = monthly_df.iloc[-1]
        m_prev = monthly_df.iloc[-2]

        monthly_below_count = sum([
            m_prev['Close'] < m_prev['SSF_50'],
            m_prev['Close'] < m_prev['SSF_100'],
            m_prev['Close'] < m_prev['SSF_250']
        ])

        monthly_cross = (
            m_prev['Close'] < m_prev['SSF_50']
            and m_latest['Close'] > m_latest['SSF_50']
        )

        if monthly_below_count >= 2 and monthly_cross:
            monthly_buy.append(stock)

        weekly_df = ticker.history(period=WEEKLY_HISTORY, interval="1wk")
        if len(weekly_df) < 260:
            continue

        w_close = weekly_df['Close'].values

        weekly_df['SSF_50'] = super_smoother(w_close, 50)
        weekly_df['SSF_100'] = super_smoother(w_close, 100)
        weekly_df['SSF_250'] = super_smoother(w_close, 250)

        w_latest = weekly_df.iloc[-1]
        w_prev = weekly_df.iloc[-2]

        weekly_below_count = sum([
            w_prev['Close'] < w_prev['SSF_50'],
            w_prev['Close'] < w_prev['SSF_100'],
            w_prev['Close'] < w_prev['SSF_250']
        ])

        weekly_cross = (
            w_prev['Close'] < w_prev['SSF_50']
            and w_latest['Close'] > w_latest['SSF_50']
        )

        if weekly_below_count >= 2 and weekly_cross:
            weekly_buy.append(stock)

        rsi = RSIIndicator(monthly_df['Close'], window=14)
        monthly_df['RSI'] = rsi.rsi()
        monthly_df['RSI_MA'] = monthly_df['RSI'].rolling(14).mean()

        rsi_prev = monthly_df.iloc[-2]
        rsi_latest = monthly_df.iloc[-1]

        if rsi_prev['RSI'] > rsi_prev['RSI_MA'] and rsi_latest['RSI'] < rsi_latest['RSI_MA']:
            sell_signals.append(stock)

    except Exception:
        continue


# ================= PORTFOLIO MEMORY =================

if os.path.exists(PORTFOLIO_FILE):
    portfolio_df = pd.read_excel(PORTFOLIO_FILE, sheet_name="Portfolio")
    owned_stocks = set(portfolio_df['Stock'])
else:
    portfolio_df = pd.DataFrame(columns=["Stock", "Status"])
    owned_stocks = set()

signals = []

# BUY LOGIC → add if not owned
for stock in weekly_buy + monthly_buy:
    if stock not in owned_stocks:
        portfolio_df.loc[len(portfolio_df)] = [stock, "OWNED"]
        signals.append([stock, "BUY"])

# SELL LOGIC → only if owned
for stock in sell_signals:
    if stock in owned_stocks:
        portfolio_df = portfolio_df[portfolio_df['Stock'] != stock]
        signals.append([stock, "SELL"])


signals_df = pd.DataFrame(signals, columns=["Stock", "Signal"])


# ================= WRITE EXCEL =================
with pd.ExcelWriter(PORTFOLIO_FILE, engine="openpyxl", mode="w") as writer:
    portfolio_df.to_excel(writer, sheet_name="Portfolio", index=False)
    signals_df.to_excel(writer, sheet_name="Signals", index=False)

print("Excel file updated successfully.")
