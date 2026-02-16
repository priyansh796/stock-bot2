import yfinance as yf
import pandas as pd
import numpy as np
from ta.momentum import RSIIndicator
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


def load_portfolio():
    if os.path.exists(PORTFOLIO_FILE):
        try:
            df = pd.read_excel(PORTFOLIO_FILE, sheet_name="Portfolio")
            return df
        except:
            pass

    return pd.DataFrame(columns=["Stock", "Status"])


def main():

    stocks = pd.read_csv("nse_stocks.csv")

    portfolio_df = load_portfolio()
    owned_stocks = set(portfolio_df["Stock"])

    weekly_buy = []
    monthly_buy = []
    sell_signals = []

    signals = []

    for symbol in stocks["Symbol"]:

        ticker = f"{symbol}.NS"
        print(f"Processing {ticker} ...")

        try:
            stock = yf.Ticker(ticker)
            info = stock.info

            market_cap = info.get("marketCap")

            if market_cap is None or market_cap < MARKET_CAP_LIMIT:
                continue

            hist = stock.history(period=MONTHLY_HISTORY)

            if hist.empty:
                continue

            close = hist["Close"].values

            smooth = super_smoother(close, 10)

            rsi = RSIIndicator(pd.Series(close), window=14).rsi()

            if smooth[-1] > smooth[-2] and rsi.iloc[-1] < 30:
                monthly_buy.append(symbol)

            if smooth[-1] < smooth[-2] and rsi.iloc[-1] > 70:
                sell_signals.append(symbol)

        except Exception as e:
            print("Error:", e)

    # ================= SIGNAL GENERATION =================

    for stock in monthly_buy:
        if stock not in owned_stocks:
            signals.append([stock, "BUY"])
            portfolio_df.loc[len(portfolio_df)] = [stock, "OWNED"]

    for stock in sell_signals:
        if stock in owned_stocks:
            signals.append([stock, "SELL"])
            portfolio_df = portfolio_df[portfolio_df["Stock"] != stock]

    # ================= DATA SAFETY (VERY IMPORTANT) =================

    signals_df = pd.DataFrame(
        signals if signals else [["NONE", "NO SIGNAL"]],
        columns=["Stock", "Signal"]
    )

    if portfolio_df.empty:
        portfolio_df = pd.DataFrame(
            [["NONE", "EMPTY"]],
            columns=["Stock", "Status"]
        )

    # ================= WRITE EXCEL (INDENTATION SAFE) =================

    with pd.ExcelWriter(PORTFOLIO_FILE, engine="openpyxl", mode="w") as writer:
        portfolio_df.to_excel(writer, sheet_name="Portfolio", index=False)
        signals_df.to_excel(writer, sheet_name="Signals", index=False)

    print("portfolio.xlsx force-created")


if __name__ == "__main__":
    main()



