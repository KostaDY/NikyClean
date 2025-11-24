# TRADES GUI — Transaction Manager

## Overview
A Python/tkinter GUI for entering and managing stock trades in TRADES.xlsx.

## Features
- Add, delete, undo delete
- Reload tickers from Stock sheet
- Date stored as true Excel date
- Safe delete logic
- Logging in CSV + Excel

## Workbook Structure
- Transactions sheet: Date, Prefix, Ticker, Number, Price, Act
- Log sheet: auto-built from TRADES_log.csv
- Stock sheet: Ticker + Description

## Requirements
Python 3.13.x, tkinter, openpyxl.

## Running
```
python3 trades_extended_gui.py
```
