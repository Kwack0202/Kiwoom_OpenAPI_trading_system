@ECHO ON

title Kiwoom Start

cd C:\Kiwoom_trading\Kiwoom_stock_analysis

call C:\Users\coden\anaconda3\Scripts\activate.bat trading_hyun
python __init__.py

cmd.exe
