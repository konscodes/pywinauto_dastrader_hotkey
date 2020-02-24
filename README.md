# pywinauto_dastrader_hotkey
##GUI automation script for DAS trader

Idea:
Calculate stop distance based on current position size and set dollar risk

Idealy should be done by the following DAS hotkey (Long):
Route=Stop;Share=Pos;StopPrice=50/Pos;StopPrice=AvgCost-StopPrice;StopType=MARKET;TIF=DAY;SELL=Send;

Problem:
DAS scripting language doesnt support Risk/Position opperation; moreover DAS API is 100$ per month

Solution:
Read montage controls with python library for GUI automation pywinauto. Perform calculation in python and return desired value of StopDistance. Put the value back into montage and send the order using DAS hotkey.

Notes:
- script will put the new stop distance in Price field and hit the hotkey to send the order
- the montage window should be active
- two hotkeys are defined within a script; one for activating the python part; another for passing into DAS

Requirements:
Pythin 3.8.0
Pywinauto 0.6.8
Pynput 1.6.7

Preparation:
- Download working working_das_stop script and install all the requirements above
- Set first key combination in python to start the script (StopUpdateHotkey in following example)
- Add the below hotkey to DAS
Long: Route=Stop;StopPrice=AvgCost-Price;StopType=MARKET;TIF=DAY;SELL=Send;Route=Limit; 
Short: Route=Stop;StopPrice=AvgCost+Price;StopType=MARKET;TIF=DAY;BUY=Send;Route=Limit;
- Set second key combination in python to pass into your DAS (DasStopUpdateShort and DasStopUpdateLong in following example)
- Set your dollar risk in python script (SetRisk variale)
- Run the script (open cmd in the directory with script; enter: py script_name.py)

How to use:
1. Make sure that montage window is active
2. Hit the hotkey defined in Preparation Step1
3. Profit ;) the correct order should be placed


Change log:
v0.7
Added the ability to assign different hotkey sets through subfunction
Removed debug and testing code to speedup the execution

v0.6
Added position direction detection
Added DAS hotkeys for long and short

v0.5
Added hotkey module to run the script
Added error handling

v0.4
Basic functionality reached