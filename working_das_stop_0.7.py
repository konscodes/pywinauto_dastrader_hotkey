'''
GUI automation script to suplement DAS hotkey

Idea: calculate stop distance based on current position size and set dollar risk

DAS hotkey script (Long): 
Route=Stop;Share=Pos;StopPrice=50/Pos;StopPrice=AvgCost-StopPrice;StopType=MARKET;TIF=DAY;SELL=Send;

Problem: DAS scripting language doesnt support Risk/Position opperation; moreover DAS API is 100$ per month
Solution: read montage controls with python library for GUI automation pywinauto; perform calculation in python; return desired value;

Notes: 
- script will put the new stop distance in Price field and hit the hotkey to send the order
- the montage window should be active
- two hotkeys are defined within a script; one for activating the python part; another for passing into DAS

Requirements:
- Pythin 3.8.0
- Pywinauto 0.6.8
- Pynput 1.6.7

Preparation:
1. Set first key combination in python to start the script (StopUpdateHotkey in following example)
2. Add the below hotkey to DAS
Long:  Route=Stop;StopPrice=AvgCost-Price;StopType=MARKET;TIF=DAY;SELL=Send;Route=Limit;
Short: Route=Stop;StopPrice=AvgCost+Price;StopType=MARKET;TIF=DAY;BUY=Send;Route=Limit;
3. Set second key combination in python to pass into your DAS (DasStopUpdateShort and DasStopUpdateLong in following example)
4. Set your dollar risk in python script (SetRisk variale)
5. Run the script

How to use:
1. Make sure that montage window is active
2. Hit the hotkey defined in Preparation Step1
3. Profit ;) the correct order should be placed

Change log:
v0.7
- Added the ability to assign different hotkey sets through subfunction
- Removed debug and testing code to speedup the execution

v0.6 
- Added position direction detection
- Added DAS hotkeys for long and short

v0.5
- Added hotkey module to run the script
- Added error handling

v0.4
- Basic functionality reached

''' 

from pywinauto import Application                       # for window detection and connection
from pywinauto import keyboard as PywinautoKeyboard     # for sending key commands into DAS
from pynput import keyboard as PynputKeyboard           # for hotkeys to start and stop the script
from pywinauto.findwindows import ElementNotFoundError  # for error handling
from pywinauto.base_wrapper import ElementNotEnabled    # for error handling
import time                                             # for debug
import ctypes                                           # for message box display

def MainFunction(SubFunction):
    try:
        print("\nExecuting MainFunction \n")
        start_time = time.time()
        # Connect to active app
        app = Application(backend="win32").connect(class_name="AfxMDIFrame100")

        # Select active window
        MontageWindow = app.window(class_name="AfxMDIFrame100")

        # Define controls (Need P button, Shares field, Price field, Route dropdown)
        Price = MontageWindow.PEdit
        Shares = MontageWindow.Edit2

        # Get current position
        MontageWindow.P.click()
        PositionDirectional = Shares.get_line(0)
        if  "-" in PositionDirectional:
            TradeDirection = "Short"
            PositionSize = PositionDirectional.replace("-", "")
        else:
            TradeDirection = "Long"  
            PositionSize = PositionDirectional

        # Perform calculations
        StopDistance = round((SetRisk/int(PositionSize)), 2)
        Price.set_edit_text(StopDistance)   # Replace variables
        print("\n Calculated StopDistance: ", StopDistance)
        
        # For StopUpdate hotkey
        if SubFunction == "StopUpdate":
            if TradeDirection == "Short":
                PywinautoKeyboard.send_keys(DasStopUpdateShort)
                print("\nDetected Short position. Sending ", DasStopUpdateShort, " DasStopUpdateShort hotkey to DAS")
            else:    
                PywinautoKeyboard.send_keys(DasStopUpdateLong)
                print("\nDetected Long position. Sending ", DasStopUpdateLong, " DasStopUpdateLong hotkey to DAS")
        
        print("\n--- %s seconds ---" % (time.time() - start_time))
        PrintMessage()

    except ElementNotFoundError:
        print("\nError: active element not found. Make sure DAS trader montage window is selected")
        MessageBox("Error", "Active element not found. \nMake sure DAS trader montage window is selected.", 0)
        quit()
    except ElementNotEnabled:
        print("\nError: active element is not enabled. Make sure Price field is enabled")
        MessageBox("Error", "Active element is not enabled. \nMake sure Price field is enabled.", 0)
        quit()
    except Exception as Error:
        print("\nUnexpected error... ", str(Error))
        MessageBox("Error", "Unexpected error...", 0)
        PynputKeyboard.Listener.StopException
        quit()

def StopUpdateFunction():
    MainFunction("StopUpdate")

def MessageBox(title, text, style):     # Popup message box function
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

def ExitFunction():
    print("\nExecuting ExitFunction \n")
    quit()

def PrintMessage():                     # Print welcome message in console
    print("\nPress assigned hotkey to run the script or ", SystemHotkey, " to exit.")

# Define user variables
StopUpdateHotkey = "<Alt>+<F12>"        # User hotkey to run the script (free text format)
SystemHotkey = "<Ctrl>+<Alt>+k"         # System hotkey to exit the script (free text format)
DasStopUpdateShort = "^%{F4}"           # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
DasStopUpdateLong = "^%+{F4}"           # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
SetRisk = 65                            # Set dollar risk per trade here

PrintMessage()

# Script hotkey body
with PynputKeyboard.GlobalHotKeys({
        StopUpdateHotkey : StopUpdateFunction,
        SystemHotkey : ExitFunction}) as MappedResult:
    MappedResult.join()