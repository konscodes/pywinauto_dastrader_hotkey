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