from pywinauto import Application                       # for window detection and connection
from pywinauto import keyboard as pywinauto_keyboard    # for sending key commands into DAS
from pynput import keyboard as pynput_keyboard          # for hotkeys to start and stop the script
from pywinauto.findwindows import ElementNotFoundError  # for error handling
from pywinauto.base_wrapper import ElementNotEnabled    # for error handling
import time                                             # for debug
import ctypes                                           # for message box display

# Connect to active app
print('\nLocating application.. \n')
app = Application(backend='win32').connect(class_name='AfxMDIFrame100').window(class_name='AfxMDIFrame100')
montage_window = app.Montage32770
#montage_window.print_control_identifiers()

# Define controls (Need P button, shares field, price field, Route dropdown)
price = montage_window.PEdit
shares = montage_window.Edit1
position_button = montage_window.PButton

def main_function(sub_function):
    try:
        print('\nExecuting main_function \n')
        start_time = time.time()
        
        # Get current position
        #position_button.click()
        position_direction = shares.get_line(0)
        print(position_direction)
        if  '-' in position_direction:
            trade_direction = 'Short'
            position_size = position_direction.replace('-', '')
        else:
            trade_direction = 'Long'  
            position_size = position_direction
        print(position_size)
        # Perform calculations
        stop_distance = round((SET_RISK/int(position_size)), 2)
        price.set_edit_text(stop_distance)
        print('\n Calculated stop_distance: ', stop_distance)
        
        # For StopUpdate hotkey
        if sub_function == 'StopUpdate':
            if trade_direction == 'Short':
                pywinauto_keyboard.send_keys(das_stop_update_short)
                print('\nDetected Short position. Sending ', das_stop_update_short, ' das_stop_update_short hotkey to DAS')
            else:    
                pywinauto_keyboard.send_keys(das_stop_update_long)
                print('\nDetected Long position. Sending ', das_stop_update_long, ' das_stop_update_long hotkey to DAS')
        
        # For TargetUpdate hotkeys
        elif sub_function == 'TargetUpdate1':
            if trade_direction == 'Short':       # Send DAS hotkey
                pywinauto_keyboard.send_keys(das_target_update_short1)
                print('\nDetected Short position. Sending ', das_target_update_short1, ' das_target_update_short1 hotkey to DAS')
            else:    
                pywinauto_keyboard.send_keys(das_target_update_long1)
                print('\nDetected Long position. Sending ', das_target_update_long1, ' das_target_update_long1 hotkey to DAS')
        
        elif sub_function == 'TargetUpdate2':
            if trade_direction == 'Short':       # Send DAS hotkey
                pywinauto_keyboard.send_keys(das_target_update_short2)
                print('\nDetected Short position. Sending ', das_target_update_short2, ' das_target_update_short2 hotkey to DAS')
            else:    
                pywinauto_keyboard.send_keys(das_target_update_long2)
                print('\nDetected Long position. Sending ', das_target_update_long2, ' das_target_update_long2 hotkey to DAS')
        
        elif sub_function == 'TargetUpdate3':
            if trade_direction == 'Short':       # Send DAS hotkey
                pywinauto_keyboard.send_keys(das_target_update_short3)
                print('\nDetected Short position. Sending ', das_target_update_short3, ' das_target_update_short3 hotkey to DAS')
            else:    
                pywinauto_keyboard.send_keys(das_target_update_long3)
                print('\nDetected Long position. Sending ', das_target_update_long3, ' das_target_update_long3 hotkey to DAS')
        
        print('\n--- %s seconds ---' % (time.time() - start_time))
        print_message()

    except ElementNotFoundError:
        print('\nError: active element not found. Make sure DAS trader montage window is selected')
        message_box('Error', 'Active element not found. \nMake sure DAS trader montage window is selected.', 0)
        quit()
    except ElementNotEnabled:
        print('\nError: active element is not enabled. Make sure price field is enabled')
        message_box('Error', 'Active element is not enabled. \nMake sure price field is enabled.', 0)
        quit()
    # except Exception as error:
    #     print('\nUnexpected error... ', str(error))
    #     message_box('Error', 'Unexpected error... {}'.format(str(error)), 0)
    #     pynput_keyboard.Listener.StopException
    #     quit()

def stop_update_function():
    main_function('StopUpdate')

def target_update_function1():
    main_function('TargetUpdate1')

def target_update_function2():
    main_function('TargetUpdate2')

def target_update_function3():
    main_function('TargetUpdate3')

def message_box(title, text, style):     # Popup message box function
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

def exit_function():
    print('\nExecuting exit_function \n')
    quit()

def print_message():                     # Print welcome message in console
    print('\nPress assigned hotkey to run the script or ', system_hotkey, ' to exit.')

# Define user variables
target_update_hotkey1 = '<Alt>+<F9>'      # User hotkey to run the script (free text format)
target_update_hotkey2 = '<Alt>+<F10>'     # User hotkey to run the script (free text format)
target_update_hotkey3 = '<Alt>+<F11>'     # User hotkey to run the script (free text format)
stop_update_hotkey = '<Alt>+<F12>'        # User hotkey to run the script (free text format)

system_hotkey = '<Ctrl>+<Alt>+k'         # System hotkey to exit the script (free text format)

das_target_update_short1 = '^%{F1}'        # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
das_target_update_short2 = '^%{F2}'        # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
das_target_update_short3 = '^%{F3}'        # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
das_stop_update_short = '^%{F4}'           # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)

das_target_update_long1 = '^%+{F1}'        # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
das_target_update_long2 = '^%+{F2}'        # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
das_target_update_long3 = '^%+{F3}'        # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)
das_stop_update_long = '^%+{F4}'           # User hotkey to pass into DAS ('+' for Shift, '^' for Ctrl, '%' for Alt)

SET_RISK = 65                            # Set dollar risk per trade here

print_message()

# Script hotkey body
with pynput_keyboard.GlobalHotKeys({
        stop_update_hotkey : stop_update_function,
        target_update_hotkey1 : target_update_function1,
        target_update_hotkey2 : target_update_function2,
        target_update_hotkey3 : target_update_function3,
        system_hotkey : exit_function}) as mapped_result:
    mapped_result.join()