'''
    Author: Brenden Morton
    Date Started: 6/05/2021
    Description: Test code for validating SD cards data corruption protection.
'''

#################################################
# Interfacing with powersupply (PSU)
import serial

# Interfacing with the IFD over CAN bus
#import can

# Timeing intervals
import time

# Writing data to excel spreadsheet
from openpyxl import Workbook

# File system/directories
import os  

# Obtaining current date when a test begins
from datetime import date
##################################################


# Iterations for testing
num_iterations = 10000

# (seconds) Minimum delay in between cycling power 
min_delay = 30

# (seconds) Maximum delay in between cycling power 
max_delay = 150

# Using the minimum delay for testing intially
delay = min_delay

# Baudrate for the device (default)
baud = 109200

# Port for the device (default)
port = 'COM1'

# Date
current_date = date.today()

# Instantiate a workbook object for storing data
wb = Workbook()

# Creating a sheet in the workbook
test_data = wb.create_sheet(f"Test Data ({current_date})")

# Initializes the workbook for data storage
def init_wb():

    # Start in the first column
    col = 1
    
    # Write test number
    test_data.cell(column=col, row=1).value = 'test_no'
    
    # Next column
    col += 1

    # Write test result    
    test_data.cell(column=col, row=1).value = 'result'

    # Next column
    col +=1 

    # Write test data
    test_data.cell(column=col, row=1).value = 'test_data'

    # wb.remove("sheet")
    wb.save("SD_TEST_DATA.xlsx")

    return

# Records the test data in an Excel spreadsheet
def record_test(test_no, test_data, result):

    # Start in the first column
    col = 1
    
    # Write test number
    test_data.cell(column=col, row=(test_no + 1)).value = test_no
    
    # Next column
    col += 1

    # Write test result    
    test_data.cell(column=col, row=(test_no + 1)).value = result

    # Next column
    col +=1 

    # Write test data
    test_data.cell(column=col, row=(test_no + 1)).value = test_data

    # Saving the workbook
    wb.save("SD_TEST_DATA.xlsx")

    return

# Assess the result of a test
def assert_test(test_no, test_data):

    # TODO
    """
        Need to write logic for when:
            - Fail and Pass occurs
        
        This is dependent on testing.


        Cody said that he used another serial port to read the data of the device
        This can be used to determine whether or not a failure has occured
        In order to figure this out, we need to test the device
    """
    # TODO

    fail_occurs = None
    pass_occurs = None

    if fail_occurs:
        result = 'FAIL'
    elif pass_occurs:
        result = 'PASS'

    record_test(test_no, test_data, result)

def run():

    init_wb()

    # Instaniating the serial interface safely
    with serial.Serial() as dev:

        ''' Initializing the device '''

        # Setting baud
        dev.baudrate = baud

        # Claiming port
        dev.port = port

        # Opening the connection to the device
        dev.open()

        '''
            To write to the device:

                dev.write()

        '''

        res = input("Select one of the following:\n1. Default\n2. Custom\n")

        if res == '2':
            num_iterations = int(input("Number of iterations: "))
            delay = int(input("Delay (in seconds): "))
        else:
            print("Default mode")

        
        for i in range(num_iterations):
            # do something with pyserial (ps)
            # turn on
            print("turn on")
            pass

            # delay amount of time
            time.sleep(delay)
            print("End of delay")

            # do something with pyserial
            # turn off
            print("turn off")
            pass


            # TODO
            '''
            Need to figure out what a failure looks like in the
            automation process for logging
            '''
            # Not sure if this will be used
            # delay amount of time
            time.sleep(delay)
            print("End of delay")


def run_2(port):
    print(f"Inside run_2: {port}.\nTypeof: {type(port)}")
    with serial.Serial() as dev:

        dev.baudrate = 9600
        # Claiming port
        dev.port = port
        
        # print(f"PORT {port}")

        for i in range(100):
            print("Turning on")

            # Opening the connection to the device
            try:
                dev.open()
            except Exception as e:
                print(f"Error: {e}")
                return        
            time.sleep(5)

            print("Turning off")
            dev.close()
            time.sleep(5)

# Loop to find the port
def loop():
    for i in range(1,100):
        port = f"COM{i}"
        print(f"COM {i}")
        run_2(port) 
        time.sleep(3)

def test():
    dev = serial.Serial()
    dev.port = ""
    dev.baudrate = 9600
    dev.open()
    time.sleep(10)
    dev.close()


# Works for testing
def ps_fun_time(time_delay):

    ps = serial.Serial(
        port='COM3',
        baudrate=9600,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.EIGHTBITS,
        rtscts=True,
        timeout=0
    )

    while True:

        try:
            ps.open()
            print("Port opened!")
        except:
            print("Could not open port -> Exception thrown")
            ps.close()
        
        for i in range(time_delay):
            print(i)
            time.sleep(1)

    ps.open()
    time.sleep(5)
    ps.close()

# ps_fun_time(15)

def input_mode():
    start = time.time()
    # state = True
    ps = serial.Serial(
        port='COM3',
        baudrate=9600,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.EIGHTBITS,
        rtscts=True,
        timeout=0
    )
    ps.close()

    ifd = serial.Serial(
        port='COM4',
        baudrate=9600,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.EIGHTBITS,
        rtscts=True,
        timeout=0
    )



    while True:
        print(f"\n~Enter~\n[1] Turn on\n[2] Turn off\n[3] To end program")
        res = input()
        res = int(res)

        if res == 1:
            try:
                ps.open()
                print("Turning on")
                
            except Exception as e:
                print(f"Exception: {e}")
                pass
            # os.system("clear")

        elif res == 2:
            try:
                ps.close()
                print("Turning off")
            except Exception as e:
                print(f"Exception: {e}")
                pass
            # os.system("clear")
        else:
            print("Ending program")
            return


def can_test(delay):
    can.rc['interface'] = 'serial'
    can.rc['channel'] = '/dev/ttyUSB0'
    can.rc['bitrate'] = 9600

    from can.interface import Bus

    bus = Bus()

    



# ps_fun_time(15)
# can_test(15)
input_mode()
