'''
    Author: Brenden Morton
    Date Started: 6/05/2021
    Description: Test code for validating SD cards data corruption protection.
'''


#################################################
# Interfacing with powersupply (PSU)
import serial                       

# Timeing intervals
import time

# Writing data to excel spreadsheet
from openpyxl import Workbook

# File system/directories
import os   

# Writing output to an excel file for analysis
import excel_writer
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



# Instantiate a workbook object for storing data
wb = Workbook()

# Creating a sheet in the workbook
test_data = wb.create_sheet("Test Data")


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

    wb.save("SD_TEST_DATA.xlsx")

    return

# Assess the result of a test
def assert_test(test_no, test_data):

    # TODO
    """
        Need to write logic for when:
            - Fail and Pass occurs
        
        This is dependent on testing.
    """

    # TODO
    fail_occurs = None
    pass_occurs = None


    if fail_occurs:
        result = 'FAIL'
    elif pass_occurs:
        result = 'PASS'

    record_test(test_no, test_data, result)


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

init_wb()
