'''
    Author: Brenden Morton
    Date Started: 6/05/2021
    Description: Test code for validating SD cards data corruption protection.
'''


#################################################
# Interfacing with powersupply (PSU)            
import serial as ps                       

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
