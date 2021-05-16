# ISS Control Program
# handles ISS systems

# imports
from random import random, randint
from openpyxl import load_workbook
import xlwings as xw
import os
import time

# set filepath
wb = xw.Book.caller()
filename = '/Console.xlsm'
path = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + filename
owb = load_workbook(path)

# doesn't do anything, but is here just in case
class Control:
    print('SOMETHING')


# clears and resets the Console.xlsm file
def end_sim():
    wb = xw.Book.caller()
    wb.sheets[0].range('H1').value = 'SIMULATION TERMINATED'
    time.sleep(3)
    wb.sheets[0].range(
        'H1,O1,A6,A13,A11,B8:B10,E6,E8,E10,E12,E14,E16,I6,I8,I10,I12,I14,I16,M6,M8,M10,M12,M14,M16').value = ''
    wb.sheets[1].range('A2:D2,A5:C5').value = ''
    time.sleep(1)
    wb.sheets[0].range('O1').value = ''


# health check commands listed below
# divided into systems and then subsystems
def system_check(system):
    # access workbook and sheet
    wb = xw.Book.caller()
    a_e = wb.sheets[1].range('A2').value
    c_e = wb.sheets[1].range('B2').value
    e_e = wb.sheets[1].range('C2').value
    t_e = wb.sheets[1].range('D2').value

    # ADCS health check
    if system == 'ADCS' or system == 'adcs' or system == 1:
        if a_e == 2:
            wb.sheets[0].range('A6').value = 'SYSTEM ERROR'
        else:
            wb.sheets[0].range('A6').value = 'SYSTEM NORMAL'

    # CRONUS health check
    elif system == 'CRONUS' or system == 'cronus' or system == 2:
        if c_e == 2:
            # Error example generator
            MAX_LIMIT = 255
            random_string = ''
            for _ in range(10):
                random_integer = randint(0, MAX_LIMIT)
                # Keep appending random characters using chr(x)
                random_string += (chr(random_integer))
            wb.sheets[0].range('E8').value = random_string
            wb.sheets[1].range('B8').value = random_string
            wb.sheets[0].range('E6').value = 'SYSTEM ERROR'
            wb.sheets[0].range('E10').value = 'DATA ERROR'
            wb.sheets[0].range('E12').value = 'UPLINK ERROR'
            wb.sheets[0].range('E14').value = 'DOWNLINK ERROR'
            wb.sheets[0].range('E16').value = 'TRANSMISSION ERROR'
        else:
            wb.sheets[0].range('E6').value = 'SYSTEM NORMAL'
            wb.sheets[0].range('E8').value = 'SEND DATA CORRECTION'
            wb.sheets[0].range('E10').value = 'SEND DATA CORRECTION'
            wb.sheets[0].range('E12').value = 'SEND DATA CORRECTION'
            wb.sheets[0].range('E14').value = 'SEND DATA CORRECTION'
            wb.sheets[0].range('E16').value = 'SEND DATA CORRECTION'

    # EPS health check
    elif system == 'EPS' or system == 'eps' or system == 3:
        if e_e == 2:
            wb.sheets[0].range('I6').value = 'SYSTEM ERROR'
            wb.sheets[0].range('I8').value = 'POWER FAILURE'
            wb.sheets[0].range('I10').value = 'ARRAY FAILURE'
            wb.sheets[0].range('I12').value = 'PDU FAILURE'
        elif e_e == 1:
            wb.sheets[0].range('I6').value = 'SYSTEM NORMAL'
            wb.sheets[0].range('I8').value = 'ACTIVATE POWER SYSTEMS'
            wb.sheets[0].range('I10').value = 'ALIGN ARRAYS'
        else:
            wb.sheets[0].range('I6').value = 'SYSTEM NORMAL'
            wb.sheets[0].range('I8').value = 'POWER NORMAL'

    # TCS health check
    elif system == 'TCS' or system == 'tcs' or system == 4:
        if t_e == 2:
            wb.sheets[0].range('M6').value = 'SYSTEM ERROR'
            wb.sheets[0].range('M8').value = 'DATA ERROR'
            wb.sheets[0].range('M10').value = 'RAD. FAILURE'
            wb.sheets[0].range('M12').value = 'PUMP FAILURE'
            wb.sheets[0].range('M14').value = 'RAD. FAILURE'
            wb.sheets[0].range('M16').value = 'TEMP. WARNING'
        else:
            wb.sheets[0].range('M6').value = 'SYSTEM NORMAL'

    # end


# function to verify system health. Selectable by acronym or place in console order left to right
def verify_system(system):
    # access workbook and sheet
    wb = xw.Book.caller()
    a_e = wb.sheets[1].range('A2').value
    c_e = wb.sheets[1].range('B2').value
    e_e = wb.sheets[1].range('C2').value
    t_e = wb.sheets[1].range('D2').value

    # ADCS verify command
    if system == 'ADCS' or system == 'adcs' or system == 1:
        if a_e == 2:
            wb.sheets[0].range('B8').value = random()
            wb.sheets[0].range('B9').value = random()
            wb.sheets[0].range('B10').value = random()
        elif a_e == 1:
            print('')
        else:
            wb.sheets[0].range('B8').value = wb.sheets[1].range('A5').value
            wb.sheets[0].range('B9').value = wb.sheets[1].range('B5').value
            wb.sheets[0].range('B10').value = wb.sheets[1].range('C5').value

    # CRONUS verify command
    elif system == 'CRONUS' or system == 'cronus' or system == 2:
        if c_e == 2:
            # Error example generator
            MAX_LIMIT = 255
            random_string = ''
            for _ in range(10):
                random_integer = randint(0, MAX_LIMIT)
                # Keep appending random characters using chr(x)
                random_string += (chr(random_integer))
            wb.sheets[0].range('E8').value = random_string
            wb.sheets[1].range('B8').value = random_string

        elif c_e == 1:
            # DATA CORRUPTED print
            wb.sheets[0].range('E8').value = 'DATA CORRUPTED'

        else:
            wb.sheets[0].range('E8').value = 'DATA TRANSMITTING'

    # EPS verify command (included in case of implementation)
    elif system == 'EPS' or system == 'eps' or system == 3:
        print()

    # TCS verify command (included in case of implementation)
    elif system == 'TCS' or system == 'tcs' or system == 4:
        print()


# system reset method. Each system is selectable by its acronym or order in the console.
def system_reset(system):
    wb = xw.Book.caller()

    # ADCS reset
    if system == 'ADCS' or system == 'adcs' or system == 1:
        wb.sheets[1].range('A2').value = 1
        wb.sheets[0].range('A13').value = "CRAFT MISALIGNED"
        wb.sheets[0].range('A6').value = 'SYSTEM RESETTING'
        time.sleep(3)
        wb.sheets[0].range('A6').value = 'PLEASE WAIT'
        time.sleep(3)
        wb.sheets[0].range('A6').value = 'PERFORM HEALTH CHECK'

    # CRONUS reset
    elif system == 'CRONUS' or system == 'cronus' or system == 2:
        wb.sheets[1].range('B2').value = 1
        wb.sheets[0].range('E8').value = 'DATA CORRUPTED'
        wb.sheets[0].range('E6').value = 'SYSTEM RESETTING'
        time.sleep(3)
        wb.sheets[0].range('E6').value = 'PLEASE WAIT'
        time.sleep(3)
        wb.sheets[0].range('E6').value = 'PERFORM HEALTH CHECK'

    # EPS reset
    elif system == 'EPS' or system == 'eps' or system == 3:
        wb.sheets[1].range('C2').value = 1
        wb.sheets[0].range('I16').value = '3%'
        wb.sheets[0].range('I6').value = 'SYSTEM RESETTING'
        time.sleep(3)
        wb.sheets[0].range('I6').value = 'PLEASE WAIT'
        time.sleep(3)
        wb.sheets[0].range('I6').value = 'PERFORM HEALTH CHECK'
        wb.sheets[0].range('I8').value = 'ACTIVATE POWER SYSTEMS'
        wb.sheets[0].range('I10').value = 'ALIGN ARRAYS'
        wb.sheets[0].range('I12').value = 'ACTIVATE PDU'

    # TCS reset
    elif system == 'TCS' or system == 'tcs' or system == 4:
        wb.sheets[1].range('D2').value = 1
        wb.sheets[0].range('M8').value = 'ACTIVATE CONTROLS'
        wb.sheets[0].range('M10').value = 'ALIGN ARRAYS'
        wb.sheets[0].range('M12').value = 'ACTIVATE PUMPS'
        wb.sheets[0].range('M14').value = 'SET RAD. LEVEL'
        wb.sheets[0].range('M16').value = 'SET TEMP.'
        wb.sheets[0].range('M6').value = 'SYSTEM RESETTING'
        time.sleep(3)
        wb.sheets[0].range('M6').value = 'PLEASE WAIT'
        time.sleep(3)
        wb.sheets[0].range('M6').value = 'PERFORM HEALTH CHECK'


# blanket function for console correction commands. fixes are separated into 'system' and 'subsystem' i.e. to
# correct EPS solar alignment, system would be 'eps' or 3 and subsystem would be 0
def fix_command(system, subsystem):
    # access workbook and sheet
    wb = xw.Book.caller()
    a_e = wb.sheets[1].range('A2').value
    c_e = wb.sheets[1].range('B2').value
    e_e = wb.sheets[1].range('C2').value
    t_e = wb.sheets[1].range('D2').value

    # ADCS fixes
    if system == 'ADCS' or system == 'adcs' or system == 1:
        # fixes attitude if system has been reset
        if subsystem == 'att':
            if a_e == 1:
                wb.sheets[1].range('A2').value = 0
                wb.sheets[0].range('B8').value = wb.sheets[1].range('A5').value
                wb.sheets[0].range('B9').value = wb.sheets[1].range('B5').value
                wb.sheets[0].range('B10').value = wb.sheets[1].range('C5').value
            else:
                print('nope')
        # fixes coupling if system has been reset
        elif subsystem == 'coupling':
            if a_e == 2:
                wb.sheets[0].range('A13').value = "ALIGNMENT ERROR"
                print()
            elif a_e == 1:
                wb.sheets[0].range('A13').value = "CRAFT MISALIGNED"
            elif a_e == 0:
                wb.sheets[0].range('A13').value = "CRAFT READY FOR COUPLING"

    # CRONUS fixes
    elif system == 'CRONUS' or system == 'cronus' or system == 2:
        # fixes data irregularity if system has been reset
        if subsystem == 0:
            if c_e == 2:
                print()
            if c_e == 1:
                wb.sheets[1].range('B2').value = 0
                wb.sheets[0].range('E8').value = 'DATA NORMAL'
                wb.sheets[0].range('E10').value = 'SYSTEMS NORMAL'
            elif c_e == 0:
                wb.sheets[0].range('E10').value = 'SYSTEMS NORMAL'
        # tests uplink capabilities if system has been reset
        elif subsystem == 1 or subsystem == 'uplink':
            if c_e == 2:
                wb.sheets[0].range('E12').value = wb.sheets[0].range('B8')
            elif c_e == 1:
                wb.sheets[0].range('E12').value = 'SEND DATA CORRECTION'
            elif c_e == 0:
                wb.sheets[0].range('E12').value = 'UPLINK ACTIVE'
        # tests downlink capabilities if system has been reset
        elif subsystem == 2 or subsystem == 'downlink':
            if c_e == 2:
                wb.sheets[0].range('E14').value = wb.sheets[0].range('B8')
            elif c_e == 1:
                print('something')
                wb.sheets[0].range('E14').value = 'SEND DATA CORRECTION'
            elif c_e == 0:
                wb.sheets[0].range('E14').value = 'DOWNLINK ACTIVE'
        # transmits GPS data if system has been reset
        elif subsystem == 3 or subsystem == 'GPS':
            if c_e == 2:
                wb.sheets[0].range('E16').value = 'TRANSMISSION ERROR'
            elif c_e == 1:
                wb.sheets[0].range('E16').value = 'SEND DATA CORRECTION'
            elif c_e == 0:
                wb.sheets[0].range('E16').value = 'DATA TRANSMITTING'

    # EPS fixes
    elif system == 'EPS' or system == 'eps' or system == 3:
        # align array if system has been reset
        if subsystem == 'array':
            if e_e == 1:
                wb.sheets[1].range('C2').value = 0
                wb.sheets[0].range('I10').value = 'ARRAYS ALIGNED'
            else:
                print()
        # engage pdu if system has been reset
        elif subsystem == 'pdu':
            wb.sheets[0].range('I12').value = 'ACTIVE'
            wb.sheets[0].range('I8').value = 'POWER NORMAL'
        # request battery levels if system has been reset
        elif subsystem == 'checkbatt':
            if e_e == 2:
                wb.sheets[0].range('I14').value = '25%'
            elif e_e == 1:
                wb.sheets[0].range('I14').value = '50%'
            elif e_e == 0:
                wb.sheets[0].range('I14').value = 'CHARGING'
        # check array efficiency if system has been reset
        elif subsystem == 'checkarr':
            if e_e == 2:
                wb.sheets[0].range('I16').value = 'POWER FAILURE'
            elif e_e == 1:
                wb.sheets[0].range('I16').value = '3%'
            elif e_e == 0:
                wb.sheets[0].range('I16').value = '14%'

    # TCS fixes
    elif system == 'TCS' or system == 'tcs' or system == 4:
        # align radiator arrays if system has been reset
        if subsystem == 'radiator':
            if t_e == 1:
                wb.sheets[1].range('D2').value = 0
                wb.sheets[0].range('M10').value = 'ARRAYS ALIGNED'
            else:
                print()
        # engage pumps if system has been reset
        elif subsystem == 'pumps':
            wb.sheets[0].range('M12').value = 'PUMPS ACTIVE'
            wb.sheets[0].range('M8').value = 'CONTROLS ACTIVE'
        # set radiator levels if system has been reset
        elif subsystem == 'setrad':
            if t_e == 0:
                wb.sheets[0].range('M14').value = 'DEFAULT SET.'
        # set temperature if system has been reset
        elif subsystem == 'settemp':
            if t_e == 0:
                wb.sheets[0].range('M16').value = 'ADJUSTING TEMP.'


def main():
    # initialize console
    wb.sheets[0].range('H1').value = 'SIMULATION RUNNING'
    wb.sheets[0].range('A6,E6,I6,M6').value = 'SYSTEM NORMAL'
    wb.sheets[0].range('A13').value = "ALIGNMENT ERROR"
    wb.sheets[0].range('E8').value = 'DATA TRANSMITTING'
    wb.sheets[0].range('E10').value = 'SYSTEMS NORMAL'
    wb.sheets[0].range('E12').value = 'UPLINK ACTIVE'
    wb.sheets[0].range('E14').value = 'DOWNLINK ACTIVE'
    wb.sheets[0].range('E16').value = 'DATA TRANSMITTING'
    wb.sheets[0].range('I8').value = 'POWER NORMAL'
    wb.sheets[0].range('I10').value = 'ARRAYS ALIGNED'
    wb.sheets[0].range('I12').value = 'ACTIVE'
    wb.sheets[0].range('I14').value = '25%'
    wb.sheets[0].range('I16').value = 'POWER FAILURE'
    wb.sheets[0].range('M8').value = 'THERMAL CONTROLS ACTIVE'
    wb.sheets[0].range('M10').value = 'ARRAYS ALIGNED'
    wb.sheets[0].range('M12').value = 'PUMPS ACTIVE'
    wb.sheets[0].range('M14').value = 'DEFAULT SET.'
    wb.sheets[0].range('M16').value = 'TEMP. LOW'
    wb.sheets[0].range('B8').value = 0.79
    wb.sheets[0].range('B9').value = -3.36
    wb.sheets[0].range('B10').value = -3.99
    wb.sheets[1].range('A5').value = 0.79
    wb.sheets[1].range('B5').value = -3.36
    wb.sheets[1].range('C5').value = -3.99

    # initialize errors
    wb.sheets[1].range('A2:D2').value = 2

    # Error example generator
    MAX_LIMIT = 255
    random_string = ''
    for _ in range(10):
        random_integer = randint(0, MAX_LIMIT)
        # Keep appending random characters using chr(x)
        random_string += (chr(random_integer))
    wb.sheets[1].range('B8').value = random_string

    # I have no clue why I wrote the following rather than implementing it -
    # similarly to the other code, however, I am afraid to change it -
    # in case I had a good reason to write it this way.
    wb.sheets[0].range('A11').value = ("""=IFS('Code Tab'!A2 = 2, "ATTITUDE WARNING", 'Code Tab'!A2 = 1, 
    "ADJUST ATTITUDE", 'Code Tab'!A2 = 0, "ATTITUDE NORMAL")""")
