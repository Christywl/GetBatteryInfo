import os 
import win32con 
import sys 
import time 
import csv
import argparse
import wmi

from ctypes import * 


class PowerClass(Structure): 
    _fields_ = [('ACLineStatus', c_byte), 
            ('BatteryFlag', c_byte), 
            ('BatteryLifePercent', c_byte), 
            ('Reserved1',c_byte), 
            ('BatteryLifeTime',c_ulong), 
            ('BatteryFullLifeTime',c_ulong)] 

def init_csv(filename):

    f = open(filename, "wb")
    csv_writer = csv.writer(f)
    csv_writer.writerow(["Minutes", "Date", "Time", "BatteryLifePercent", "BatteryRemainingCapacity(mWh)"])
    f.close()

def GetBatteryLifePercent():

    powerclass = PowerClass()
    windll.kernel32.GetSystemPowerStatus( byref(powerclass) ) 

    return str(powerclass.BatteryLifePercent)

def GetDesignandFullChargeCapacity():

    c = wmi.WMI()
    t = wmi.WMI(moniker = "//./root/wmi")
    batts1 = c.CIM_Battery(Caption = 'Portable Battery')
    for i, b in enumerate(batts1):
        print '\nBattery %d Design Capacity: %d mWh' % (i, b.DesignCapacity or 0)
    
    batts2 = t.ExecQuery('Select * from BatteryFullChargedCapacity')
    for i, b in enumerate(batts2):
        print ('\nBattery %d Fully Charged Capacity: %d mWh' % 
            (i, b.FullChargedCapacity))

def GetBatteryRemainingCapacity():

    t = wmi.WMI(moniker = "//./root/wmi")   
    batts = t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
    for i, b in enumerate(batts):
        remaningBatteryCapacity = str(b.RemainingCapacity)

    return remaningBatteryCapacity

def save_to_csv(filename, intervals, num):

    counter = 0
    for counter in range(num + 1):
        print '\n' + '*' * 40
        if counter == 0:
            print '\nInitial Battery Data'
            print 'Log Number: %d, Test Duration: %d mins' %(counter, intervals * counter)
        elif counter > 0:
            print '\nLog Num: %d, Test Duration: %d mins' %(counter, intervals * counter)
        
        date = time.strftime("%Y-%m-%d", time.localtime(time.time()))
        battery_time = time.strftime("%H:%M:%S", time.localtime(time.time()))

        BatteryLifePercent = GetBatteryLifePercent()
        print 'BatteryLifePercent: %s%%' %BatteryLifePercent
        
        RemainingCapacity = GetBatteryRemainingCapacity()
        print 'RemainingCapacity: ' + str(RemainingCapacity)
        
        csvFile = open(filename, "ab")
        csv_writer = csv.writer(csvFile)
        csv_writer.writerow([str(counter * intervals) + " mins", date, battery_time, BatteryLifePercent + "%", RemainingCapacity])
        csvFile.close()

        if counter < num:
            print 'Waiting for %d minutes...' % intervals
            time.sleep(intervals)
        elif counter == num:
            print '\n' + '*' * 40
            print '\n******** Completing the testing ********'
            
        counter += 1

def main():
    
    parser = argparse.ArgumentParser();
    parser.add_argument("--i", required=True, type=int, metavar='Number', help="the time interval(minutes) that you choose")
    parser.add_argument("--n", required=True, type=int, metavar='Number', help="the log numbers that you choose")
    args = parser.parse_args()
    print 'The time interval(minutes) that you choose is %d mins' %args.i
    print 'The log number that you choose is %d' %args.n
    
    GetDesignandFullChargeCapacity()
    
    file_time = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime(time.time()))
    filename = "BatteryInfo" + "-" + file_time + ".csv"
    init_csv(filename)
    save_to_csv(filename, args.i, args.n)
    
if __name__ == "__main__": 
    try:
        main()
    except IOError as e:
        print e
    except KeyboardInterrupt:
        pass
    except Exception:
        raise