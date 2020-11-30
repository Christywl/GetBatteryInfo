import argparse
import csv
import time
import wmi
from ctypes import *


class PowerClass(Structure): 
    _fields_ = [('ACLineStatus', c_byte), 
                ('BatteryFlag', c_byte),
                ('BatteryLifePercent', c_byte),
                ('Reserved1', c_byte),
                ('BatteryLifeTime', c_ulong),
                ('BatteryFullLifeTime', c_ulong)]


def get_designed_and_full_charged_capacity():

    c = wmi.WMI() # http://www.voidcn.com/article/p-bmwrycny-btk.html
    t = wmi.WMI(moniker="//./root/wmi")
    battery_designed = c.CIM_Battery(Caption='Portable Battery')
    for i, b in enumerate(battery_designed):
        print('\nBattery %d Design Capacity: %d mWh' % (i, b.DesignCapacity or 0))
    
    battery_full_charged = t.ExecQuery('Select * from BatteryFullChargedCapacity')
    for i, b in enumerate(battery_full_charged):
        print('\nBattery %d Fully Charged Capacity: %d mWh' % (i, b.FullChargedCapacity))


def get_battery_life_percent():

    power_class = PowerClass()  # http://www.itxm.cn/post/2069.html
    windll.kernel32.GetSystemPowerStatus(byref(power_class))

    return str(power_class.BatteryLifePercent)


def get_battery_remaining_capacity():

    t = wmi.WMI(moniker="//./root/wmi")
    battery = t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
    for i, b in enumerate(battery):
        return str(b.RemainingCapacity)


def init_csv(filename):

    with open(filename, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Minutes", "Date", "Time", "BatteryLifePercent", "BatteryRemainingCapacity(mWh)"])


def save_to_csv(filename, intervals, num):

    for counter in range(num + 1):
        print('\n' + '*' * 40)
        if counter == 0:
            print('\nInitial Battery Data')
            print('Log Number: %d, Testing Duration: %d minutes' % (counter, intervals * counter))
        elif counter > 0:
            print('Log Number: %d, Testing Duration: %d minutes' % (counter, intervals * counter))
        
        date = time.strftime("%Y-%m-%d", time.localtime(time.time()))
        battery_time = time.strftime("%H:%M:%S", time.localtime(time.time()))

        battery_life_percent = get_battery_life_percent()
        print('BatteryLifePercent: %s%%' % battery_life_percent)
        
        remaining_capacity = get_battery_remaining_capacity()
        print('RemainingCapacity: %s mWh' % remaining_capacity)

        with open(filename, 'a', newline='') as csvFile:
            csv_writer = csv.writer(csvFile)
            csv_writer.writerow([str(counter * intervals) + ' minutes', date, battery_time,
                                 battery_life_percent + "%", remaining_capacity])

        if counter < num:
            print('Waiting for %d minutes ...' % intervals)
            time.sleep(intervals * 60)
        elif counter == num:
            print('\n' + '*' * 40)
            print('\n********* The testing finished *********')
            
        counter += 1


def main():
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--i", required=True, type=int, metavar='Number',
                        help="the time interval(minutes) that you choose")
    parser.add_argument("--n", required=True, type=int, metavar='Number',
                        help="the log numbers that you choose")
    args = parser.parse_args()
    print('\nThe time interval(minutes) that you choose is %d minutes' % args.i)
    print('The log number that you choose is %d' % args.n)
    
    get_designed_and_full_charged_capacity()
    
    file_time = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime(time.time()))
    filename = "BatteryInfo" + "-" + file_time + ".csv"
    init_csv(filename)
    save_to_csv(filename, args.i, args.n)


if __name__ == "__main__": 
    try:
        main()
    except IOError as e:
        print(e)
    except KeyboardInterrupt:
        pass
    except Exception:
        raise
