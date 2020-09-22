# GetBatteryInfo
Get the battery info for laptops on Windows and log the battery info to a .CSV file which contains information about BatteryLifePercent(in %) and RemainingCapacity in mWh units (milliwatt-hours).

## How do I use it?
### For Python 2.x on Windows
  1. Install Python 2.x on Windows and add Python to your PATH environment variable 
  2. Install Python modules with pip: `pywin32` and `wmi`
  3. Run the script:

```sh
     Usage: getbatteryinfo_py2.py [-h] --i Number --n Number

     optional arguments:
     -h, --help show this help message and exit
     --i  Number  the time interval(minutes) that you choose
     --n  Number  the log numbers that you choose

```

For example:

```sh
    > python getbatteryinfo_py2.py --i 5 --n 10
```
  4. The output file `BatteryInfo-****.csv` is saved in the current directo
  
### For Python 3.x on Windows
  1. Install Python 3.x on Windows and add Python to your PATH environment variable
  2. Install Python modules with pip: `pywin32` and `wmi`
  3. Run the script:

```sh
     Usage: getbatteryinfo_py2.py [-h] --i Number --n Number

     optional arguments:
     -h, --help show this help message and exit
     --i  Number  the time interval(minutes) that you choose
     --n  Number  the log numbers that you choose

```

For example:

```sh
    > python getbatteryinfo_py3.py --i 5 --n 10
```

  4. The output file `BatteryInfo-****.csv` is saved in the current directory.

