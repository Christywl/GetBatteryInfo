# GetBatteryInfo
Get the battery info for laptops on Windows and log the battery info to a .CSV file which contains information about BatteryLifePercent(in %) and RemainingCapacity in mWh units (milliwatt-hours).

## Using the script
  1. Install python 2.x on Windows
  2. Install python packages with pip: `pywin32` and `wmi`
  3. Run the script:

```sh
     Usage: getbatteryinfo.py [-h] --i Number --n Number

     optional arguments:
     -h, --help show this help message and exit
     --i  Number  the time interval(minutes) that you choose
     --n  Number  the log numbers that you choose

```

For example:

```sh
    > python getbatteryinfo.py --i 5 --n 10
```
  4. The output file `BatteryInfo-****.csv` is saved in the current directory.

