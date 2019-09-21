## sp_schedule

This script will help transform the output of a shedule.xlsx file so that it can be easy to create a screenshot of each day

### command

    poetry install
    # add the current Ask schedule (Excel file; Sheet1) to this folder
    # rename that file schedule.xlsx
    # Change the Time column (10:00:00 AM) format to Integer or Text (10)
    poetry run python horaire.py

### output 


![alt text](https://github.com/guinslym/sp_transform_schedule/blob/master/example_output.png "Schedule")


## todo
1. Put every "standalone statement" in a function 
2. Removed unused column automatically
3. test it
4. create package and publish it to Pypi 
