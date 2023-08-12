import time
from datetime import datetime


def k():
    print("k function active")

while True:
    now = datetime.now()
    current_time=now.strftime("%I:%M:%S:%P")
    print("TIME : " + current_time)
    current_date=now.strftime("%d-%m-%Y")
    print("DATE : " + current_date)

    h=now.strftime("%H")
    hour=str(h)
    print(hour)
    

    full=["0","1","2","3","4","5","6","7"]
    day=["8","9","10","11","12","13","14","15","16"]
    half=["17","18","19","20","21","22","23"]
    
    if hour in day:
        print("day")
        k()
    elif hour in half:
        print("half")
    elif hour in full:
        print("full")
    