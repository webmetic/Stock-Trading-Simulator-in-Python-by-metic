# metic
I probably won't be updating this code. You can do whatever you'd like with it. Please drop me an email if you use it. It would make me happy :D

You could download stock data for different stocks or update the stocks prices. You would have to convert all the dates into text and prices need to be limited to 2 decimal places.

![image](https://user-images.githubusercontent.com/74499053/134215255-d8804994-3b99-4dfe-9e86-753b8a9c2759.png)

I have 50 popular stocks in my list and they update every 4 seconds (4 Seconds in Simulation = 1 Day in Real Life). This gives better performance.

Download stock data on the excel file using this:
https://stackoverflow.com/questions/32545316/how-to-write-data-to-excel-using-python-for-stock-data-being-pulled-from-yahoo
(I DID NOT CODE THE ABOVE AND PLEASE CREDIT THEM IF POSSIBLE)

Download the above folder and run the python file (named 'Stock Trading Simulator.py')
Make sure all the files (inside the folder are always) inside the same folder(same file path).

Display Settings
Display Resolution: 1920x1080
Size: 125%
(sorry these are my default settings)

You can view 50 stock prices update in the first window (top left). You can make purchases in the second one (top right). You can view the performance of a stock in the third window (bottom left) and view your transactions in the fourth one (bottom right). 

Bug Fixes:
Use `pip install xlrd==1.2.0` to solve the xlrd module issue. 
