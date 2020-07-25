@echo off
title Divan Bleu Invoicing System
echo Make sure you've downloaded the latest ClientRevenueReportList.xlsx from GoRendezVous and saved in Software folder
pause
echo Would you like to make invoices for this month?(Y/N)
set/p  "choicecurrent=>"
if %choicecurrent%==Y goto RUN
if %choicecurrent%==y goto RUN
if %choicecurrent%==N goto SELECTMONTH
if %choicecurrent%==n goto SELECTMONTH


:RUN
"C:\Users\bimon\AppData\Local\Programs\Python\Python38\python.exe" "_Software\Divan_Bleu_Invoicing.py"
pause
goto END

:SELECTMONTH
echo Would you like to also change the year?(Y/N)
set/p  "choiceyear=>"
if %choiceyear%==Y goto SELECTYEAR
if %choiceyear%==y goto SELECTYEAR
if %choiceyear%==N goto RUNMONTH
if %choiceyear%==n goto RUNMONTH

:SELECTYEAR
echo Please enter the month of the invoices
set/p "choicemonth=>"
echo Please enter the year of the invoices
set/p "choiceyear=>"
"C:\Users\bimon\AppData\Local\Programs\Python\Python38\python.exe" "_Software\Divan_Bleu_Invoicing.py" --month=%choicemonth% --year=%choiceyear%
pause
goto END

:RUNMONTH
echo Please enter the number for the month of the invoices
set/p "choicemonth=>"
"C:\Users\bimon\AppData\Local\Programs\Python\Python38\python.exe" "_Software\Divan_Bleu_Invoicing.py" --month=%choicemonth%
pause
goto END

:END