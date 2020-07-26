
@echo off
title Divan Bleu Invoicing System
echo.
echo Make sure you've downloaded the latest ClientRevenueReportList.xlsx from GoRendezVous and saved in Invoices folder
echo.
pause

:HOME
echo.
echo Main Menu
echo 1 - Create and E-mail this month's invoices
echo 2 - Create this month's invoices
echo 3 - Create past invoices
echo 0 - Exit program
echo.
echo Enter Selection: 
set/p "choice=>"
if %choice%==1 goto RUN_EMAIL
if %choice%==2 goto RUN
if %choice%==3 goto SELECTPAST
if %choice%==0 goto END
goto HOME
pause

:RUN
"C:\Toolkits\anaconda3-5.2.0\python.exe" "_Software\Divan_Bleu_Invoicing.py"
echo.
goto HOME

:RUN_EMAIL
"C:\Toolkits\anaconda3-5.2.0\python.exe" "_Software\Divan_Bleu_Invoicing.py" --email
echo.
goto HOME

:SELECTPAST
echo.
echo 1 - Change Month
echo 2 - Change Month and Year
echo 0 - Go Back
echo.
echo Enter Selection: 
set/p "selectionpast=>"
if %selectionpast%==1 goto RUNMONTH
if %selectionpast%==2 goto RUNYEAR
if %selectionpast%==0 goto HOME



:RUNYEAR
echo Please enter the month of the invoices
set/p "choicemonth=>"
echo Please enter the year of the invoices
set/p "choiceyear=>"
"C:\Toolkits\anaconda3-5.2.0\python.exe" "_Software\Divan_Bleu_Invoicing.py" --month=%choicemonth% --year=%choiceyear%
echo.
goto HOME

:RUNMONTH
echo Please enter the number for the month of the invoices
set/p "choicemonth=>"
"C:\Toolkits\anaconda3-5.2.0\python.exe" "_Software\Divan_Bleu_Invoicing.py" --month=%choicemonth%
echo.
goto HOME

:END