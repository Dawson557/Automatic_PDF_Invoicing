
@echo off
title Divan Bleu Invoicing System
echo.
echo Make sure you've downloaded the latest ClientRevenueReportList.xlsx from GoRendezVous and saved in Invoices folder
echo.
pause

:HOME
echo.
echo Main Menu
echo 1 - Create and E-mail commission invoices
echo 2 - Create and E-mail rent invoices
echo 3 - Create commission invoices
echo 4 - Create rent invoices
echo 0 - Exit program
echo.
echo Enter Selection: 
set/p "choice=>"
if %choice%==1 goto RUN_EMAIL
if %choice%==2 goto RUN_RENT_EMAIL
if %choice%==3 goto RUN
if %choice%==4 goto RUN_RENT
if %choice%==0 goto END
goto HOME
pause


:RUN
echo Please enter the month of the invoices
set/p "choicemonth=>"
echo Please enter the year of the invoices
set/p "choiceyear=>"
"C:\Users\bimon\AppData\Local\Programs\Python\Python38\python.exe" "_Software\Divan_Bleu_Invoicing.py" --month=%choicemonth% --year=%choiceyear%
echo.
goto HOME

:RUN_EMAIL
echo Please enter the month of the invoices
set/p "choicemonth=>"
echo Please enter the year of the invoices
set/p "choiceyear=>"
"C:\Users\bimon\AppData\Local\Programs\Python\Python38\python.exe" "_Software\Divan_Bleu_Invoicing.py" --email --month=%choicemonth% --year=%choiceyear%
echo.
goto HOME

:RUN_RENT
echo Please enter the month of the invoices
set/p "choicemonth=>"
echo Please enter the year of the invoices
set/p "choiceyear=>"
"C:\Users\bimon\AppData\Local\Programs\Python\Python38\python.exe" "_Software\Divan_Bleu_Invoicing.py" --rent --month=%choicemonth% --year=%choiceyear%
echo.
goto HOME


:RUN_RENT_EMAIL
echo Please enter the month of the invoices
set/p "choicemonth=>"
echo Please enter the year of the invoices
set/p "choiceyear=>"
"C:\Users\bimon\AppData\Local\Programs\Python\Python38\python.exe" "_Software\Divan_Bleu_Invoicing.py" --rent --email --month=%choicemonth% --year=%choiceyear%
echo.
goto HOME
:END