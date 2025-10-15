This a data driven testing project.
Once setup is completed (putting the files in a java project) users are good to go.
To run the script locate the FD_Calculator_Main.java and run.
This script makes use of excelUtills.java which contains all the methods and veriables. This file helps in avoiding creating methods and veriables again and again there by saving time, memory space.
When file runs it gets test data from an Excel file like: Principle, Interest rate, Tenure, Frequency and it sends this information to the web application (URL:https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI-BSB001.html?classic=true) and calculates the Amount (principal+interest)
Lastly, we make a validation of for each amount, if its a match then it is marked PASSED and if now then FAILED.
