# MSExcel_macro
Automate Sending mail when there is a change in cell inside MS excel Eg. Lets say there is a cell in A coulmn in excel with "Pending" written and there 
need to send mail when status changes from Pending to Complte .And when status changes from pending to complete automatic mail will be sent to desired mail id using
outlook mail server

**
Steps to use this file**

1st create a excel file and name it sheet1.xlxs and then go to saveas and change its xlsm to enable support for macro(This is due to security reason)

Then open newly saved xlsm file and press Alt+F11 

Under Interset Tab--->Module and paste send_mail code snippet

Then go to Sheet1 Worksheet and paste sheet1.xlxs file code snippet

Dont forget to Enable Microsfot Outlook 16.0 Object Library from Tools--->Refrences
