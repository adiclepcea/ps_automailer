# ps_automailer
powershell script to send mails from an excel file using [mailsender](https://github.com/adiclepcea/mailsender)

##Usage

This is a script that is intended to be run once a day (for example in Scheduled Tasks).
This script should have at it's disposal am excel file that will be read and the columns passed in should have the data needed for the script.

This scripts needs a column will send mails to suppliers if the contracts expire after the specified number of days.

From the powershell console you could run it with:

```
.\AutoMailer.ps1 -excelFile 'excel file.xlsx' -sheetName Sheet1 -dateLocation 'expiration date' -nameLocation SupplierName -mailLocation 'Contact 1' -mail2Location 'Contact 2' -daysBefore 40
```


