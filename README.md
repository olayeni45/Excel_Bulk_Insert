
# Insert Bulk Transactions to Microsoft Excel

This repository contains a UiPath library for automating the insertion of bulk transactions from a datatable to an excel file at a predetermined interval.

## Authors

- [@Olayeni Anifowose](https://github.com/olayeni45)


## Installation

* Download the library folder and publish the .xaml file to your local UiPath Orchestrator.
* Install the package as an activity in your project.
* Configure the activity with the required arguments.
    
## Arguments

To run this activity, you will need to configure the following arguments:

`Breakpoint`\
If the row count in the datatable exceeds this value, the bot breaks the datatable into smaller chunks at a predefined interval. \
**Direction:** In

`Datatable`\
The datatable containing the transactions.\
**Direction:** In

`ExcelWorkbook`\
The excel workbook's absolute path for inserting transactions.\
**Direction:** In

`IntervalPerSheet`\
The number of transactions that must be entered into each worksheet.\
**Direction:** In

`WorksheetName`\
The name of the worksheet to insert the transactions.\
**Direction:** In

`Status`\
The execution status, which can either be "Success" or an exception that the bot has thrown.\
**Direction:** Out
## Screenshots

![Activity Screenshot](https://res.cloudinary.com/edconnect/image/upload/v1663784879/Github/Excel_Bulk_Insert_iqpunp.png)

