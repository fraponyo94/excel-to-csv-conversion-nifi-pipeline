# Use Nifi to convert Excel Spreadsheets to Csv
In BigData solutions,use of data in Excel spreadsheets is common and most developers will tend to transform that data to csv when developing an ingestion pipeline thatwill enable them achieve their objective. These can be achieved using various implementation strategies as per the developer's expertise. However,java provides a way of reading and writing to Excel through **Apache's POI Java API**. 

There is an out-of-the-box Nifi processor called **ConvertExcelToCSV Processor** that uses POI to convert the binary format of Excel into a CSV file for each workbook (sheet) in the spreadsheet. However,as per the  time of writing this article,it only supports  2007 excel format(.xlsx) onwards with no backward support of Excel 2003(.xls) format.

This "excel-to-csv-conversion-nifi-pipeline" have a **custom python script** that converts .xls excel spreadsheets to csv called in the pipeline via **ExecuteStreamCommand processor**.

# Technologies
  - Apache Nifi
  - Python
      - Libraries used
         - xlrd
         - unicodecsv
         - uuid
         - glob
        
# Usage 
- Load the nifi template called **excel_to_csv_nifi_template.xml** into nifi workspace
- Create the following variables in your Nifi workspace
    - **input.dir** path to location of your excel files
    - **output.dir** path to location where converted csv files will be outputed
    - **working.dir** path to where your script file is located
- Move **xls_to_csv_nifi.py** script to **working.dir** directory 
- Start your processor

# Queries 
Incase of any query concerning the pipeline,you can get in touch via **fraponyo94@gmail.com
