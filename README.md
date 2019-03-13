# PHP_ExcelExporter
  Small library written specially for export/convert_from_CSV huge tables to XLS, XLSX and CSV tables. 
Agregate fastest solutions for XLSX export (https://github.com/mk-j/PHP_XLSXWriter) and XLS export (https://github.com/MAXakaWIZARD/xls-writer). And XLSX import (https://github.com/gneustaetter/XLSXReader), XLS import (https://github.com/shuchkin/simplexls).

To work with library you have to create an instance of class kulikovdev\ConverterService.
For export JSON data call function 
* ExportJsonToFile($jsonTable, $exportType)

For convert CSV file to Excel file include ConverterService.php and call function
* ConvertCsvToExcel($inputFilePath, $exportType) 

Also you can set:
* $relativeExportPath: relative path from service folder to folder for saving exported files
* $delimiter and $enclosure: arguments for fgetcsv when convert csv to Excel.
 
 
