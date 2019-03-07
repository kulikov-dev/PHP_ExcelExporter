# PHP_ExcelExporter
  Small library written specially for export/convert_from_CSV huge tables to XLS, XLSX and CSV tables. 
Agregate fastest solutions for XLSX export (https://github.com/mk-j/PHP_XLSXWriter) and XLS export (https://github.com/MAXakaWIZARD/xls-writer). And XLSX import (https://github.com/gneustaetter/XLSXReader), XLS import (https://github.com/shuchkin/simplexls).
  
For export JSON data include ExporterService.php and call function 
* kulikovdev\ExporterService::ExportJsonToFile($json, $fileFormat);

where $json is an exporting table in JSON format, $fileFormat is the type of exported file (ExportTypeEnum)

For convert CSV file to Excel file include ConverterService.php and call function
* kulikovdev\ConverterService::ConvertCsvToExcel($$inputFilePath, $fileFormat);

where $$inputFilePath is a relative path to CSV file, $fileFormat is the type of exported file (ExportTypeEnum)

Also you have to setup folders for saving exported files. There are two settings inside ExporterService.php:
* settingsUrlExportPath: path to a folder for saving from the website root folder
* settingsExportPath.: relative path to a folder for saving from the library folder.
 
 
