# PHP_ExcelExporter
  Small library written specially for export huge tables to XLS, XLSX and CSV tables.
Agregate solutions for XLSX export (https://github.com/mk-j/PHP_XLSXWriter) and XLS export (https://github.com/thoroc/php_writeexcel)
  
For using include ExporterService.php and call function 
* kulikovdev\ExporterService::ExportJsonToFile($json, $fileFormat);

Also you have to setup folders for saving exported files. There are two settings inside ExporterService.php:
* settingsUrlExportPath: path to a folder for saving from the website root folder
* settingsExportPath.: relative path to a folder for saving from the library folder.
 
 
