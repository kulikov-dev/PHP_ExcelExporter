<?php
namespace kulikovdev;
include 'Autoloader.php';

use kulikovdev\ConverterService as ConverterService;
use kulikovdev\ExporterService as ExporterService;

class ConverterService {
    /**
     * Config: path from site root folder to folder for saving exported files
     */
    const settingsUrlExportPath = 'TempExport';
    /**
     * Config: relative path from service folder to folder for saving exported files
     */
    const settingsExportPath = "../TempExport/";

    /**
     * Copy csv file to another place;
     * $inputFilePath - relative path to csv file;
     */
    private static function CopyCsv($inputFilePath) {
        $fileName = uniqid() . ".csv";
        $outputFilePath = ConverterService::settingsExportPath . $fileName;
        copy($inputFilePath, $outputFilePath);

        $context = stream_context_create();
        $orig_file = fopen($outputFilePath, 'r', 1, $context);
        $temp_filename = tempnam(sys_get_temp_dir(), 'php_prepend_');
        file_put_contents($temp_filename, chr(0xEF).chr(0xBB).chr(0xBF));		// for unicode supporting
        file_put_contents($temp_filename, $orig_file, FILE_APPEND);
        fclose($orig_file);
        unlink($outputFilePath);
        rename($temp_filename, $outputFilePath);
        return $fileName;
    }
    /**
     * Convert csv file to XLSX format;
     * $inputFilePath - relative path to csv file;
     */
    private static function ConvertToXlsx($inputFilePath) {
        $handle = fopen($inputFilePath, "r");
        $fileName = uniqid() . ".xlsx";
        $filePath = ConverterService::settingsExportPath . $fileName;
        $writer = new \XLSXWriter();

        while ( ($data = fgetcsv($handle,0,';') ) !== FALSE ) {
            $writer->writeSheetRow('data', $data);
        }
        $writer->writeToFile($filePath);
        fclose($handle);
        return $fileName;
    }
    /**
     * Convert csv file to XLS format;
     * $inputFilePath - relative path to csv file;
     */
    private static function ConvertToXls($inputFilePath) {
        $handle = fopen($inputFilePath, "r");
        $fileName = uniqid() . ".xls";
        $filePath = ConverterService::settingsExportPath . $fileName;
        $workbook = &new \writeexcel_workbookbig($filePath);
        $worksheet = &$workbook->addworksheet();
        $lineCount = 0;
        while ( ($data = fgetcsv($handle,0,';') ) !== FALSE ) {
            $array = array_values($data);
            $subLength = count($array);
            for ($j = 0; $j < $subLength; $j++) {
                $worksheet->write($lineCount,$j, (string)$array[$j]);
            }
            ++$lineCount;
        }

        $workbook->close();
        fclose($handle);
        return $fileName;
    }

    /**
     * Convert csv file to Excel file;
     * $inputFilePath - relative path to csv file;
     * $exportType - output file format
     */
    public static function ConvertCsvToExcel($inputFilePath, $exportType) {
        if (!file_exists($inputFilePath)) {
            echo 'File not found!';
            return;
        }

        ini_set('auto_detect_line_endings',TRUE);
        switch ($exportType) {
            case ExportTypeEnum.XLSX:
                $fileName = ConverterService::ConvertToXlsx($inputFilePath);
                break;
            case ExportTypeEnum.XLS:
                $fileName = ConverterService::ConvertToXls($inputFilePath);
                break;
            case ExportTypeEnum.CSV:
                $fileName = ConverterService::CopyCsv($inputFilePath);
                break;
        }

        ini_set('auto_detect_line_endings',FALSE);
        echo $url = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://" . $_SERVER[HTTP_HOST] . "/" . ConverterService::settingsUrlExportPath . "/" . $fileName;
    }

    /**
     * Convert xls and xlsx files to csv file;
     * $inputFilePath - relative path to Excel file;
     */
    public static function ConvertExcelToCsv($inputFilePath) {
        if (!file_exists($inputFilePath)) {
            echo 'File not found!';
            return;
        }
        $hasError = false;
        $ext = pathinfo($inputFilePath, PATHINFO_EXTENSION);
        switch ($ext) {
            case 'xls':
                if ( $xlsx = \SimpleXLS::parse($inputFilePath) ) {
                    $fileName = ExporterService::ExportToCsv($xlsx->rows());
                } else {
                    $hasError = true;
                    echo SimpleXLSX::parseError();
                }
                break;
            case 'xlsx':
                if ( $xlsx = \SimpleXLSX::parse($inputFilePath) ) {
                    $fileName = ExporterService::ExportToCsv($xlsx->rows());
                } else {
                    $hasError = true;
                    echo SimpleXLSX::parseError();
                }
                break;
        }
        if (!$hasError) {
            echo $url = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://" . $_SERVER[HTTP_HOST] . "/" . ExporterService::settingsUrlExportPath . "/" . $fileName;
        }
    }
}