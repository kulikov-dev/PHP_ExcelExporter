<?php
namespace kulikovdev;
use kulikovdev\ExporterService as ExporterService;
include 'Autoloader.php';

class ExporterService {
    /**
     * Config: path from site root folder to folder for saving exported files
     */
    const settingsUrlExportPath = 'TempExport';
    /**
     * Config: relative path from service folder to folder for saving exported files
     */
    const settingsExportPath = "../TempExport/";

    /**
     * Export php-array to csv file;
     * @param $arrayData datatable in array
     * @return generated filename
     */
    public static function ExportToCsv($arrayData) {
        $fileName = uniqid() . ".csv";
        $filePath = ExporterService::settingsExportPath . $fileName;
        $delimiter = ';';
        $temp_memory = fopen($filePath, 'w');
        fprintf($temp_memory, chr(0xEF).chr(0xBB).chr(0xBF));
        foreach ($arrayData as $line) {
            fputcsv($temp_memory, $line, $delimiter);
        }

        fclose($temp_memory);
        return $fileName;
    }
    /**
     * Export php-array to xls file;
     * @param $arrayData datatable in array
     * @return generated filename
     */
    private static function ExportToXls($arrayData) {
        $fileName = uniqid() . ".xls";
        $filePath = ExporterService::settingsExportPath . $fileName;
        $workbook = new \Xls\Workbook();
        $worksheet = &$workbook->addworksheet();

        $length = count($arrayData);
        for ($i = 0; $i < $length; $i++) {
            $subLength = count($arrayData[$i]);
            $subArray = $arrayData[$i];
            for ($j = 0; $j < $subLength; $j++) {
                $array = array_values($subArray);
                $worksheet->write($i,$j, (string)$array[$j]);
            }
        }

        $workbook->save($filePath);
        return $fileName;
    }
    /**
     * Export php-array to xlsx file;
     * @param $arrayData datatable in array
     * @return generated filename
     */
    private static function ExportToXlsx($arrayData) {
        $fileName = uniqid() . ".xlsx";
        $filePath = ExporterService::settingsExportPath . $fileName;
        $writer = new \XLSXWriter();
        $writer->writeSheet($arrayData);
        $writer->writeToFile($filePath);
        return $fileName;
    }

    /**
     * Export json string to table file;
     * @param $jsonTable dataTable data in JSON format
     * @param $exportType output file format
     */
    public static function ExportJsonToFile($jsonTable, $exportType) {
        $arrayData = json_decode($jsonTable, true);
        switch ($exportType) {
            case ExportTypeEnum.XLSX:
                $fileName = ExporterService::ExportToXlsx($arrayData);
                break;
            case ExportTypeEnum.XLS:
                $fileName = ExporterService::ExportToXls($arrayData);
                break;
            case ExportTypeEnum.CSV:
                $fileName = ExporterService::ExportToCsv($arrayData);
                break;
        }
        echo $url = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://" . $_SERVER[HTTP_HOST] . "/" . ExporterService::settingsUrlExportPath . "/" . $fileName;
    }
}
?>