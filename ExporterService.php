<?php
namespace kulikovdev;
use kulikovdev\ExporterService as ExporterService;
require_once 'lib/xlsxwriter.php';
include 'Autoloader.php';

class ExporterService {

	const settingsUrlExportPath = 'TempExport';
	/**
 	* Config: path to save exported files.
 	*/ 
	const settingsExportPath = "../TempExport/";

	/**
	 * Export php-array to csv file;
	 */ 
	private static function ConvertToCsv($someArray) {
		$fileName = uniqid() . ".csv";
		$filePath = ExporterService::settingsExportPath . $fileName;
   		$delimiter = ';';
		$temp_memory = fopen($filePath, 'w');
	    fprintf($temp_memory, chr(0xEF).chr(0xBB).chr(0xBF));
		foreach ($someArray as $line) {
			fputcsv($temp_memory, $line, $delimiter);
		}

		fclose($temp_memory);
		return $fileName;
	}
	/**
	 * Export php-array to xls file;
	 */ 
	private static function ConvertToXls($someArray) {	
		$fileName = uniqid() . ".xls";
		$filePath = ExporterService::settingsExportPath . $fileName;
		$workbook = &new \writeexcel_workbookbig($filePath);
		$worksheet = &$workbook->addworksheet();

		$length = count($someArray);
		for ($i = 0; $i < $length; $i++) {
			$subLength = count($someArray[$i]);
			$subArray = $someArray[$i];
			for ($j = 0; $j < $subLength; $j++) {
				$array = array_values($subArray);
				$worksheet->write($i,$j, (string)$array[$j]);
			}
		}
	
		$workbook->close();
		return $fileName;
	}
	/**
	 * Export php-array to xlsx file;
	 */ 
	private static function ConvertToXlsx($someArray) {
		$fileName = uniqid() . ".xlsx";
		$filePath = ExporterService::settingsExportPath . $fileName;
		$writer = new \XLSXWriter();
		$writer->writeSheet($someArray);
		$writer->writeToFile($filePath);
		return $fileName;
	}

	/**
	 * Export json string to table file;
	 * $json - dataTable data in JSON format
	 * $exportType - output file format
	 */ 
	public static function ExportJsonToFile($json, $exportType) {
		$someArray = json_decode($json, true);
		switch ($exportType) {
    		case ExportTypeEnum.XLSX:
				$fileName = ExporterService::ConvertToXlsx($someArray);
			break;
			case ExportTypeEnum.XLS:
				$fileName = ExporterService::ConvertToXls($someArray);
			break;
			case ExportTypeEnum.CSV:
        		$fileName = ExporterService::ConvertToCsv($someArray);
        	break;
		}	
		echo $url = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://" . $_SERVER[HTTP_HOST] . "/" . ExporterService::settingsUrlExportPath . "/" . $fileName;
	}
}	
?>
