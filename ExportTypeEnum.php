<?php
namespace kulikovdev;
/**
 * Enumerable for type of an exported file
 */ 
class ExportTypeEnum {
    const XLSX = 0;
    const XLS = 1;
    const CSV = 2;
	
/**
 * Convert file index to enum value.
 */ 
	public static function GetTypeFromNumber($fileFormatIndex) {
		switch ($fileFormatIndex) {
    		case 0:
				return ExportTypeEnum.XLSX;
			case 1:
				return ExportTypeEnum.XLS;
			case 2:
				return ExportTypeEnum.CSV;
		}
		return ExportTypeEnum.XLSX;
	}
}
?>