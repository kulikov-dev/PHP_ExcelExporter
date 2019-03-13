<?php
namespace kulikovdev;
include 'Autoloader.php';

class ConverterService {
    /**
     * @var string Config: relative path from service folder to folder for saving exported files
     */
    private $relativeExportPath = "../TempExport/";
    /**
     * @return string Relative path from service folder to folder for saving exported files
     */
    public function getRelativeExportPath()
    {
        return $this->relativeExportPath;
    }
    /**
     * @param string $relativeExportPath Relative path from service folder to folder for saving exported files
     */
    public function setRelativeExportPath($relativeExportPath)
    {
        $this->relativeExportPath = $relativeExportPath;
    }

    /**
     * @var string Delimiter for CSV parsing (fgetCsv function argument)
     */
    private $delimiter = ',';
    /**
     * @return string Delimiter for CSV parsing
     */
    public function getDelimiter()
    {
        return $this->delimiter;
    }
    /**
     * @param string $delimiter Delimiter for CSV parsing
     */
    public function setDelimiter($delimiter)
    {
        $this->delimiter = $delimiter;
    }

    /**
     * @var string Enclosure for CSV parsing (fgetCsv function argument)
     */
    private $enclosure = '"';
    /**
     * @return string Enclosure for CSV parsing
     */
    public function getEnclosure()
    {
        return $this->enclosure;
    }
    /**
     * @param string $enclosure Enclosure for CSV parsing
     */
    public function setEnclosure($enclosure)
    {
        if ($enclosure == ''){
            $this->enclosure = chr(0);  // as we can't place empty enclosure to fgetcsv func;
        }
        else {
            $this->enclosure = $enclosure;
        }
    }


    /**
     * Convert csv file to Excel file;
     * @param $inputFilePath relative path to csv file;
     * @param $exportType output file format
     * @return string filename of created file
     */
    public function ConvertCsvToExcel($inputFilePath, $exportType) {
        if (!file_exists($inputFilePath)) {
            throw new Exception('File not found!');
        }

        ini_set('auto_detect_line_endings',TRUE);
        $fileName = "";
        switch ($exportType) {
            case ExportTypeEnum.XLSX:
                $fileName = self::ConvertCsvToXlsx($inputFilePath);
                break;
            case ExportTypeEnum.XLS:
                $fileName = self::ConvertCsvToXls($inputFilePath);
                break;
            case ExportTypeEnum.CSV:
                $fileName = self::CopyCsv($inputFilePath);
                break;
        }

        ini_set('auto_detect_line_endings',FALSE);
        return $fileName;
    }
    /**
     * Convert xls and xlsx files to csv file;
     * @param $inputFilePath relative path to Excel file;
     * @return string filename of created file
     */
    public function ConvertExcelToCsv($inputFilePath) {
        if (!file_exists($inputFilePath)) {
            throw new Exception('File not found!');
        }

        $fileName = "";
        $ext = pathinfo($inputFilePath, PATHINFO_EXTENSION);
        switch (strtolower($ext)) {
            case "xls":
                $fileName = self::ConvertXlsToCsv($inputFilePath);
                break;
            case "xlsx":
                $fileName = self::ConvertXlsxToCsv($inputFilePath);
                break;
        }

        return $fileName;
    }
    /**
     * Export json string to table file;
     * @param $jsonTable dataTable data in JSON format
     * @param $exportType output file format
     * @return string filename of created file
     */
    public function ExportJsonToFile($jsonTable, $exportType) {
        $arrayData = json_decode($jsonTable, true);
        $fileName = "";
        switch ($exportType) {
            case ExportTypeEnum.XLSX:
                $fileName = self::ExportArrayTableToXlsx($arrayData);
                break;
            case ExportTypeEnum.XLS:
                $fileName = self::ExportArrayTableToXls($arrayData);
                break;
            case ExportTypeEnum.CSV:
                $fileName = self::ExportArrayTableToCsv($arrayData);
                break;
        }
        return $fileName;
    }

    /**
     * Copy csv file to another place;
     * @param $inputFilePath - relative path to csv file;
     * @return generated filename
     */
    private function CopyCsv($inputFilePath) {
        $fileName = uniqid() . ".csv";
        $outputFilePath = $this->relativeExportPath . $fileName;
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
     * @param $inputFilePath - relative path to csv file;
     * @return generated filename
     */
    private function ConvertCsvToXlsx($inputFilePath) {
        $handle = fopen($inputFilePath, "r");
        $fileName = uniqid() . ".xlsx";
        $filePath = $this->relativeExportPath . $fileName;
        $writer = new \XLSXWriter();

        while ( ($data = fgetcsv($handle,0,$this->delimiter, $this->enclosure) ) !== FALSE ) {
            $writer->writeSheetRow('data', $data);
        }
        $writer->writeToFile($filePath);
        fclose($handle);
        return $fileName;
    }
    /**
     * Convert csv file to XLS format;
     * @param $inputFilePath - relative path to csv file;
     * @return generated filename
     */
    private function ConvertCsvToXls($inputFilePath) {
        $handle = fopen($inputFilePath, "r");
        $fileName = uniqid() . ".xls";
        $filePath = $this->relativeExportPath . $fileName;
        $workbook = new \Xls\Workbook();
        $worksheet = &$workbook->addworksheet();
        $lineCount = 0;
        while ( ($data = fgetcsv($handle,0, $this->delimiter, $this->enclosure) ) !== FALSE ) {
            $array = array_values($data);
            $subLength = count($array);
            for ($j = 0; $j < $subLength; $j++) {
                $worksheet->write($lineCount,$j, (string)$array[$j]);
            }
            ++$lineCount;
        }

        $workbook->save($filePath);
        fclose($handle);
        return $fileName;
    }

    /**
     * Convert Xls to Csv file
     * @param $inputFilePath - relative path to xls file;
     * @return generated filename
     */
    private function ConvertXlsToCsv($inputFilePath) {
        if ($xlsx = \SimpleXLS::parse($inputFilePath) ) {
            $fileName = ConverterService::ExportToCsv($xlsx->rows());
            return $fileName;
        } else {
            throw new Exception(SimpleXLSX::parseError());
        }
    }
    /**
     * Convert Xlsx to Csv file
     * @param $inputFilePath - relative path to xls file;
     * @return generated filename
     */
    private function ConvertXlsxToCsv($inputFilePath) {
        $xlsx = new \XLSXReader($inputFilePath);
        $sheets = $xlsx->getSheetNames();
        if (!empty($sheets)) {
            $values = array_values($sheets);
            $data = $xlsx->getSheetData($values[0]);
            return ConverterService::ExportToCsv($data);
        } else {
            throw new Exception("Empty file");
        }
    }

    /**
     * Export php-array to csv file;
     * @param $arrayData datatable in array
     * @return generated filename
     */
    private function ExportArrayTableToCsv($arrayData) {
        $fileName = uniqid() . ".csv";
        $filePath = $this->relativeExportPath . $fileName;
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
    private function ExportArrayTableToXls($arrayData) {
        $fileName = uniqid() . ".xls";
        $filePath = $this->relativeExportPath . $fileName;
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
    private function ExportArrayTableToXlsx($arrayData) {
        $fileName = uniqid() . ".xlsx";
        $filePath = $this->relativeExportPath . $fileName;
        $writer = new \XLSXWriter();
        $writer->writeSheet($arrayData);
        $writer->writeToFile($filePath);
        return $fileName;
    }
}

/**
 * Enumerable for type of an exported file
 */
class ExportTypeEnum
{
    const XLSX = 0;
    const XLS = 1;
    const CSV = 2;

    /**
     * Convert file index to enum value.
     */
    public static function GetTypeFromNumber($fileFormatIndex)
    {
        switch ($fileFormatIndex) {
            case 0:
                return ExportTypeEnum . XLSX;
            case 1:
                return ExportTypeEnum . XLS;
            case 2:
                return ExportTypeEnum . CSV;
        }
        return ExportTypeEnum . XLSX;
    }
}
