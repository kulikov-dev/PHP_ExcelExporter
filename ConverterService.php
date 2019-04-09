<?php
namespace kulikovdev;
use Exception;

include 'Autoloader.php';

/**
 * Class ConverterService
 * Provides functions for converting from csv to Excel xls/xlsx formats and vice versa
 * @package kulikovdev
 */
class ConverterService {

    private $bomString;
    function __construct() {
        $this->bomString = chr(0xEF).chr(0xBB).chr(0xBF);
    }

    /**
     * @var string Config: relative path from service folder to folder for saving exported files
     */
    private $relativeExportPath = "../reports/";
    /**
     * @return string Relative path from service folder to folder for saving exported files
     */
    public function getRelativeExportPath() {
        return $this->relativeExportPath;
    }
    /**
     * @param string $relativeExportPath Relative path from service folder to folder for saving exported files
     */
    public function setRelativeExportPath($relativeExportPath) {
        $this->relativeExportPath = $relativeExportPath;
    }

    /**
     * @var string Delimiter for CSV parsing (fgetCsv function argument)
     */
    private $delimiter = ',';
    /**
     * @return string Delimiter for CSV parsing
     */
    public function getDelimiter() {
        return $this->delimiter;
    }
    /**
     * @param string $delimiter Delimiter for CSV parsing
     */
    public function setDelimiter($delimiter) {
        $this->delimiter = $delimiter;
    }

    /**
     * @var string Enclosure for CSV parsing (fgetCsv function argument)
     */
    private $enclosure = '"';
    /**
     * @return string Enclosure for CSV parsing
     */
    public function getEnclosure() {
        return $this->enclosure;
    }
    /**
     * @param string $enclosure Enclosure for CSV parsing
     */
    public function setEnclosure($enclosure) {
        if ($enclosure == ''){
            $this->enclosure = chr(0);  // as we can't place empty enclosure to fgetcsv func;
        }
        else {
            $this->enclosure = $enclosure;
        }
    }


    const defaultCsvEncoding = "Windows-1252";
    /**
     * @var string Encoding for input CSV file in CSV to Excel converting
     */
    private $csvEncoding = defaultCsvEncoding;
    /**
     * @return string Encoding for input CSV file in CSV to Excel converting
     */
    public function getCsvEncoding() {
        return $this->csvEncoding;
    }
    /**
     * @param string $encoding Encoding for input CSV file in CSV to Excel converting
     */
    public function setCsvEncoding($encoding) {
        if ($encoding == ''){
            $this->csvEncoding = defaultCsvEncoding;
        }
        else {
            $this->csvEncoding = $encoding;
        }
    }

    private $warningsHandle = WarningHandlingEnum::TryToFix;
    /**
     * @return int Method to work with warnings during converting
     */
    public function getWarningsHandle()
    {
        return $this->warningsHandle;
    }
    /**
     * @param int $warningsHandle  Method to work with warnings during converting
     */
    public function setWarningsHandle($warningsHandle)
    {
        $this->warningsHandle = $warningsHandle;
    }

    /**
     * @var string Config: relative path from service folder to folder for saving warnings files
     */
    private $relativeWarningsExportPath = "../reports/";
    /**
     * @return string Relative path from service folder to folder for saving exported warnings files
     */
    public function getRelativeWarningsExportPath() {
        return $this->relativeWarningsExportPath;
    }
    /**
     * @param string $relativeWarningsExportPath Relative path from service folder to folder for saving warnings files
     */
    public function setRelativeWarningsExportPath($relativeWarningsExportPath) {
        $this->relativeWarningsExportPath = $relativeWarningsExportPath;
    }

    private $warnings = array();

    /** warnings. if empty - no warnings.
     * @return array warnings, occured during converting.
     */
    public function GetWarnings() {
        $stringWarnings = array();
        foreach ($this->warnings as $warning) {
           $stringWarnings[] = $warning->ToString();
        }
        return $stringWarnings;
    }

    /**
     * Convert csv file to Excel file;
     * @param string $inputFilePath Relative path to csv file;
     * @param ExportTypeEnum $exportType Output file format
     * @return string Filename of created file
     * @throws Exception File not found exception
     * @throws Exception wrong handle method or didn't setted up relativeWarningExportPath
     */
    public function ConvertCsvToExcel($inputFilePath, $exportType) {
        if (!file_exists($inputFilePath)) {
            throw new Exception('File not found!');
        }

        if (($this->warningsHandle ==  WarningHandlingEnum::WriteToFile) and ($this->relativeWarningsExportPath == '')){
            throw new Exception("Change warning handle method or set up relativeWarningExportPath.");
        }

        ini_set('auto_detect_line_endings',TRUE);
        $fileName = "";
        switch ($exportType) {
            case ExportTypeEnum::XLSX:
                $fileName = self::ConvertCsvToXlsx($inputFilePath);
                break;
            case ExportTypeEnum::XLS:
                $fileName = self::ConvertCsvToXls($inputFilePath);
                break;
            case ExportTypeEnum::CSV:
                $fileName = self::CopyCsv($inputFilePath);
                break;
        }

        ini_set('auto_detect_line_endings',FALSE);
        return $fileName;
    }
    /**
    * Convert xls and xlsx files to csv file;
    * @param string $inputFilePath Relative path to Excel file;
    * @param string $customInputFileExtension Custom extention for the input file.
    * @return string Filename of created file
    * @throws Exception File not found exception
    */
    public function ConvertExcelToCsv($inputFilePath, $customInputFileExtension = '') {
        if (!file_exists($inputFilePath)) {
            throw new Exception('File not found!');
        }
        $fileName = "";
        $ext = $customInputFileExtension;
        if (empty($customInputFileExtension))
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
     * @param mixed $jsonTable DataTable data in JSON format
     * @param ExportTypeEnum $exportType Output file format
     * @return string Filename of created file
     */
    public function ExportJsonToFile($jsonTable, $exportType) {
        $arrayData = json_decode($jsonTable, true);
        $fileName = "";
        switch ($exportType) {
            case ExportTypeEnum::XLSX:
                $fileName = self::ExportArrayTableToXlsx($arrayData);
                break;
            case ExportTypeEnum::XLS:
                $fileName = self::ExportArrayTableToXls($arrayData);
                break;
            case ExportTypeEnum::CSV:
                $fileName = self::ExportArrayTableToCsv($arrayData);
                break;
        }
        return $fileName;
    }

    /** As fgetcsv depends on locals, we have to update encoding.
     * @param $str Input string
     * @return string Encoded string
     */
    private function ConvertEncoding($str ) {
        $utf8 = 'UTF-8';
        $user_encoding = $this->getCsvEncoding();
        if ((strcasecmp($user_encoding, $utf8) == 0) or (mb_check_encoding($str, $utf8))) {
            return $str;
        }

        return iconv( $user_encoding, $utf8, $str );
    }
    /**
     * Copy csv file to another place;
     * @param string $inputFilePath Relative path to csv file;
     * @return string Filename of created file
     */
    private function CopyCsv($inputFilePath) {
        $fileName = uniqid() . ".csv";
        $outputFilePath = $this->relativeExportPath . $fileName;
        copy($inputFilePath, $outputFilePath);

        $context = stream_context_create();
        $orig_file = fopen($outputFilePath, 'r', 1, $context);
        $temp_filename = tempnam(sys_get_temp_dir(), 'php_prepend_');
        file_put_contents($temp_filename, $this->bomString);		// for unicode supporting
        file_put_contents($temp_filename, $orig_file, FILE_APPEND);
        fclose($orig_file);
        unlink($outputFilePath);
        rename($temp_filename, $outputFilePath);
        return $fileName;
    }
    /**
     * If we move problem values to separate files - write the links to the manifest file
     */
    private function WriteWarningsManifest() {
        if ($this->warningsHandle == WarningHandlingEnum::WriteToFile) {
            $manifestHandle = fopen($this->relativeWarningsExportPath . "manifest.txt","wb");
            foreach ($this->warnings  as $warning) {
                if ($warning->fileName != '') {
                    fwrite($manifestHandle, $warning->fileName . PHP_EOL);
                }
            }
            fclose($manifestHandle);
        }
    }

    /**
     * Convert csv file to XLSX format;
     * @param string $inputFilePath Relative path to csv file;
     * @return string Filename of created file
     */
    private function ConvertCsvToXlsx($inputFilePath) {
        $handle = fopen($inputFilePath, "r");
        $fileName = uniqid() . ".xlsx";
        $filePath = $this->relativeExportPath . $fileName;
        $writer = new \XLSXWriter();
        $writer->warningHandling = $this->warningsHandle;
        $writer->waningsOutputPath = $this->relativeWarningsExportPath;

        while ( ($data = fgetcsv($handle,0,$this->getDelimiter(), $this->getEnclosure()) ) !== FALSE ) {
            $row = array_map(array($this,"ConvertEncoding"), $data );
            $writer->writeSheetRow('data', $row);
        }
        $writer->writeToFile($filePath);
        $this->warnings = $writer->getWarnings();
        fclose($handle);

        $this->WriteWarningsManifest();
        return $fileName;
    }
    /**
     * Convert csv file to XLS format;
     * @param string $inputFilePath Relative path to csv file;
     * @return string Filename of created file
     */
    private function ConvertCsvToXls($inputFilePath) {
        $handle = fopen($inputFilePath, "r");
        $fileName = uniqid() . ".xls";
        $filePath = $this->relativeExportPath . $fileName;
        $workbook = new \Xls\Workbook();
        $workbook->warningHandling = $this->warningsHandle;
        $workbook->waningsOutputPath = $this->relativeWarningsExportPath;
        $worksheet = &$workbook->addworksheet();
        $lineCount = 0;
        while ( ($data = fgetcsv($handle,0, $this->getDelimiter(), $this->getEnclosure()) ) !== FALSE ) {
            $row = array_map(array($this,"ConvertEncoding"), $data );
            $array = array_values($row);
            $subLength = count($array);
            for ($j = 0; $j < $subLength; $j++) {
                $worksheet->write($lineCount,$j, (string)$array[$j]);
            }
            ++$lineCount;
        }

        $workbook->save($filePath);
        $this->warnings = $workbook->getWarnings();
        fclose($handle);

        $this->WriteWarningsManifest();
        return $fileName;
    }

    /**
     * Convert Xls to Csv file
     * @param string $inputFilePath Relative path to xls file;
     * @return string Filename of created file
     * @throws Exception Incorrect XLS file
     */
    private function ConvertXlsToCsv($inputFilePath) {
        if ($xlsx = \SimpleXLS::parse($inputFilePath) ) {
            $fileName = self::ExportArrayTableToCsv($xlsx->rows());
            return $fileName;
        } else {
            throw new Exception(SimpleXLSX::parseError());
        }
    }
    /**
     * Convert Xlsx to Csv file
     * @param string $inputFilePath Relative path to xls file;
     * @return string Filename of created file
     * @throws Exception Incorrect XLSX file
     */
    private function ConvertXlsxToCsv($inputFilePath) {
        $xlsx = new \XLSXReader($inputFilePath);
        $sheets = $xlsx->getSheetNames();
        if (!empty($sheets)) {
            $values = array_values($sheets);
            $data = $xlsx->getSheetData($values[0]);
            return self::ExportArrayTableToCsv($data);
        } else {
            throw new Exception("Empty file");
        }
    }

    /**
     * Export php-array to csv file;
     * @param mixed $arrayData Datatable in array
     * @return string Filename of created file
     */
    private function ExportArrayTableToCsv($arrayData) {
        $fileName = uniqid() . ".csv";
        $filePath = $this->relativeExportPath . $fileName;
        $delimiter = ';';
        $temp_memory = fopen($filePath, 'w');
        fprintf($temp_memory, $this->bomString);
        foreach ($arrayData as $line) {
            fputcsv($temp_memory, $line, $delimiter);
        }

        $stat = fstat($temp_memory);
        ftruncate($temp_memory, $stat['size']-1);

        fclose($temp_memory);
        return $fileName;
    }
    /**
     * Export php-array to xls file;
     * @param mixed $arrayData Datatable in array
     * @return string Filename of created file
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
     * @param mixed $arrayData Datatable in array
     * @return string Filename of created file
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
    public static function GetTypeFromNumber($fileFormatIndex) {
        switch ($fileFormatIndex) {
            case 0:
                return ExportTypeEnum::XLSX;
            case 1:
                return ExportTypeEnum::XLS;
            case 2:
                return ExportTypeEnum::CSV;
        }
        return ExportTypeEnum::XLSX;
    }
}

/**
 * Enumeration for warnings handle during converting.
 * TryToFix - trying to fix problems in place. Can cause data loss (for example in case of long cell value it will be cutted)
 * WriteToFix - moving problem cell values to separate files.
 */
class WarningHandlingEnum
{
    const TryToFix = 0;
    const WriteToFile = 1;
}