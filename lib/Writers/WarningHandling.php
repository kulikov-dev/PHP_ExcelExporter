<?php

class WarningInfo {
    const LongCellValue = 'The cell value too long.';

    protected $row=0;
    protected $col=0;

    public $fileName = '';
    protected $message = '';

    /**
     * XLSXWriter_Warning constructor.
     * @param $rowIndex Row index
     * @param $colIndex Column index
     * @param $infoMessage Information about warning
     */
    public function __construct($rowIndex, $colIndex, $infoMessage) {
        $this->row = $rowIndex;
        $this->col = $colIndex;
        $this->message = $infoMessage;
    }

    private function getCell($row_number, $column_number) {
        $n = $column_number;
        for($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n%26 + 0x41) . $r;
        }
        return $r . ($row_number+1);
    }
    /**
     * @return string generate readable information about warning
     */
    public function ToString() {
        return "Cell: " . $this->getCell($this->row, $this->col) . ". Warning: " . $this->message;
    }
}