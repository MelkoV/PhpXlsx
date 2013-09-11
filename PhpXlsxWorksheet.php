<?php
/**
 * User: smorodin
 * Date: 10.09.13
 * Time: 16:06
 */

class PhpXlsxWorksheet {
    const STORAGE_TYPE_ARRAY = 'STORAGE_TYPE_ARRAY';
    const STORAGE_TYPE_STRING = 'STORAGE_TYPE_STRING';

    public $id;
    public $name;

    private $storageType;

    private $data;
    /** @var  PhpXlsx $xlsx */
    private $xlsx;
    private $rowsCount = 0;

    private $dimensionAStart = 'A';
    private $dimensionNStart = '1';
    private $dimensionAEnd = 'A';
    private $dimensionNEnd = '1';

    public function __construct($id, $name, $xlsx, $storageType = self::STORAGE_TYPE_ARRAY)
    {
        $this->id = $id;
        $this->name = $name;
        $this->storageType = $storageType;

        $this->xlsx = $xlsx;

        if ($storageType = self::STORAGE_TYPE_ARRAY) {
            $this->data = array();
        }
    }

    /**
     * @param array $data
     */
    public function addRow(array $data = array())
    {
        switch ($this->storageType) {
            case self::STORAGE_TYPE_ARRAY:
                $this->addDataAsArray($data);
                break;
            case self::STORAGE_TYPE_STRING:
                $this->addDataAsString($data);
                break;
            default:
                break;
        }
    }

    public function getData()
    {
        switch ($this->storageType) {
            case self::STORAGE_TYPE_ARRAY:
                return $this->getDataAsArray();
                break;
            case self::STORAGE_TYPE_STRING:
                return $this->getDataAsString();
                break;
            default:
                return '';
                break;
        }
    }

    private function getDataAsString()
    {

    }

    private function getDataAsArray()
    {
        $data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $data .= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
            <dimension ref="' . $this->dimensionAStart . $this->dimensionNStart . ':' . $this->dimensionAEnd . $this->dimensionNEnd . '"/>
            <sheetViews>
            <sheetView workbookViewId="0">
            <selection activeCell="A1" sqref="A1"/>
            </sheetView>
            </sheetViews>
            <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
            <sheetData>';

        foreach ($this->data as $i => $cell) {
            $data .= '<row r="'.$i.'" spans="1:24" x14ac:dyDescent="0.25">';
            foreach ($cell as $name => $value) {
                if ($value['shared']) {
                    $data .= '<c r="' . $name . '" t="s"><v>' . $value['value'] . '</v></c>';
                } else {
                    $data .= '<c r="' . $name . '"><v>' . $value['value'] . '</v></c>';
                }
            }
            $data .= '</row>';
        }

        $data .= '</sheetData>
            <sheetProtection formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" sort="0" autoFilter="0" pivotTables="0"/>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>';

        return $data;
    }

    private function addDataAsArray(array $data = array())
    {
        $this->rowsCount++;
        $this->dimensionNEnd = $this->rowsCount;

        $j = 'A';

        $this->data[$this->rowsCount] = array();

        foreach ($data as $row) {
            if ($row !== null && trim($row) != '') {
                $value = array('shared' => true, 'value' => '');

                if (gettype($row) == 'integer') {
                    $value['shared'] = false;
                    $value['value'] = $row;
                } else {
                    $value['value'] = $this->xlsx->addSharedString($row);
                }
                $this->data[$this->rowsCount][$j . $this->rowsCount] = $value;
                if ($j > $this->dimensionAEnd) {
                    $this->dimensionAEnd = $j;
                }
            }
            $j++;
        }
    }

    private function addDataAsString(array $data = array())
    {

    }

}