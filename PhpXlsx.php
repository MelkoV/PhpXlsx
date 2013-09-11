<?php
/**
 * Генерирует xlsx файлы.
 * Умеет работать с листами.
 * Умеет только добавлять строчные данные в сетку таблицы построчно массивом ячеек.
 *
 * User: smorodin
 * Date: 10.09.13
 * Time: 15:58
 */

class PhpXlsx extends ZipArchive {

    public $title = 'Xlsx file';
    public $subject = 'Xlsx file';
    public $creator = 'PhpXlsx';
    public $lastModifiedBy = 'PhpXlsx';

    private $application = 'Microsoft Excel';
    private $docSecurity = 0;
    private $scaleCrop = 'false';
    /** @var PhpXlsxWorksheet[] $worksheets */
    private $worksheets = array();
    private $company = 'Microsoft Corporation';
    private $linksUpToDate = 'false';
    private $sharedDoc = 'false';
    private $hyperlinksChanged = 'false';
    private $appVersion = '14.0300';
    private $worksheetsCount = 0;
    private $templateDir;

    private $sstCount = 0;
    private $sstUniqueCount = 0;
    private $sharedStrings = array();

    /**
     * @param null $templateDir
     */
    public function __construct($templateDir = null)
    {
        if (!$templateDir) {
            $this->templateDir = dirname(__FILE__) . '/template/';
        } else {
            $this->templateDir = $templateDir;
        }
    }

    /**
     * @param $string
     * @return int
     */
    public function addSharedString($string)
    {
        $string = htmlspecialchars($string);
        $this->sstCount++;
        /*$keys = array_keys($this->sharedStrings, $string);
        if ($keys && count($keys) == 1) {
            return $keys[0];
        }*/

        $this->sharedStrings[] = $string;
        return $this->sstUniqueCount++;
    }

    /**
     * @param $name
     * @param $storageType
     * @return PhpXlsxWorksheet
     */
    public function addWorksheet($name, $storageType = PhpXlsxWorksheet::STORAGE_TYPE_ARRAY)
    {
        $this->worksheetsCount++;
        $worksheet = new PhpXlsxWorksheet($this->worksheetsCount, $name, $this, $storageType);
        $this->worksheets[$this->worksheetsCount] = $worksheet;
        return $worksheet;
    }

    /**
     * @param $id
     * @return null|PhpXlsxWorksheet
     */
    public function getWorksheet($id)
    {
        if (isset($this->worksheets[$id])) {
            return $this->worksheets[$id];
        }
        return null;
    }

    public function save($fileName)
    {
        /** Создаем файл для архива */
        if ($this->open($fileName, ZipArchive::CREATE) !== true) {
            return false;
        }

        $this->addFile($this->templateDir . '_rels/.rels', '_rels/.rels');
        $this->addContentTypes();
        $this->addDocPropsApp();
        $this->addDocPropsCore();
        $this->addXlRels();
        $this->addXlTheme();
        $this->addXlWorksheets();
        $this->addXlStyles();
        $this->addXlWorkbook();
        $this->addXlSharedStrings();

        $this->close();
        return true;
    }

    private function addContentTypes()
    {
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $content .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $content .= '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $content .= '<Default Extension="xml" ContentType="application/xml"/>';
        $content .= '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
        $content .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        foreach ($this->worksheets as $worksheet) {
            $content .= '<Override PartName="/xl/worksheets/sheet' . $worksheet->id . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $content .= '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
        $content .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $content .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
        $content .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $content .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $content .= '</Types>';

        $this->addFromString('[Content_Types].xml', $content);
    }

    private function addDocPropsApp()
    {
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $content .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
        $content .= '<Application>' . $this->application . '</Application>';
        $content .= '<DocSecurity>' . $this->docSecurity . '</DocSecurity>';
        $content .= '<ScaleCrop>' . $this->scaleCrop . '</ScaleCrop>';
        $content .= '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Листы</vt:lpstr></vt:variant><vt:variant><vt:i4>' . $this->worksheetsCount . '</vt:i4></vt:variant></vt:vector></HeadingPairs>';
        $content .= '<TitlesOfParts><vt:vector size="' . $this->worksheetsCount . '" baseType="lpstr">';
        foreach ($this->worksheets as $worksheet) {
            $content .= '<vt:lpstr>' . $worksheet->name . '</vt:lpstr>';
        }
        $content .= '</vt:vector></TitlesOfParts>';
        $content .= '<Company>' . $this->company . '</Company>';
        $content .= '<LinksUpToDate>' . $this->linksUpToDate . '</LinksUpToDate>';
        $content .= '<SharedDoc>' . $this->sharedDoc . '</SharedDoc>';
        $content .= '<HyperlinksChanged>' . $this->hyperlinksChanged . '</HyperlinksChanged>';
        $content .= '<AppVersion>' . $this->appVersion . '</AppVersion>';
        $content .= '</Properties>';

        $this->addFromString('docProps/app.xml', $content);
    }

    private function addDocPropsCore()
    {
        $date = date('Y-m-d') . 'T' . date('H:i:s') . 'Z';

        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $content .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        $content .= '<dc:title>' . $this->title . '</dc:title>';
        $content .= '<dc:subject>' . $this->subject . '</dc:subject>';
        $content .= '<dc:creator>' . $this->creator . '</dc:creator>';
        $content .= '<cp:lastModifiedBy>' . $this->lastModifiedBy . '</cp:lastModifiedBy>';
        $content .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . $date . '</dcterms:created>';
        $content .= '<dcterms:modified xsi:type="dcterms:W3CDTF">' . $date . '</dcterms:modified>';
        $content .= '</cp:coreProperties>';

        $this->addFromString('docProps/core.xml', $content);
    }

    private function addXlRels()
    {
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $content .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $rels = array(
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        );
        $i = 1;
        foreach ($this->worksheets as $worksheet) {
            $content .= '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $worksheet->id . '.xml"/>';
            $i++;
        }
        foreach ($rels as $rel) {
            $content .= '<Relationship Id="rId' . $i . '" ' . $rel;
            $i++;
        }
        $content .= '</Relationships>';

        $this->addFromString('xl/_rels/workbook.xml.rels', $content);
    }

    private function addXlTheme()
    {
        $this->addFile($this->templateDir . 'xl/theme/theme1.xml', 'xl/theme/theme1.xml');
    }

    private function addXlWorksheets()
    {
        foreach ($this->worksheets as $worksheet) {
            $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
            $this->addFromString('xl/worksheets/_rels/sheet' . $worksheet->id . '.xml.rels', $content);
            $this->addXlWorksheet($worksheet);
        }
    }

    private function addXlWorksheet(PhpXlsxWorksheet $worksheet)
    {
        $this->addFromString('xl/worksheets/sheet' . $worksheet->id . '.xml', $worksheet->getData());
    }

    private function addXlStyles()
    {
        $this->addFile($this->templateDir . 'xl/styles.xml', 'xl/styles.xml');
    }

    private function addXlWorkbook()
    {
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $content .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $content .= '<fileVersion appName="xl" lastEdited="2" lowestEdited="1" rupBuild="9302"/>';
        $content .= '<workbookPr codeName="ThisWorkbook"/>';
        $content .= '<bookViews>';
        $content .= '<workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>';
        $content .= '</bookViews>';
        $content .= '<sheets>';
        $i = 1;
        foreach ($this->worksheets as $worksheet) {
            $content .= '<sheet name="' . $worksheet->name . '" sheetId="' . $worksheet->id . '" r:id="rId' . $i . '"/>';
            $i++;
        }

        $content .= '</sheets>';
        $content .= '<calcPr calcId="124519"/>';
        $content .= '</workbook>';

        $this->addFromString('xl/workbook.xml', $content);
    }

    private function addXlSharedStrings()
    {
        $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        $content .= '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $this->sstCount . '" uniqueCount="' . $this->sstUniqueCount . '">';
        foreach ($this->sharedStrings as $string) {
            $content .= '<si><t>' . $string . '</t></si>';
        }
        $content .= '</sst>';

        $this->addFromString('xl/sharedStrings.xml', $content);
    }

}