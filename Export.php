<?php
/**
 * Description of Export
 */
class Export
{
  /**
   * nazev vysledneho xml souboru
   * @var string
   */
  public $outputFilePath;
  
  /**
   * data z formulare
   * @var array
   */
  private $data = array();
  
  /**
   * data z Excelu
   * @var array
   */
  private $excelData = array();

  /**
   * uploadovany soubor
   * @var file
   */
  private $file;
  
  /** @var \DOMDocument */
  private $DOMDocument;

  
  /**
   * listy z Excelu, ktere se maji prochazet
   * @var array
   */
  private $sheetNames = array(
    'A2',
    'A4',
    'A5',
    'B2',
    'B3',
  );
  
  /**
   * pole pro ulozeni souctu
   * @var array
   */
  private $totals = array(
    'celk_zd_a2'            => 0,
    'a4_zakl_dane1'         => 0,
    'a4_zakl_dane_snizena'  => 0,
    'a5_zakl_dane1'         => 0,
    'a5_zakl_dane_snizena'  => 0,
    'b2_zakl_dane1'         => 0,
    'b2_zakl_dane_snizena'  => 0,
    'b3_zakl_dane1'         => 0,
    'b3_zakl_dane_snizena'  => 0,
  );
  
  
  
  /**
   * @param string $outputFilePath
   */
  public function __construct($outputFilePath)
  {
    $this->outputFilePath = $outputFilePath ? $outputFilePath : 'temp/file.xml';
    $this->DOMDocument = new DOMDocument('1.0', 'UTF-8');
  }

  /**
   * export dat z Excelu do XML
   * @param file $file
   * @param array $data
   */
  public function exportFile($file, $data)
  {
    $this->file = $file;
    $this->data = $data;
    
    $this->loadPhpExcel();
    $this->createXmlDoc($this->data);
  }

  /**
   * pokusi se precist uploadovany soubor
   * finfo_file(finfo_open(FILEINFO_MIME_TYPE), $fileTmpName) mi na serveru .xls vyhodnotilo jako Word
   * @param file $file
   * @return boolean
   */
  public function readExcelFile($file)
  {
    $isExcel = FALSE;
    $types = array('Excel5', 'Excel2007');
    
    foreach ($types as $type) {
      $reader = PHPExcel_IOFactory::createReader($type);
      if ($reader->canRead($file)) {
        $isExcel = TRUE;
        break;
      }
    }

    return $isExcel;
  }
  
  /**
   * inicializace souboru a ulozeni dat z Excelu do pole
   */
  private function loadPhpExcel()
  {
    $objPHPExcel = PHPExcel_IOFactory::load($this->file);
    
    // TODO - pokud zadny list neexistuje
    foreach ($this->sheetNames as $sheetName) {
      $sheet = $objPHPExcel->getSheetByName($sheetName);
      if ($sheet == NULL) { continue; }
      
      $data = array();
      $highestColumn = $sheet->getHighestColumn();
      $highestRow = $sheet->getHighestRow();

      $header = $sheet->rangeToArray('A1:' . $highestColumn . 1);

      for ($row = 2; $row <= $highestRow; $row++) {
        $rowData = $sheet->rangeToArray('A' . $row . ':'  . $highestColumn . $row, NULL, TRUE, TRUE);
        $data[] = array_combine(array_values($header[0]), array_values($rowData[0]));
      }

      $this->excelData[$sheetName] = $data;
    }
    
    if (empty($this->excelData)) {
      throw new Exception('Soubor neobsahuje relevantní data');
    }
    
  }
  
  /**
   * vytvoreni XML souboru
   */
  private function createXmlDoc()
  {
    $xmlRoot = $this->DOMDocument->createElement('Pisemnost');
    $xmlRoot->setAttribute('nazevSW', 'EPO MF ČR');
    $xmlRoot->setAttribute('verzeSW', '38.12.1');
    
    $dphElement = $this->DOMDocument->createElement('DPHKH1');
    $dphElement->setAttribute('verzePis', '01.02');

    $sentenceD = $this->DOMDocument->createElement('VetaD');
    $sentenceD->setAttribute('dokument', 'KH1');
    $sentenceD->setAttribute('k_uladis', 'DPH');
    $sentenceD->setAttribute('mesic', $this->data['mesic']);
    $sentenceD->setAttribute('khdph_forma', 'B');
    $sentenceD->setAttribute('rok', $this->data['rok']);
    $sentenceD->setAttribute('d_poddp', $this->data['d_poddp']);
    
    $sentenceP = $this->DOMDocument->createElement('VetaP');
    $sentenceP->setAttribute('naz_obce', $this->data['naz_obce']);
    $sentenceP->setAttribute('opr_postaveni', $this->data['opr_postaveni']);
    $sentenceP->setAttribute('psc', $this->data['psc']);
    $sentenceP->setAttribute('c_pop', $this->data['c_pop']);
    $sentenceP->setAttribute('id_dats', $this->data['id_dats']);
    $sentenceP->setAttribute('zast_ic', $this->data['zast_ic']);
    $sentenceP->setAttribute('zast_kod', $this->data['zast_kod']);
    $sentenceP->setAttribute('zast_typ', $this->data['zast_typ']);
    $sentenceP->setAttribute('c_pracufo', $this->data['c_pracufo']);
    $sentenceP->setAttribute('c_ufo', $this->data['c_ufo']);
    $sentenceP->setAttribute('opr_jmeno', $this->data['opr_jmeno']);
    $sentenceP->setAttribute('stat', $this->data['stat']);
    $sentenceP->setAttribute('zkrobchjm', $this->data['zkrobchjm']);
    $sentenceP->setAttribute('zast_nazev', $this->data['zast_nazev']);
    $sentenceP->setAttribute('ulice', $this->data['ulice']);
    $sentenceP->setAttribute('typ_ds', $this->data['typ_ds']);
    $sentenceP->setAttribute('opr_prijmeni', $this->data['opr_prijmeni']);
    $sentenceP->setAttribute('c_orient', $this->data['c_orient']);
    $sentenceP->setAttribute('dic', $this->data['dic']);

    $dphElement->appendChild($sentenceD);
    $dphElement->appendChild($sentenceP);
    
    // data z Excelu
    foreach ($this->sheetNames as $sheetName) {
      $sheetData = $this->excelData[$sheetName];
      
      if ($sheetData) {
        switch ($sheetName) {
          case 'A2':
            $this->getXmlA2($sheetData, $dphElement);
            break;
          case 'A4':
            $this->getXmlA4($sheetData, $dphElement);
            break;
          case 'A5':
            $this->getXmlA5($sheetData, $dphElement);
            break;
          case 'B2':
            $this->getXmlB2($sheetData, $dphElement);
            break;
          case 'B3':
            $this->getXmlB3($sheetData, $dphElement);
            break;
        }
      }
    }
    
    // soucty
    $sentenceC = $this->DOMDocument->createElement('VetaC');
    $sentenceC->setAttribute('celk_zd_a2', $this->totals['celk_zd_a2']);
    $sentenceC->setAttribute('obrat23', $this->totals['a4_zakl_dane1'] + $this->totals['a5_zakl_dane1']);
    $sentenceC->setAttribute('pln23', $this->totals['b2_zakl_dane1'] + $this->totals['b3_zakl_dane1']);
    
    if (($this->totals['a4_zakl_dane_snizena'] + $this->totals['a5_zakl_dane_snizena']) != 0) {
      $sentenceC->setAttribute('obrat5', $this->totals['a4_zakl_dane_snizena'] + $this->totals['a5_zakl_dane_snizena']);
    }
    
    if (($this->totals['b2_zakl_dane_snizena'] + $this->totals['b3_zakl_dane_snizena']) != 0) {
      $sentenceC->setAttribute('pln5', $this->totals['b2_zakl_dane_snizena'] + $this->totals['b3_zakl_dane_snizena']);
    }
    $dphElement->appendChild($sentenceC);
    
    
    $xmlRoot->appendChild($dphElement);
    $this->DOMDocument->appendChild($xmlRoot);

    if (file_exists($this->outputFilePath)) {
      unlink($this->outputFilePath);
    }
    
    $this->DOMDocument->save($this->outputFilePath);
  }
  
  
  /**
   * ziskani XML pro A2
   * @param array $sheetData
   * @param \DOMElement $dphElement
   */
  private function getXmlA2($sheetData, $dphElement)
  {
    $sumZaklDane = 0;
    
    foreach ($sheetData as $rowId => $row) {
      if ($row['vat_type'] == NULL) { continue; }
      
      array_key_exists('zakl_dane1', $row) ? $sumZaklDane += str_replace(',', '', $row['zakl_dane1']) * -1 : '';
      array_key_exists('zakl_dane2', $row) ? $sumZaklDane += str_replace(',', '', $row['zakl_dane2']) * -1 : '';
      array_key_exists('zakl_dane3', $row) ? $sumZaklDane += str_replace(',', '', $row['zakl_dane3']) * -1 : '';
      
      $element = $this->createElementA2($rowId, $row);
      $element ? $dphElement->appendChild($element) : '';
    }
    
    $this->totals['celk_zd_a2'] = $sumZaklDane;
  }
  
  /**
   * ziskani XML pro A4
   * @param array $sheetData
   * @param \DOMElement $dphElement
   */
  private function getXmlA4($sheetData, $dphElement)
  {
    $sumZaklDane = 0;
    $sumZaklDaneSnizena = 0;
    
    foreach ($sheetData as $rowId => $row) {
      if ($row['vat_type'] == NULL) { continue; }

      array_key_exists('zakl_dane1', $row) ? $sumZaklDane += str_replace(',', '', $row['zakl_dane1']) * -1 : '';
      array_key_exists('zakl_dane2', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane2']) * -1 : '';
      array_key_exists('zakl_dane3', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane3']) * -1 : '';
      
      $element = $this->createElementA4($rowId, $row);
      $element ? $dphElement->appendChild($element) : '';
    }
    
    $this->totals['a4_zakl_dane1'] = $sumZaklDane;
    $this->totals['a4_zakl_dane_snizena'] = $sumZaklDaneSnizena;
  }

  /**
   * ziskani XML pro A5
   * @param array $sheetData
   * @param \DOMElement $dphElement
   */
  private function getXmlA5($sheetData, $dphElement)
  {
    $sumZaklDane = 0;
    $sumZaklDaneSnizena = 0;
    
    foreach ($sheetData as $row) {
      array_key_exists('zakl_dane1', $row) ? $sumZaklDane += str_replace(',', '', $row['zakl_dane1']) * -1 : '';
      array_key_exists('zakl_dane2', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane2']) * -1 : '';
      array_key_exists('zakl_dane3', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane3']) * -1 : '';
      
      $element = $this->createElementA5($row);
      $element ? $dphElement->appendChild($element) : '';
    }
    
    $this->totals['a5_zakl_dane1'] = $sumZaklDane;
    $this->totals['a5_zakl_dane_snizena'] = $sumZaklDaneSnizena;
  }

  /**
   * ziskani XML pro B2
   * @param array $sheetData
   * @param \DOMElement $dphElement
   */
  private function getXmlB2($sheetData, $dphElement)
  {
    $sumZaklDane = 0;
    $sumZaklDaneSnizena = 0;
    
    foreach ($sheetData as $rowId => $row) {
      if ($row['vat_type'] == NULL) { continue; }

      array_key_exists('zakl_dane1', $row) ? $sumZaklDane += str_replace(',', '', $row['zakl_dane1']): '';
      array_key_exists('zakl_dane2', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane2']): '';
      array_key_exists('zakl_dane3', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane3']) : '';
      
      $element = $this->createElementB2($rowId, $row);
      $element ? $dphElement->appendChild($element) : '';
    }
    
    $this->totals['b2_zakl_dane1'] = $sumZaklDane;
    $this->totals['b2_zakl_dane_snizena'] = $sumZaklDaneSnizena;
  }

  /**
   * ziskani XML pro B3
   * @param array $sheetData
   * @param \DOMElement $dphElement
   */
  private function getXmlB3($sheetData, $dphElement)
  {
    $sumZaklDane = 0;
    $sumZaklDaneSnizena = 0;
    
    foreach ($sheetData as $row) {
      if (count(array_unique($row)) != 1) {
        
        array_key_exists('zakl_dane1', $row) ? $sumZaklDane += str_replace(',', '', $row['zakl_dane1']) : '';
        array_key_exists('zakl_dane2', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane2']) : '';
        array_key_exists('zakl_dane3', $row) ? $sumZaklDaneSnizena += str_replace(',', '', $row['zakl_dane3']) : '';
        
        $element = $this->createElementB3($row);
        $element ? $dphElement->appendChild($element) : '';
      }
    }
    
    $this->totals['b3_zakl_dane1'] = $sumZaklDane;
    $this->totals['b3_zakl_dane_snizena'] = $sumZaklDaneSnizena;
  }

  
  /**
   * atributy pro element A2
   * @param int $rowId
   * @param array $row
   * @return \DOMElement
   */
  private function createElementA2($rowId, $row)
  {
    $element = $this->DOMDocument->createElement('VetaA2');
    $element->setAttribute('c_radku', ++$rowId);
    $element->setAttribute('zakl_dane1', str_replace(',', '', $row['zakl_dane1']) * -1);
    $element->setAttribute('dan1', str_replace(',', '', $row['dan1']) * -1);
    $element->setAttribute('c_evid_dd', $row['c_evid_dd']);
    $element->setAttribute('k_stat', substr($row['k_stat_vatid_dod'], 0, 2));
    $element->setAttribute('vatid_dod', substr($row['k_stat_vatid_dod'], 2));
    $element->setAttribute('dppd', $this->getDppdDate($row['dppd']));
    
    $this->setTaxAttributes($row, $element, -1);
    
    return $element;
  }
  
  /**
   * atributy pro element A4
   * @param int $rowId
   * @param array $row
   * @return \DOMElement
   */
  private function createElementA4($rowId, $row)
  {
    $element = $this->DOMDocument->createElement('VetaA4');
    $element->setAttribute('c_radku', ++$rowId);
    $element->setAttribute('zakl_dane1', str_replace(',', '', $row['zakl_dane1']) * -1);
    $element->setAttribute('dan1', str_replace(',', '', $row['dan1']) * -1);
    $element->setAttribute('c_evid_dd', $row['c_evid_dd']);
    $element->setAttribute('kod_rezim_pl', $row['kod_rezim_pl']);
    $element->setAttribute('zdph_44', $row['zdph_44']);
    $element->setAttribute('dic_odb', $row['dic_odb']);
    $element->setAttribute('dppd', $this->getDppdDate($row['dppd']));
    
    $this->setTaxAttributes($row, $element, -1);
    
    return $element;
  }

  /**
   * atributy pro element A5
   * @param array $row
   * @return \DOMElement
   */
  private function createElementA5($row)
  {
    $element = $this->DOMDocument->createElement('VetaA5');
    $element->setAttribute('dan1', str_replace(',', '', $row['dan1']) * -1);
    $element->setAttribute('zakl_dane1', str_replace(',', '', $row['zakl_dane1']) * -1);

    $this->setTaxAttributes($row, $element, -1);

    return $element;
  }
  
  /**
   * atributy pro element B2
   * @param int $rowId
   * @param array $row
   * @return \DOMElement
   */
  private function createElementB2($rowId, $row)
  {
    $element = $this->DOMDocument->createElement('VetaB2');
    $element->setAttribute('c_radku', ++$rowId);
    $element->setAttribute('dan1', str_replace(',', '', $row['dan1']));
    $element->setAttribute('zakl_dane1', str_replace(',', '', $row['zakl_dane1']));
    $element->setAttribute('c_evid_dd', $row['c_evid_dd']);
    $element->setAttribute('zdph_44', $row['zdph_44']);
    $element->setAttribute('dic_dod', substr($row['dic_dod'], 2));
    $element->setAttribute('pomer', $row['pomer']);
    $element->setAttribute('dppd', $this->getDppdDate($row['dppd']));

    $this->setTaxAttributes($row, $element, 1);
    
    return $element;
  }

  /**
   * atributy pro element B3
   * @param array $row
   * @return \DOMElement
   */
  private function createElementB3($row)
  {
    $element = $this->DOMDocument->createElement('VetaB3');
    $row['dan1'] != 0 ? $element->setAttribute('dan1', str_replace(',', '', $row['dan1'])) : '';
    $row['zakl_dane1'] != 0 ? $element->setAttribute('zakl_dane1', str_replace(',', '', $row['zakl_dane1'])) : '';
    
    $this->setTaxAttributes($row, $element, 1);
    
    return $element;
  }

  /**
   * nastavi atributy pro snizene sazby
   * @param array $row
   * @param \DOMElement $element
   * @param int $value
   */
  private function setTaxAttributes($row, $element, $value)
  {
    (array_key_exists('dan2', $row) && $row['dan2'] != 0) ? $element->setAttribute('dan2', str_replace(',', '', $row['dan2']) * $value) : '';
    (array_key_exists('zakl_dane2', $row) && $row['zakl_dane2'] != 0) ? $element->setAttribute('zakl_dane2', str_replace(',', '', $row['zakl_dane2']) * $value) : '';
    (array_key_exists('dan3', $row) && $row['dan3'] != 0) ? $element->setAttribute('dan3', str_replace(',', '', $row['dan3']) * $value) : '';
    (array_key_exists('zakl_dane3', $row) && $row['zakl_dane3'] != 0) ? $element->setAttribute('zakl_dane3', str_replace(',', '', $row['zakl_dane3']) * $value) : '';
  }
  
  /**
   * format data
   * @param string $dppd
   * @return string
   */
  private function getDppdDate($dppd)
  {
    if ($this->data['dppd_format'] == '1') {
      return str_replace('/', '.', $dppd);
    } else {
      $date = DateTime::createFromFormat('m/d/Y', $dppd);
      return $date instanceof \DateTime ? $date->format('d.m.Y') : str_replace('/', '.', $dppd);
    }
  }

  /**
   * cesta k vystupnimu souboru
   * @return string
   */
  public function getOutputFilePath()
  {
    return $this->outputFilePath;
  }
  
}
