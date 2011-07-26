<?php

/**
 * Description of POExcelWorksheet
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
class POExcelWorksheet extends POXMLFile implements POIXMLElement
{
	/**
	 * @var integer worksheet id
	 */
	private $id;
	
	/**
	 * @var string worksheet name
	 */
	private $name;
	
	/**
	 * @var array list of worksheet rows
	 */
	private $rows = array();
	
	/**
	 * @var POExcelWorkbook excel instance
	 */
	private $workbook;
	
	/**
	 * Constructor
	 * 
	 * @param POExcel $excel excel instance
	 * @param integer $id worksheet id
	 * @param string $name worksheet name
	 */
	public function __construct(POExcelWorkbook $workbook, $id, $name)
	{
		$this->workbook = $workbook;
		$this->id = $id;
		$this->name = $name;
	}

	/**
	 * @see POIXMLElement::getAsXML()
	 */
	public function getAsXML()
	{
		$xml = '';
		
		// add sheet data
		if (count($this->rows) > 0) {
			$xml .= '<sheetData>';
			
			// add rows
			ksort($this->rows);
			foreach ($this->rows as $row) {
				$xml .= $row->getAsXML();
			}
			
			$xml .= '</sheetData>';
		} else {
			$xml .= '<sheetData />';
		}
		
		return $xml;
	}
	
	/**
	 * @see POXMLFile::getAsXMLFile()
	 */
	public function getAsXMLFile()
	{
		return $this->renderXMLTemplate('sheet', array('content' => $this->getAsXML()));
	}

	/**
	 * Return cell instance by coordinate (A1, B3, etc...)
	 * 
	 * @param string $coordinate cell coordinate
	 * @return POExcelCell cell instance
	 */
	public function getCell($coordinate)
	{
		if (preg_match('/([A-Z]+)([0-9]+)/u', $coordinate, $matches)) {
			if ($matches[2] > 0) {
				return $this->getRow($matches[2])->getCell($matches[1]);
			} else {
				// @todo error
			}
		} else {
			// @todo error
		}
	}
		
	/**
	 * Return worksheet id
	 * 
	 * @return integer
	 */
	public function getId()
	{
		return $this->id;
	}
	
	/**
	 * Return worksheet name
	 * 
	 * @return string
	 */
	public function getName()
	{
		return $this->name;
	}
	
	/**
	 * Generating and return worksheet row
	 * 
	 * @param integer $name row number
	 * @return POExcelRow row instance 
	 */
	public function getRow($name)
	{
		if (!isset($this->rows[$name])) {
			$this->rows[$name] = new POExcelRow($this, $name);
		}
		return $this->rows[$name];
	}
	
	/**
	 * Return excel instance
	 * 
	 * @return POExcelWorkbook
	 */
	public function getWorkbook()
	{
		return $this->excel;
	}
	
	/**
	 * Setting cell value
	 * 
	 * @param string $cell cell coordinates
	 * @param mixed $value cell value
	 */
	public function setCellValue($cell, $value)
	{
		if (($cell = $this->getCell($cell)) instanceof POExcelCell) {
			$cell->setValue($value);
		} else {
			// @todo error
		}
	}
}