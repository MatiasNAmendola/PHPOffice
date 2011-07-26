<?php

/**
 * Description of POExcelRow
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
class POExcelRow extends POComponent implements POIXMLElement
{
	/**
	 * @var array list of row cells
	 */
	private $cells = array();
	
	/**
	 * @var integer row number
	 */
	private $name;
	
	/**
	 * @var POExcelWorksheet worksheet instance 
	 */
	private $worksheet;
	
	/**
	 * Constructor
	 * 
	 * @param POExcelWorksheet $worksheet worksheet instance
	 * @param integer $name row number
	 */
	public function __construct($worksheet, $name)
	{
		$this->name = $name;
		$this->worksheet = $worksheet;
	}
	
	/**
	 * @see POIXMLElement::getAsXML()
	 */
	public function getAsXML()
	{
		$xml = '<row r="' . $this->name . '">';
		
		// add cells
		foreach ($this->cells as $cell) {
			$xml .= $cell->getAsXML();
		}
		
		$xml .= '</row>';
		return $xml;
	}
	
	/**
	 * Generating and return cell instance
	 * 
	 * @param string $name column name
	 * @return POExcelCell cell instance
	 */
	public function getCell($name)
	{
		if (!isset($this->cells[$name])) {
			$this->cells[$name] = new POExcelCell($this, $name);
		}
		return $this->cells[$name];
	}
	
	/**
	 * Return row number
	 * 
	 * @return integer row number
	 */
	public function getName()
	{
		return $this->name;
	}
	
	/**
	 * Return worksheet instance
	 * 
	 * @return POExcelWorksheet 
	 */
	public function getWorksheet()
	{
		return $this->worksheet;
	}
}