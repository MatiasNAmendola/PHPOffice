<?php

/**
 * Description of POExcelCell
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
class POExcelCell extends POComponent implements POIXMLElement
{
	/**
	 * @var string column name
	 */
	private $name;
	
	/**
	 * @var POExcelRow row instance
	 */
	private $row;
	
	/**
	 * @var mixed cell value
	 */
	private $value;
	
	/**
	 * Constructor
	 * 
	 * @param POExcelRow $row row instance
	 * @param string $name column name
	 */
	public function __construct($row, $name)
	{
		$this->name = $name;
		$this->row = $row;
	}
	
	/**
	 * @see POIXMLElement::getAsXML()
	 */
	public function getAsXML()
	{
		// type
		$type = '';
		if (is_string($this->value)) {
			$type = ' t="s"';
		}
		
		$xml = '<c r="' . $this->name . $this->row->getName() . '"' . $type . '>';
		
		// add value
		if (is_int($this->value)) {
			$xml .= '<v>' . $this->value . '</v>';
		} elseif (is_string($this->value)) {
			$id = $this->row->getWorksheet()->getWorkbook()->getSharedStrings()->createString($this->value);
			$xml .= '<v>' . $id . '</v>';
		} else {
			// @todo error
		}
		
		$xml .= '</c>';
		return $xml;
	}
	
	/**
	 * Setting cell value
	 * 
	 * @param mixed $value
	 */
	public function setValue($value)
	{
		$this->value = $value;
	}
}