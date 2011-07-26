<?php

/**
 * Description of POExcelSharedStrings
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
class POExcelSharedStrings extends POXMLFile
{
	private $strings = array();
	
	public function createString($string)
	{
		$id = count($this->strings);
		$this->strings[] = $string;
		return $id;
	}
	
	public function getCount()
	{
		return count($this->strings);
	}
	
	/**
	 * @see POXMLFile::getAsXMLFile()
	 */
	public function getAsXMLFile()
	{
		// strings
		$strings = '';
		foreach ($this->strings as $string) {
			$strings .= '<si><t>' . $string . '</t></si>';
		}
		
		return $this->renderXMLTemplate('shared_strings', array('strings' => $strings, 'count' => $this->getCount()));
	}
}