<?php

/**
 * Description of POExcelWorkbook
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
class POExcelWorkbook extends POXMLFile
{
	/**
	 * @var array list of worksheets 
	 */
	private $worksheets = array();
	
	/**
	 * @see POXMLFile::getAsXMLFile()
	 */
	public function getAsXMLFile()
	{		
		$worksheets = '<sheets>';
		foreach ($this->worksheets as $worksheet) {
			$worksheets .= '<sheet name="' . $worksheet->getName() . '" sheetId="' . $worksheet->getId() . '" r:id="rId' . $worksheet->getId() . '" />';
		}
		$worksheets .= '</sheets>';
		
		return $this->renderXMLTemplate('workbook', array('worksheets' => $worksheets));
	}
	
	/**
	 * Return XML file with relations
	 * 
	 * @return string
	 */
	public function getRelationsXMLFile()
	{
		// worksheets
		$worksheets = '';
		foreach ($this->worksheets as $worksheet) {
			$worksheets .= '<Relationship Id="rId' . $worksheet->getId() . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $worksheet->getId() . '.xml"/>';
		}
		
		return $this->renderXMLTemplate('workbook_xml_rels', array(
			'worksheets' => $worksheets
		));
	}
	
	/**
	 * Creating new worksheet
	 * 
	 * @param string $name name of worksheet
	 * @return POExcelWorksheet worksheet instance
	 */
	public function createWorksheet($name)
	{
		$id = count($this->worksheets) + 1;
		return $this->worksheets[$name] = new POExcelWorksheet($this, $id, $name);
	}
	
	/**
	 * Return list of worksheets
	 * 
	 * @return array
	 */
	public function getWorksheets()
	{
		return $this->worksheets;
	}
}