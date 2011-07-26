<?php

/**
 * Description of POExcelContentTypes
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
class POExcelContentTypes extends POXMLFile
{
	/**
	 * @var POExcel excel instance
	 */
	private $excel;
	
	/**
	 * Constructor
	 * 
	 * @param POExcel $excel excel instance
	 */
	public function __construct(POExcel $excel)
	{
		$this->excel = $excel;
	}
	
	/**
	 * @see POXMLFile::getAsXMLFile()
	 */
	public function getAsXMLFile()
	{
		// worksheets
		$worksheets = '';
		foreach ($this->excel->getWorkbook()->getWorksheets() as $worksheet) {
			$worksheets .= '<Override PartName="/xl/worksheets/sheet' . $worksheet->getId() . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
		}
		
		// shared strings
		$sharedStrings = '';
		if ($this->excel->getWorkbook()->getSharedStrings()->getCount()) {
			$sharedStrings = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
		}
		
		return $this->renderXMLTemplate('content_types', array(
			'worksheets' => $worksheets,
			'sharedStrings' => $sharedStrings
		));
	}
}