<?php

/**
 * Description of POExcel
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
class POExcel extends POComponent
{
	/**
	 * @var string excel file path
	 */
	private $fileName;
	
	/**
	 * @var POExcelWorkbook workbook instance 
	 */
	private $workbook;
	
	/**
	 * Constructor
	 * 
	 * @param string $fileName excel file path
	 */
	public function __construct($fileName)
	{
		$this->fileName = $fileName;
	}
	
	/**
	 * Generate (if not exists) and return workbook
	 * 
	 * @return POExcelWorkbook workbook instance 
	 */
	public function getWorkbook()
	{
		if ($this->workbook === null) {
			$this->workbook = new POExcelWorkbook();
		}
		return $this->workbook;
	}
	
	/**
	 * Save XLSX
	 */
	public function save()
	{
		$zip = new ZipArchive();
		$zip->open($this->fileName, ZipArchive::CREATE);
				
		// _rels
		$zip->addEmptyDir('_rels');
		$zip->addFromString('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
		
		// xl
		$zip->addEmptyDir('xl');
				
		// xl/worksheets
		$zip->addEmptyDir('xl/worksheets');
		
		foreach ($this->getWorkbook()->getWorksheets() as $worksheet) {
			$zip->addFromString('xl/worksheets/sheet' . $worksheet->getId() . '.xml', $worksheet->getAsXMLFile());
		}
		
		$zip->addFromString('xl/workbook.xml', $this->getWorkbook()->getAsXMLFile());
		
		if ($this->getWorkbook()->getSharedStrings()->getCount()) {
			$zip->addFromString('xl/sharedStrings.xml', $this->getWorkbook()->getSharedStrings()->getAsXMLFile());
		}
		
		// xl/_rels
		$zip->addEmptyDir('xl/_rels');
		$zip->addFromString('xl/_rels/workbook.xml.rels', $this->getWorkbook()->getRelationsXMLFile());
		
		// content types
		$contentTypes = new POExcelContentTypes($this);
		$zip->addFromString('[Content_Types].xml', $contentTypes->getAsXMLFile());
		
		$zip->close();
	}
}