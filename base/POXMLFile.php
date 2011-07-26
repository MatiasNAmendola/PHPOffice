<?php

/**
 * Description of POXMLFile
 *
 * @author Aleksey Laletin <aleksey at laletin.ru>
 */
abstract class POXMLFile extends POComponent
{
	/**
	 * Render and return XML
	 */
	abstract public function getAsXMLFile();
	
	/**
	 * Render and return XML template file
	 * 
	 * @param string $_template_ template file
	 * @param array $_data_ data for template
	 * @return string xml
	 */
	protected function renderXMLTemplate($_template_, $_data_)
	{
		extract($_data_, EXTR_OVERWRITE);
		ob_start();
		echo '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
		require(PO_BASEPATH . DIRECTORY_SEPARATOR . 'excel' . DIRECTORY_SEPARATOR . 'templates' . DIRECTORY_SEPARATOR . $_template_ . '.php');
		return ob_get_clean();
	}
}