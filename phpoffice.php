<?php

/**
 * Boot
 */

// Base path to PHPOffice
define('PO_BASEPATH', rtrim(dirname(__FILE__), DIRECTORY_SEPARATOR));

// Include base files
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'base' . DIRECTORY_SEPARATOR . 'POComponent.php');
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'base' . DIRECTORY_SEPARATOR . 'POIXMLElement.php');
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'base' . DIRECTORY_SEPARATOR . 'POXMLFile.php');

// Inlcude Excel files
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'excel' . DIRECTORY_SEPARATOR . 'POExcel.php');
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'excel' . DIRECTORY_SEPARATOR . 'POExcelCell.php');
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'excel' . DIRECTORY_SEPARATOR . 'POExcelContentTypes.php');
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'excel' . DIRECTORY_SEPARATOR . 'POExcelRow.php');
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'excel' . DIRECTORY_SEPARATOR . 'POExcelWorkbook.php');
require_once(PO_BASEPATH . DIRECTORY_SEPARATOR . 'excel' . DIRECTORY_SEPARATOR . 'POExcelWorksheet.php');