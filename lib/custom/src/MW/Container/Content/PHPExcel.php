<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015-2016
 * @package MW
 * @subpackage Container
 */


namespace Aimeos\MW\Container\Content;


/**
 * Implementation of the PHPExcel content object.
 *
 * @package MW
 * @subpackage Container
 */
class PHPExcel
	extends \Aimeos\MW\Container\Content\Base
	implements \Aimeos\MW\Container\Content\Iface
{
	private $sheet;
	private $iterator;


	/**
	 * Initializes the PHPExcel content object.
	 *
	 * @param \PHPExcel_Worksheet $sheet PHPExcel sheet
	 * @param array $options Associative list of key/value pairs for configuration
	 */
	public function __construct( \PHPExcel_Worksheet $sheet, $name, array $options = array() )
	{
		parent::__construct( $sheet, $name, $options );

		$this->sheet = $sheet;
		$this->iterator = $sheet->getRowIterator();
	}


	/**
	 * Cleans up and saves the content.
	 * Does nothing for PHPExcel sheets.
	 */
	public function close()
	{
	}


	/**
	 * Adds a row to the content object.
	 *
	 * @param mixed $data Data to add
	 */
	public function add( $data )
	{
		$columnNum = 0;
		$rowNum = $this->iterator->current()->getRowIndex();

		foreach( (array) $data as $value ) {
			$this->sheet->setCellValueByColumnAndRow( $columnNum++, $rowNum, $value );
		}

		$this->iterator->next();
	}


	/**
	 * Return the current row.
	 *
	 * @return array List of values
	 */
	function current()
	{
		if( $this->iterator->valid() === false ) {
			return null;
		}

		$iterator = $this->iterator->current()->getCellIterator();
		$iterator->setIterateOnlyExistingCells( false );

		$result = array();

		foreach( $iterator as $cell ) {
			$result[] = $cell->getValue();
		}

		return $result;
	}


	/**
	 * Returns the key of the current row.
	 *
	 * @return integer Position within the PHPExcel sheet
	 */
	function key()
	{
		return $this->iterator->key();
	}


	/**
	 * Moves forward to next row.
	 */
	function next()
	{
		$this->iterator->next();
	}


	/**
	 * Resets the current row to the beginning of the sheet.
	 */
	function rewind()
	{
		$this->iterator->rewind();
	}


	/**
	 * Checks if the current position is valid.
	 *
	 * @return boolean True on success or false on failure
	 */
	function valid()
	{
		return $this->iterator->valid();
	}
}
