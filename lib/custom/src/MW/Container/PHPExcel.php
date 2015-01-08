<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015
 * @package MW
 * @subpackage Container
 */


/**
 * Implementation of PHPExcel containers.
 *
 * @package MW
 * @subpackage Container
 */
class MW_Container_PHPExcel
	extends MW_Container_Abstract
	implements MW_Container_Interface
{
	private $_container;
	private $_format;


	/**
	 * Opens an existing container or creates a new one.
	 *
	 * @param string $resourcepath Path to the resource like a file
	 * @param string $format Format of the content objects inside the container
	 * @param array $options Associative list of key/value pairs for configuration
	 */
	public function __construct( $resourcepath, $format, array $options = array() )
	{
		if( file_exists( $resourcepath ) )
		{
			$type = PHPExcel_IOFactory::identify( $resourcepath );
			$reader = PHPExcel_IOFactory::createReader( $type );
			$this->_container = $reader->load( $resourcepath );
		}
		else
		{
			$this->_container = new PHPExcel();
			$this->_container->removeSheetByIndex( 0 );

			switch( $format )
			{
				case 'Excel5':
					$resourcepath .= '.xls';
					break;
				case 'Excel2003XML':
					$resourcepath .= '.xml';
					break;
				case 'Excel2007':
					$resourcepath .= '.xlsx';
					break;
				case 'OOCalc':
					$resourcepath .= '.ods';
					break;
				case 'SYLK':
					$resourcepath .= '.slk';
					break;
				case 'Gnumeric':
					$resourcepath .= '.gnumeric';
					break;
				case 'CSV':
					$resourcepath .= '.csv';
					break;
			}
		}

		parent::__construct( $resourcepath, $options );

		$this->_iterator = $this->_container->getWorksheetIterator();

		$this->_resourcepath = $resourcepath;
		$this->_format = $format;
	}


	/**
	 * Creates a new content object.
	 *
	 * @param string $name Name of the content
	 * @return MW_Container_Content_Interface New content object
	 */
	public function create( $name )
	{
		$sheet = $this->_container->createSheet();
		$sheet->setTitle( $name );

		return new MW_Container_Content_PHPExcel( $sheet, $name, $this->_getOptions() );
	}


	/**
	 * Adds content data to the container.
	 *
	 * @param MW_Container_Content_Interface $content Content object
	 */
	public function add( MW_Container_Content_Interface $content )
	{
		// was already added to the PHPExcel object by createSheet()
	}


	/**
	 * Returns the element specified by its name.
	 *
	 * @param string $name Name of the content object that should be returned
	 * @return MW_Container_Content_Interface Content object
	 */
	function get( $name )
	{
		if( ( $sheet = $this->_container->getSheetByName( $name ) ) === null ) {
			throw new MW_Container_Exception( sprintf( 'No sheet "%1$s" available', $name ) );
		}

		return new MW_Container_Content_PHPExcel( $sheet, $sheet->getTitle(), $this->_getOptions() );
	}


	/**
	 * Cleans up and saves the container.
	 */
	public function close()
	{
		$writer = PHPExcel_IOFactory::createWriter( $this->_container, $this->_format );
		$writer->save( $this->_resourcepath );
	}


	/**
	 * Return the current element.
	 *
	 * @return MW_Container_Content_Interface Content object with PHPExcel sheet
	 */
	function current()
	{
		$sheet = $this->_iterator->current();

		return new MW_Container_Content_PHPExcel( $sheet, $sheet->getTitle(), $this->_getOptions() );
	}


	/**
	 * Returns the key of the current element.
	 *
	 * @return integer Index of the PHPExcel sheet
	 */
	function key()
	{
		return $this->_iterator->key();
	}


	/**
	 * Moves forward to next PHPExcel sheet.
	 */
	function next()
	{
		return $this->_iterator->next();
	}


	/**
	 * Rewinds to the first PHPExcel sheet.
	 */
	function rewind()
	{
		return $this->_iterator->rewind();
	}


	/**
	 * Checks if the current position is valid.
	 *
	 * @return boolean True on success or false on failure
	 */
	function valid()
	{
		return $this->_iterator->valid();
	}
}