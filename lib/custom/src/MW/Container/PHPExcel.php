<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015-2018
 * @package MW
 * @subpackage Container
 */


namespace Aimeos\MW\Container;


/**
 * Implementation of PHPExcel containers.
 *
 * @package MW
 * @subpackage Container
 */
class PHPExcel
	extends \Aimeos\MW\Container\Base
	implements \Aimeos\MW\Container\Iface
{
	private $container;
	private $format;


	/**
	 * Opens an existing container or creates a new one.
	 *
	 * @param string $resourcepath Path to the resource like a file
	 * @param string $format Format of the content objects inside the container
	 * @param array $options Associative list of key/value pairs for configuration
	 */
	public function __construct( $resourcepath, $format, array $options = [] )
	{
		if( file_exists( $resourcepath ) )
		{
			$type = \PHPExcel_IOFactory::identify( $resourcepath );
			$reader = \PHPExcel_IOFactory::createReader( $type );
			$this->container = $reader->load( $resourcepath );
		}
		else
		{
			$this->container = new \PHPExcel();
			$this->container->removeSheetByIndex( 0 );

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

		$this->iterator = $this->container->getWorksheetIterator();

		$this->resourcepath = $resourcepath;
		$this->format = $format;
	}


	/**
	 * Creates a new content object.
	 *
	 * @param string $name Name of the content
	 * @return \Aimeos\MW\Container\Content\Iface New content object
	 */
	public function create( $name )
	{
		$sheet = $this->container->createSheet();
		$sheet->setTitle( $name );

		return new \Aimeos\MW\Container\Content\PHPExcel( $sheet, $name, $this->getOptions() );
	}


	/**
	 * Adds content data to the container.
	 *
	 * @param \Aimeos\MW\Container\Content\Iface $content Content object
	 */
	public function add( \Aimeos\MW\Container\Content\Iface $content )
	{
		// was already added to the PHPExcel object by createSheet()
	}


	/**
	 * Returns the element specified by its name.
	 *
	 * @param string $name Name of the content object that should be returned
	 * @return \Aimeos\MW\Container\Content\Iface Content object
	 */
	function get( $name )
	{
		if( ( $sheet = $this->container->getSheetByName( $name ) ) === null ) {
			throw new \Aimeos\MW\Container\Exception( sprintf( 'No sheet "%1$s" available', $name ) );
		}

		return new \Aimeos\MW\Container\Content\PHPExcel( $sheet, $sheet->getTitle(), $this->getOptions() );
	}


	/**
	 * Cleans up and saves the container.
	 */
	public function close()
	{
		$writer = \PHPExcel_IOFactory::createWriter( $this->container, $this->format );
		$writer->save( $this->resourcepath );
	}


	/**
	 * Return the current element.
	 *
	 * @return \Aimeos\MW\Container\Content\Iface Content object with PHPExcel sheet
	 */
	function current()
	{
		$sheet = $this->iterator->current();

		return new \Aimeos\MW\Container\Content\PHPExcel( $sheet, $sheet->getTitle(), $this->getOptions() );
	}


	/**
	 * Returns the key of the current element.
	 *
	 * @return integer Index of the PHPExcel sheet
	 */
	function key()
	{
		return $this->iterator->key();
	}


	/**
	 * Moves forward to next PHPExcel sheet.
	 */
	function next()
	{
		return $this->iterator->next();
	}


	/**
	 * Rewinds to the first PHPExcel sheet.
	 */
	function rewind()
	{
		return $this->iterator->rewind();
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