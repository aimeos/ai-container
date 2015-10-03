<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2011
 * @copyright Aimeos (aimeos.org), 2015
 */


class Controller_ExtJS_Attribute_Export_Text_ExcelTest extends PHPUnit_Framework_TestCase
{
	private $object;
	private $context;


	/**
	 * Sets up the fixture, for example, opens a network connection.
	 * This method is called before a test is executed.
	 *
	 * @access protected
	 */
	protected function setUp()
	{
		if( !class_exists( 'PHPExcel' ) ) {
			$this->markTestSkipped( 'PHPExcel not available' );
		}

		$this->context = TestHelper::getContext();
		$this->context->getConfig()->set( 'controller/extjs/attribute/export/text/default/container/type', 'PHPExcel' );
		$this->context->getConfig()->set( 'controller/extjs/attribute/export/text/default/container/format', 'Excel5' );
		$this->object = new Controller_ExtJS_Attribute_Export_Text_Default( $this->context );
	}


	/**
	 * Tears down the fixture, for example, closes a network connection.
	 * This method is called after a test is executed.
	 *
	 * @access protected
	 */
	protected function tearDown()
	{
		$this->object = null;

		Controller_ExtJS_Factory::clear();
		MShop_Factory::clear();
	}




	public function testExportXLSFile()
	{
		$this->object = new Controller_ExtJS_Attribute_Export_Text_Default( $this->context );

		$manager = MShop_Attribute_Manager_Factory::createManager( $this->context );

		$ids = array();
		foreach( $manager->searchItems( $manager->createSearch() ) as $item ) {
			$ids[] = $item->getId();
		}

		$params = new stdClass();
		$params->lang = array( 'de' );
		$params->items = $ids;
		$params->site = 'unittest';

		$result = $this->object->exportFile( $params );

		$this->assertTrue( array_key_exists('file', $result) );

		$file = substr($result['file'], 9, -14);
		$this->assertTrue( file_exists( $file ) );

		$phpExcel = PHPExcel_IOFactory::load($file);

		if( unlink( $file ) === false ) {
			throw new Exception( 'Unable to remove export file' );
		}


		$phpExcel->setActiveSheetIndex(0);
		$sheet = $phpExcel->getActiveSheet();

		$this->assertEquals( 'Language ID', $sheet->getCell('A1')->getValue() );
		$this->assertEquals( 'Text', $sheet->getCell('G1')->getValue() );

		$this->assertEquals( 'de', $sheet->getCell('A9')->getValue() );
		$this->assertEquals( 'color', $sheet->getCell('B9')->getValue() );
		$this->assertEquals( 'red', $sheet->getCell('C9')->getValue() );
		$this->assertEquals( 'default', $sheet->getCell('D9')->getValue() );
		$this->assertEquals( 'name', $sheet->getCell('E9')->getValue() );
		$this->assertEquals( '', $sheet->getCell('G9')->getValue() );
	}
}