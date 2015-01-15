<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2011
 * @copyright Aimeos (aimeos.org), 2015
 */


class Controller_ExtJS_Attribute_Export_Text_ExcelTest extends MW_Unittest_Testcase
{
	private $_object;
	private $_context;


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

		$this->_context = TestHelper::getContext();
		$this->_context->getConfig()->set( 'controller/extjs/attribute/export/text/default/container/type', 'PHPExcel' );
		$this->_context->getConfig()->set( 'controller/extjs/attribute/export/text/default/container/format', 'Excel5' );
		$this->_object = new Controller_ExtJS_Attribute_Export_Text_Default( $this->_context );
	}


	/**
	 * Tears down the fixture, for example, closes a network connection.
	 * This method is called after a test is executed.
	 *
	 * @access protected
	 */
	protected function tearDown()
	{
		$this->_object = null;

		Controller_ExtJS_Factory::clear();
		MShop_Factory::clear();
	}




	public function testExportXLSFile()
	{
		$this->_object = new Controller_ExtJS_Attribute_Export_Text_Default( $this->_context );

		$manager = MShop_Attribute_Manager_Factory::createManager( $this->_context );

		$ids = array();
		foreach( $manager->searchItems( $manager->createSearch() ) as $item ) {
			$ids[] = $item->getId();
		}

		$params = new stdClass();
		$params->lang = array( 'de' );
		$params->items = $ids;
		$params->site = 'unittest';

		$result = $this->_object->exportFile( $params );

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

		$this->assertEquals( 'de', $sheet->getCell('A8')->getValue() );
		$this->assertEquals( 'color', $sheet->getCell('B8')->getValue() );
		$this->assertEquals( 'red', $sheet->getCell('C8')->getValue() );
		$this->assertEquals( 'default', $sheet->getCell('D8')->getValue() );
		$this->assertEquals( 'name', $sheet->getCell('E8')->getValue() );
		$this->assertEquals( '', $sheet->getCell('G8')->getValue() );


		$this->assertEquals( '', $sheet->getCell('A124')->getValue() );
		$this->assertEquals( 'width', $sheet->getCell('B124')->getValue() );
		$this->assertEquals( '29', $sheet->getCell('C124')->getValue() );
		$this->assertEquals( 'default', $sheet->getCell('D124')->getValue() );
		$this->assertEquals( 'name', $sheet->getCell('E124')->getValue() );
		$this->assertEquals( '29', $sheet->getCell('G124')->getValue() );
	}
}