<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015
 */


class Controller_ExtJS_Product_Export_Text_ExcelTest extends MW_Unittest_Testcase
{
	private $_object;
	private $_context;


	/**
	 * Runs the test methods of this class.
	 *
	 * @access public
	 * @static
	 */
	public static function main()
	{
		require_once 'PHPUnit/TextUI/TestRunner.php';

		$suite  = new PHPUnit_Framework_TestSuite( 'Controller_ExtJS_Product_Export_Text_ExcelTest' );
		$result = PHPUnit_TextUI_TestRunner::run( $suite );
	}


	/**
	 * Sets up the fixture, for example, opens a network connection.
	 * This method is called before a test is executed.
	 *
	 * @access protected
	 */
	protected function setUp()
	{
		$this->_context = TestHelper::getContext();
		$this->_context->getConfig()->set( 'controller/extjs/product/export/text/default/container/type', 'PHPExcel' );
		$this->_context->getConfig()->set( 'controller/extjs/product/export/text/default/container/format', 'Excel5' );

		$this->_object = new Controller_ExtJS_Product_Export_Text_Default( $this->_context );
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
		$this->_object = new Controller_ExtJS_Product_Export_Text_Default( $this->_context );

		$productManager = MShop_Product_Manager_Factory::createManager( $this->_context );
		$criteria = $productManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'product.code', 'CNE' );
		$criteria->setConditions( $criteria->compare( '==', 'product.code', 'CNE' ) );

		$searchResult = $productManager->searchItems( $criteria );

		if ( ( $productItem = reset( $searchResult ) ) === false ) {
			throw new Exception( 'No item with product code CNE found' );
		}

		$params = new stdClass();
		$params->site = $this->_context->getLocale()->getSite()->getCode();
		$params->items = $productItem->getId();
		$params->lang = 'de';

		$result = $this->_object->exportFile( $params );

		$this->assertTrue( array_key_exists('file', $result) );

		$file = substr($result['file'], 9, -14);

		$this->assertTrue( file_exists( $file ) );

		$inputFileType = PHPExcel_IOFactory::identify( $file );
		$objReader = PHPExcel_IOFactory::createReader( $inputFileType );
		$objReader->setLoadSheetsOnly( $params->lang );
		$objPHPExcel = $objReader->load( $file );

		if( unlink( $file ) === false ) {
			throw new Exception( 'Unable to remove export file' );
		}

		$objWorksheet = $objPHPExcel->getActiveSheet();

		$product = $productItem->toArray();

		for ( $i = 2; $i < 8; $i++ )
		{
			$this->assertEquals( $params->lang, $objWorksheet->getCellByColumnAndRow( 0, $i )->getValue() );
			$this->assertEquals( $product['product.type'], $objWorksheet->getCellByColumnAndRow( 1, $i )->getValue() );
			$this->assertEquals( $product['product.code'], $objWorksheet->getCellByColumnAndRow( 2, $i )->getValue() );
		}

		$this->assertEquals( 'List type', $objWorksheet->getCellByColumnAndRow( 3, 1 )->getValue() );
		$this->assertEquals( 'Text type', $objWorksheet->getCellByColumnAndRow( 4, 1 )->getValue() );
		$this->assertEquals( 'Text ID', $objWorksheet->getCellByColumnAndRow( 5, 1 )->getValue() );
		$this->assertEquals( 'Text', $objWorksheet->getCellByColumnAndRow( 6, 1 )->getValue() );
	}
}