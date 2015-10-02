<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015
 */


class Controller_ExtJS_Product_Export_Text_ExcelTest extends MW_Unittest_Testcase
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
		$this->context->getConfig()->set( 'controller/extjs/product/export/text/default/container/type', 'PHPExcel' );
		$this->context->getConfig()->set( 'controller/extjs/product/export/text/default/container/format', 'Excel5' );

		$this->object = new Controller_ExtJS_Product_Export_Text_Default( $this->context );
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
		$this->object = new Controller_ExtJS_Product_Export_Text_Default( $this->context );

		$productManager = MShop_Product_Manager_Factory::createManager( $this->context );
		$criteria = $productManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'product.code', 'CNE' );
		$criteria->setConditions( $criteria->compare( '==', 'product.code', 'CNE' ) );

		$searchResult = $productManager->searchItems( $criteria );

		if ( ( $productItem = reset( $searchResult ) ) === false ) {
			throw new Exception( 'No item with product code CNE found' );
		}

		$params = new stdClass();
		$params->site = $this->context->getLocale()->getSite()->getCode();
		$params->items = $productItem->getId();
		$params->lang = 'de';

		$result = $this->object->exportFile( $params );

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