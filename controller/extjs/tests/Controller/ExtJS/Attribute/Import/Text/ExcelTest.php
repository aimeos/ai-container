<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2011
 * @copyright Aimeos (aimeos.org), 2015
 */


class Controller_ExtJS_Attribute_Import_Text_ExcelTest extends MW_Unittest_Testcase
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
		$this->context->getConfig()->set( 'controller/extjs/attribute/import/text/default/container/type', 'PHPExcel' );
		$this->context->getConfig()->set( 'controller/extjs/attribute/import/text/default/container/format', 'Excel5' );

		$this->object = new Controller_ExtJS_Attribute_Import_Text_Default( $this->context );
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


	public function testImportFromXLSFile()
	{
		$attributeManager = MShop_Attribute_Manager_Factory::createManager( $this->context );

		$search = $attributeManager->createSearch();
		$search->setConditions( $search->compare( '==', 'attribute.type.code', 'color' ) );

		$ids = array();
		foreach( $attributeManager->searchItems( $search ) as $item ) {
			$ids[] = $item->getId();
		}

		if( empty( $ids ) ) {
			throw new Exception( 'Empty id list' );
		}

		$params = new stdClass();
		$params->lang = array( 'en' );
		$params->items = $ids;
		$params->site = $this->context->getLocale()->getSite()->getCode();

		$exporter = new Controller_ExtJS_Attribute_Export_Text_Default( $this->context );
		$result = $exporter->exportFile( $params );

		$this->assertTrue( array_key_exists('file', $result) );

		$filename = substr($result['file'], 9, -14);
		$this->assertTrue( file_exists( $filename ) );

		$filename2 = 'attribute-import.xls';

		$phpExcel = PHPExcel_IOFactory::load($filename);

		if( unlink( $filename ) !== true ) {
			throw new Exception( sprintf( 'Deleting file "%1$s" failed', $filename ) );
		}

		$sheet = $phpExcel->getSheet( 0 );

		$sheet->setCellValueByColumnAndRow( 6, 2, 'Root: delivery info' );
		$sheet->setCellValueByColumnAndRow( 6, 3, 'Root: long' );
		$sheet->setCellValueByColumnAndRow( 6, 4, 'Root: name' );
		$sheet->setCellValueByColumnAndRow( 6, 5, 'Root: payment info' );
		$sheet->setCellValueByColumnAndRow( 6, 6, 'Root: short' );

		$objWriter = PHPExcel_IOFactory::createWriter( $phpExcel, 'Excel5' );
		$objWriter->save( $filename2 );

		$params = new stdClass();
		$params->site = $this->context->getLocale()->getSite()->getCode();
		$params->items = $filename2;

		$this->object->importFile( $params );

		if( file_exists( $filename2 ) !== false ) {
			throw new Exception( 'Import file was not removed' );
		}

		$textManager = MShop_Text_Manager_Factory::createManager( $this->context );
		$criteria = $textManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'text.languageid', 'en' );
		$expr[] = $criteria->compare( '==', 'text.status', 1 );
		$expr[] = $criteria->compare( '~=', 'text.content', 'Root:' );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$textItems = $textManager->searchItems( $criteria );

		$textIds = array();
		foreach( $textItems as $item )
		{
			$textManager->deleteItem( $item->getId() );
			$textIds[] = $item->getId();
		}


		$listManager = $attributeManager->getSubManager( 'list' );
		$criteria = $listManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'attribute.list.domain', 'text' );
		$expr[] = $criteria->compare( '==', 'attribute.list.refid', $textIds );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$listItems = $listManager->searchItems( $criteria );

		foreach( $listItems as $item ) {
			$listManager->deleteItem( $item->getId() );
		}


		$this->assertEquals( 5, count( $textItems ) );
		$this->assertEquals( 5, count( $listItems ) );

		foreach( $textItems as $item ) {
			$this->assertEquals( 'Root:', substr( $item->getContent(), 0, 5 ) );
		}
	}
}