<?php

namespace Aimeos\Controller\ExtJS\Product\Import\Text;


/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2011
 * @copyright Aimeos (aimeos.org), 2015-2017
 */
class ExcelTest extends \PHPUnit\Framework\TestCase
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
		if( !class_exists( '\PHPExcel' ) ) {
			$this->markTestSkipped( 'PHPExcel not available' );
		}

		$this->context = \TestHelper::getContext();
		$this->context->getConfig()->set( 'controller/extjs/product/export/text/standard/container/type', 'PHPExcel' );
		$this->context->getConfig()->set( 'controller/extjs/product/export/text/standard/container/format', 'Excel5' );
		$this->context->getConfig()->set( 'controller/extjs/product/import/text/standard/container/type', 'PHPExcel' );
		$this->context->getConfig()->set( 'controller/extjs/product/import/text/standard/container/format', 'Excel5' );

		$this->object = new \Aimeos\Controller\ExtJS\Product\Import\Text\Standard( $this->context );
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

		\Aimeos\Controller\ExtJS\Factory::clear();
		\Aimeos\MShop\Factory::clear();
	}


	public function testImportFromXLSFile()
	{
		$this->object = new \Aimeos\Controller\ExtJS\Product\Import\Text\Standard( $this->context );

		$filename = PATH_TESTS . DIRECTORY_SEPARATOR . 'tmp' . DIRECTORY_SEPARATOR . 'product-import-test.xlsx';

		$phpExcel = new \PHPExcel();
		$phpExcel->setActiveSheetIndex(0);
		$sheet = $phpExcel->getActiveSheet();

		$sheet->setCellValueByColumnAndRow( 0, 2, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 3, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 4, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 5, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 6, 'en' );
		$sheet->setCellValueByColumnAndRow( 0, 7, 'en' );

		$sheet->setCellValueByColumnAndRow( 1, 2, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 3, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 4, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 5, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 6, 'product' );
		$sheet->setCellValueByColumnAndRow( 1, 7, 'product' );

		$sheet->setCellValueByColumnAndRow( 2, 2, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 3, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 4, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 5, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 6, 'ABCD' );
		$sheet->setCellValueByColumnAndRow( 2, 7, 'ABCD' );

		$sheet->setCellValueByColumnAndRow( 3, 2, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 3, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 4, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 5, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 6, 'default' );
		$sheet->setCellValueByColumnAndRow( 3, 7, 'default' );

		$sheet->setCellValueByColumnAndRow( 4, 2, 'long' );
		$sheet->setCellValueByColumnAndRow( 4, 3, 'meta-description' );
		$sheet->setCellValueByColumnAndRow( 4, 4, 'meta-keyword' );
		$sheet->setCellValueByColumnAndRow( 4, 5, 'metatitle' );
		$sheet->setCellValueByColumnAndRow( 4, 6, 'name' );
		$sheet->setCellValueByColumnAndRow( 4, 7, 'short' );

		$sheet->setCellValueByColumnAndRow( 6, 2, 'ABCD: long' );
		$sheet->setCellValueByColumnAndRow( 6, 3, 'ABCD: meta desc' );
		$sheet->setCellValueByColumnAndRow( 6, 4, 'ABCD: meta keywords' );
		$sheet->setCellValueByColumnAndRow( 6, 5, 'ABCD: meta title' );
		$sheet->setCellValueByColumnAndRow( 6, 6, 'ABCD: name' );
		$sheet->setCellValueByColumnAndRow( 6, 7, 'ABCD: short' );

		$objWriter = \PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');
		$objWriter->save($filename);


		$params = new \stdClass();
		$params->site = $this->context->getLocale()->getSite()->getCode();
		$params->items = basename( $filename );

		$this->object->importFile( $params );


		$textManager = \Aimeos\MShop\Text\Manager\Factory::createManager( $this->context );
		$criteria = $textManager->createSearch();

		$expr = [];
		$expr[] = $criteria->compare( '==', 'text.domain', 'product' );
		$expr[] = $criteria->compare( '==', 'text.languageid', 'en' );
		$expr[] = $criteria->compare( '==', 'text.status', 1 );
		$expr[] = $criteria->compare( '~=', 'text.content', 'ABCD:' );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$textItems = $textManager->searchItems( $criteria );

		$textIds = [];
		foreach( $textItems as $item )
		{
			$textManager->deleteItem( $item->getId() );
			$textIds[] = $item->getId();
		}


		$productManager = \Aimeos\MShop\Product\Manager\Factory::createManager( $this->context );
		$listManager = $productManager->getSubManager( 'lists' );
		$criteria = $listManager->createSearch();

		$expr = [];
		$expr[] = $criteria->compare( '==', 'product.lists.domain', 'text' );
		$expr[] = $criteria->compare( '==', 'product.lists.refid', $textIds );
		$criteria->setConditions( $criteria->combine( '&&', $expr ) );

		$listItems = $listManager->searchItems( $criteria );

		foreach( $listItems as $item ) {
			$listManager->deleteItem( $item->getId() );
		}


		foreach( $textItems as $item ) {
			$this->assertEquals( 'ABCD:', substr( $item->getContent(), 0, 5 ) );
		}

		$this->assertEquals( 6, count( $textItems ) );
		$this->assertEquals( 6, count( $listItems ) );

		if( file_exists( $filename ) !== false ) {
			throw new \RuntimeException( 'Import file was not removed' );
		}
	}
}