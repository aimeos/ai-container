<?php

namespace Aimeos\Controller\ExtJS\Catalog\Import\Text;


/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2011
 * @copyright Aimeos (aimeos.org), 2015
 */
class ExcelTest extends \PHPUnit_Framework_TestCase
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
		$this->context->getConfig()->set( 'controller/extjs/catalog/export/text/standard/container/type', 'PHPExcel' );
		$this->context->getConfig()->set( 'controller/extjs/catalog/export/text/standard/container/format', 'Excel5' );
		$this->context->getConfig()->set( 'controller/extjs/catalog/import/text/standard/container/type', 'PHPExcel' );
		$this->context->getConfig()->set( 'controller/extjs/catalog/import/text/standard/container/format', 'Excel5' );

		$this->object = new \Aimeos\Controller\ExtJS\Catalog\Import\Text\Standard( $this->context );
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
		$this->object = new \Aimeos\Controller\ExtJS\Catalog\Import\Text\Standard( $this->context );

		$catalogManager = \Aimeos\MShop\Catalog\Manager\Factory::createManager( $this->context );

		$node = $catalogManager->getTree( null, array(), \Aimeos\MW\Tree\Manager\Base::LEVEL_ONE );

		$params = new \stdClass();
		$params->lang = array( 'en' );
		$params->items = $node->getId();
		$params->site = $this->context->getLocale()->getSite()->getCode();


		$exporter = new \Aimeos\Controller\ExtJS\Catalog\Export\Text\Standard( $this->context );
		$result = $exporter->exportFile( $params );

		$this->assertTrue( array_key_exists('file', $result) );

		$filename = substr($result['file'], 9, -14);
		$this->assertTrue( file_exists( $filename ) );

		$filename2 = PATH_TESTS . DIRECTORY_SEPARATOR . 'tmp' . DIRECTORY_SEPARATOR . 'catalog-import.xls';

		$phpExcel = \PHPExcel_IOFactory::load($filename);

		if( unlink( $filename ) !== true ) {
			throw new \RuntimeException( sprintf( 'Deleting file "%1$s" failed', $filename ) );
		}

		$sheet = $phpExcel->getSheet( 0 );

		$sheet->setCellValueByColumnAndRow( 6, 2, 'Root: delivery info' );
		$sheet->setCellValueByColumnAndRow( 6, 3, 'Root: long' );
		$sheet->setCellValueByColumnAndRow( 6, 4, 'Root: name' );
		$sheet->setCellValueByColumnAndRow( 6, 5, 'Root: payment info' );
		$sheet->setCellValueByColumnAndRow( 6, 6, 'Root: short' );

		$objWriter = \PHPExcel_IOFactory::createWriter( $phpExcel, 'Excel5' );
		$objWriter->save( $filename2 );


		$params = new \stdClass();
		$params->site = $this->context->getLocale()->getSite()->getCode();
		$params->items = basename( $filename2 );

		$this->object->importFile( $params );

		if( file_exists( $filename2 ) !== false ) {
			throw new \RuntimeException( 'Import file was not removed' );
		}

		$textManager = \Aimeos\MShop\Text\Manager\Factory::createManager( $this->context );
		$criteria = $textManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'text.domain', 'catalog' );
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


		$listManager = $catalogManager->getSubManager( 'lists' );
		$criteria = $listManager->createSearch();

		$expr = array();
		$expr[] = $criteria->compare( '==', 'catalog.lists.domain', 'text' );
		$expr[] = $criteria->compare( '==', 'catalog.lists.refid', $textIds );
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