<?php

namespace Aimeos\Controller\ExtJS\Catalog\Export\Text;


/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
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

		$this->object = new \Aimeos\Controller\ExtJS\Catalog\Export\Text\Standard( $this->context );
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


	public function testExportXLSFile()
	{
		$this->object = new \Aimeos\Controller\ExtJS\Catalog\Export\Text\Standard( $this->context );

		$manager = \Aimeos\MShop\Catalog\Manager\Factory::createManager( $this->context );
		$node = $manager->getTree( null, array(), \Aimeos\MW\Tree\Manager\Base::LEVEL_ONE );

		$search = $manager->createSearch();
		$search->setConditions( $search->compare( '==', 'catalog.label', array( 'Root', 'Tee' ) ) );

		$ids = array();
		foreach ( $manager->searchItems( $search ) as $item ) {
			$ids[$item->getLabel()] = $item->getId();
		}

		$params = new \stdClass();
		$params->lang = array( 'de', 'fr' );
		$params->items = array( $node->getId() );
		$params->site = $this->context->getLocale()->getSite()->getCode();

		$result = $this->object->exportFile( $params );
		$this->assertTrue( array_key_exists('file', $result) );

		$file = substr($result['file'], 9, -14);
		$this->assertTrue( file_exists( $file ) );


		$inputFileType = \PHPExcel_IOFactory::identify( $file );
		$objReader = \PHPExcel_IOFactory::createReader( $inputFileType );
		$objPHPExcel = $objReader->load( $file );
		$objPHPExcel->setActiveSheetIndex( 0 );

		if( unlink( $file ) === false ) {
			throw new \RuntimeException( 'Unable to remove export file' );
		}

		$sheet = $objPHPExcel->getActiveSheet();

		$this->assertEquals( 'Language ID', $sheet->getCell( 'A1' )->getValue() );
		$this->assertEquals( 'Text', $sheet->getCell( 'G1' )->getValue() );

		$this->assertEquals( 'de', $sheet->getCell( 'A4' )->getValue() );
		$this->assertEquals( 'Root', $sheet->getCell( 'B4' )->getValue() );
		$this->assertEquals( $ids['Root'], $sheet->getCell( 'C4' )->getValue() );
		$this->assertEquals( 'default', $sheet->getCell( 'D4' )->getValue() );
		$this->assertEquals( 'name', $sheet->getCell( 'E4' )->getValue() );
		$this->assertEquals( '', $sheet->getCell( 'G4' )->getValue() );

		$this->assertEquals( 'de', $sheet->getCell( 'A24' )->getValue() );
		$this->assertEquals( 'Tee', $sheet->getCell( 'B24' )->getValue() );
		$this->assertEquals( $ids['Tee'], $sheet->getCell( 'C24' )->getValue() );
		$this->assertEquals( 'unittype8', $sheet->getCell( 'D24' )->getValue() );
		$this->assertEquals( 'long', $sheet->getCell( 'E24' )->getValue() );
		$this->assertEquals( 'Dies würde die lange Beschreibung der Teekategorie sein. Auch hier machen Bilder einen Sinn.', $sheet->getCell( 'G24' )->getValue() );
	}
}