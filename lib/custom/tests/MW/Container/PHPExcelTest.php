<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015
 */


class MW_Container_PHPExcelTest extends MW_Unittest_Testcase
{
	protected function setUp()
	{
		if( !class_exists( 'PHPExcel' ) ) {
			$this->markTestSkipped( 'PHPExcel not available' );
		}
	}


	public function testExistingFile()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'excel5.xls';

		new MW_Container_PHPExcel( $filename, 'Excel5', array() );
	}


	public function testNewFile()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'tempfile';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );
		$container->close();

		$result = file_exists( $container->getName() );
		unlink( $container->getName() );

		$this->assertTrue( $result );
		$this->assertEquals( '.xls', substr( $container->getName(), -4 ) );
		$this->assertFalse( file_exists( $container->getName() ) );
	}


	public function testFormat()
	{
		$container = new MW_Container_PHPExcel( 'tempfile', 'Excel2007', array() );
		$this->assertEquals( '.xlsx', substr( $container->getName(), -5 ) );

		$container = new MW_Container_PHPExcel( 'tempfile', 'OOCalc', array() );
		$this->assertEquals( '.ods', substr( $container->getName(), -4 ) );

		$container = new MW_Container_PHPExcel( 'tempfile', 'CSV', array() );
		$this->assertEquals( '.csv', substr( $container->getName(), -4 ) );
	}


	public function testAdd()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'tempfile';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );
		$container->add( $container->create( 'test' ) );

		$result = 0;
		foreach( $container as $content ) {
			$result++;
		}

		$container->close();
		unlink( $container->getName() );

		$this->assertEquals( 1, $result );
	}


	public function testGet()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'excel5.xls';
		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );

		$this->assertInstanceOf( 'MW_Container_Content_Interface', $container->get( 'Sheet2' ) );
		
		$this->setExpectedException( 'MW_Container_Exception' );
		$container->get( 'abc' );
	}


	public function testIterator()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'excel5.xls';

		$container = new MW_Container_PHPExcel( $filename, 'Excel5', array() );

		$result = 0;
		foreach( $container as $key => $content ) {
			$result++;
		}

		$this->assertEquals( 3, $result );
	}

}
