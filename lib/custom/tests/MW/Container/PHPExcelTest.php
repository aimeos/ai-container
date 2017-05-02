<?php

namespace Aimeos\MW\Container;


/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015-2016
 */
class PHPExcelTest extends \PHPUnit\Framework\TestCase
{
	protected function setUp()
	{
		if( !class_exists( '\PHPExcel' ) ) {
			$this->markTestSkipped( 'PHPExcel not available' );
		}
	}


	public function testExistingFile()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'excel5.xls';

		new \Aimeos\MW\Container\PHPExcel( $filename, 'Excel5', [] );
	}


	public function testNewFile()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'tempfile';

		$container = new \Aimeos\MW\Container\PHPExcel( $filename, 'Excel5', [] );
		$container->close();

		$result = file_exists( $container->getName() );
		unlink( $container->getName() );

		$this->assertTrue( $result );
		$this->assertEquals( '.xls', substr( $container->getName(), -4 ) );
		$this->assertFalse( file_exists( $container->getName() ) );
	}


	public function testFormat()
	{
		$container = new \Aimeos\MW\Container\PHPExcel( 'tempfile', 'Excel2007', [] );
		$this->assertEquals( '.xlsx', substr( $container->getName(), -5 ) );

		$container = new \Aimeos\MW\Container\PHPExcel( 'tempfile', 'OOCalc', [] );
		$this->assertEquals( '.ods', substr( $container->getName(), -4 ) );

		$container = new \Aimeos\MW\Container\PHPExcel( 'tempfile', 'CSV', [] );
		$this->assertEquals( '.csv', substr( $container->getName(), -4 ) );
	}


	public function testAdd()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'tempfile';

		$container = new \Aimeos\MW\Container\PHPExcel( $filename, 'Excel5', [] );
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
		$container = new \Aimeos\MW\Container\PHPExcel( $filename, 'Excel5', [] );

		$this->assertInstanceOf( '\\Aimeos\\MW\\Container\\Content\\Iface', $container->get( 'Sheet2' ) );

		$this->setExpectedException( '\\Aimeos\\MW\\Container\\Exception' );
		$container->get( 'abc' );
	}


	public function testIterator()
	{
		$filename = __DIR__ . DIRECTORY_SEPARATOR . 'excel5.xls';

		$container = new \Aimeos\MW\Container\PHPExcel( $filename, 'Excel5', [] );

		$result = 0;
		foreach( $container as $key => $content ) {
			$result++;
		}

		$this->assertEquals( 3, $result );
	}

}
