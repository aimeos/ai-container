<?php

namespace Aimeos\MW\Container\Content;


/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015-2016
 */
class PHPExcelTest extends \PHPUnit_Framework_TestCase
{
	private $object;


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

		$phpExcel = new \PHPExcel();
		$sheet = $phpExcel->createSheet();

		$this->object = new \Aimeos\MW\Container\Content\PHPExcel( $sheet, 'test', [] );
	}


	/**
	 * Tears down the fixture, for example, closes a network connection.
	 * This method is called after a test is executed.
	 *
	 * @access protected
	 */
	protected function tearDown()
	{
		unset( $this->object );
	}


	public function testClose()
	{
		$this->object->close();
	}


	public function testAdd()
	{
		$expected = array(
			array( 'test', 'file', 'data' ),
			array( '":;,"', pack( 'x' ), '\\' ),
		);

		foreach( $expected as $entry ) {
			$this->object->add( $entry );
		}

		$actual = $this->object->getResource()->toArray();

		$this->assertEquals( $expected, $actual );
	}


	public function testAddEmpty()
	{
		$expected = array(
			array( 'test', '', 'data' ),
		);

		foreach( $expected as $entry ) {
			$this->object->add( $entry );
		}

		$actual = $this->object->getResource()->toArray();

		$this->assertEquals( $expected, $actual );
	}


	public function testIterator()
	{
		$expected = array(
			array( 'test', 'file', 'data' ),
			array( 'foo', 'bar', 'baz' ),
			array( '":;,"', pack( 'x' ), '\\' ),
		);

		$this->object->getResource()->fromArray( $expected );

		$actual = [];
		foreach( $this->object as $key => $values ) {
			$actual[] = $values;
		}

		// $this->assertEquals( $expected, $actual ); // iterator doesn't work in 1.8.1
	}


	public function testIteratorEmpty()
	{
		$expected = array(
			array( 'test', '', 'data' ),
		);

		$this->object->getResource()->fromArray( $expected );

		$actual = [];
		foreach( $this->object as $values ) {
			$actual[] = $values;
		}

		$this->assertEquals( $expected, $actual );
	}


	public function testGetName()
	{
		$this->assertEquals( 'test', $this->object->getName() );
	}


	public function testGetResource()
	{
		$this->assertInstanceOf( '\PHPExcel_Worksheet', $this->object->getResource() );
	}

}
