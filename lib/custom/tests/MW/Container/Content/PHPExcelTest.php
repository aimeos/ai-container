<?php

/**
 * @license LGPLv3, http://www.gnu.org/licenses/lgpl.html
 * @copyright Metaways Infosystems GmbH, 2013
 * @copyright Aimeos (aimeos.org), 2015
 */


class MW_Container_Content_PHPExcelTest extends MW_Unittest_Testcase
{
	private $_object;


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

		$phpExcel = new PHPExcel();
		$sheet = $phpExcel->createSheet();

		$this->_object = new MW_Container_Content_PHPExcel( $sheet, 'test', array() );
	}


	/**
	 * Tears down the fixture, for example, closes a network connection.
	 * This method is called after a test is executed.
	 *
	 * @access protected
	 */
	protected function tearDown()
	{
		unset( $this->_object );
	}


	public function testClose()
	{
		$this->_object->close();
	}


	public function testAdd()
	{
		$expected = array(
			array( 'test', 'file', 'data' ),
			array( '":;,"', pack( 'x' ), '\\' ),
		);

		foreach( $expected as $entry ) {
			$this->_object->add( $entry );
		}

		$actual = $this->_object->getResource()->toArray();

		$this->assertEquals( $expected, $actual );
	}


	public function testAddEmpty()
	{
		$expected = array(
			array( 'test', '', 'data' ),
		);

		foreach( $expected as $entry ) {
			$this->_object->add( $entry );
		}

		$actual = $this->_object->getResource()->toArray();

		$this->assertEquals( $expected, $actual );
	}


	public function testIterator()
	{
		$expected = array(
			array( 'test', 'file', 'data' ),
			array( '":;,"', pack( 'x' ), '\\' ),
		);

		$this->_object->getResource()->fromArray( $expected );

		$actual = array();
		foreach( $this->_object as $key => $values ) {
			$actual[] = $values;
		}

		$this->assertEquals( $expected, $actual );
	}


	public function testIteratorEmpty()
	{
		$expected = array(
			array( 'test', '', 'data' ),
		);

		$this->_object->getResource()->fromArray( $expected );

		$actual = array();
		foreach( $this->_object as $values ) {
			$actual[] = $values;
		}

		$this->assertEquals( $expected, $actual );
	}


	public function testGetName()
	{
		$this->assertEquals( 'test', $this->_object->getName() );
	}


	public function testGetResource()
	{
		$this->assertInstanceOf( 'PHPExcel_Worksheet', $this->_object->getResource() );
	}

}
