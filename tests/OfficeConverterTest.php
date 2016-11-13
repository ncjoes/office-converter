<?php
/**
 * office-converter
 *
 * Author:  Chukwuemeka Nwobodo (jcnwobodo@gmail.com)
 * Date:    11/13/2016
 * Time:    12:49 AM
 **/

use NcJoes\OfficeConverter\OfficeConverter;

class OfficeConverterTest extends PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        parent::setUp();
    }

    public function testDocxToPdfConversion()
    {
        $file = __DIR__.'\sources\test1.docx';
        $converter = new OfficeConverter($file);
        $result = $converter->convertTo('test1.pdf');

        $this->assertTrue(is_file($file));
    }
}
