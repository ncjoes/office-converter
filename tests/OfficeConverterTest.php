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
    /**
     * @var OfficeConverter $converter
     */
    private $converter;
    private $outDir;

    public function setUp()
    {
        parent::setUp();

        $file = __DIR__.'\sources\test1.docx';
        $this->outDir = __DIR__.'\results';

        $this->converter = new OfficeConverter($file, $this->outDir);
    }

    public function testDocxToPdfConversion()
    {
        $output = $this->converter->convertTo('result1.pdf');

        $this->assertFileExists($output);
    }

    public function testDocxToHtmlConversion()
    {
        $output = $this->converter->convertTo('result1.html');

        $this->assertFileExists($output);
    }
}
