<?php
namespace CatDocXls\Test;

class ParserTest extends \PHPUnit_Framework_TestCase
{
    public function provider_parse_common()
    {
        return array(
            array('empty_1sheet', array()),
            array('empty_2sheets', array()),
            array('empty_1sheet_full_2sheet', array(array(array('test')))),
            array('1sheet', array(array(array('a', 'b')))),
            array('2sheets', array(array(array('a', 'b')), array(array('c', 'd')))),
            array('quotes', array(array(array('"double quoted"', "'single quoted'")))),
            array('semicolon', array(array(array('field with ; semicolon')))),
            array('empty_line', array(array(array('line 1'), array('line 3')))),
        );
    }

    /**
     * @dataProvider provider_parse_common
     */
    public function test_parse_common($filename, array $expected)
    {
        $Parser = new \CatDocXls\Parser;

        $result = $Parser->xls(__DIR__ . '/../../fixture/' . $filename . '.xls');
        $this->assertEquals($expected, $result, 'xls fail');

        $result = $Parser->xlsx(__DIR__ . '/../../fixture/' . $filename . '.xlsx');
        $this->assertEquals($expected, $result, 'xlsx fail');
    }

    public function provider_parse_xlsx()
    {
        return array(
            //xls2cvs always uses his own date format, so dates are converted correct only with xlsx2csv
            array('date', array(array(array('2012-01-01 00:00:00')))),
        );
    }

    /**
     * @dataProvider provider_parse_xlsx
     */
    public function test_parse_xlsx($filename, array $expected, $apply = 'both')
    {
        $Parser = new \CatDocXls\Parser;

        $result = $Parser->xlsx(__DIR__ . '/../../fixture/' . $filename . '.xlsx');
        $this->assertEquals($expected, $result, 'xlsx fail');
    }

    /**
     * @expectedException \CatDocXls\Exception
     */
    public function test_parse_unexisted_file_xls()
    {
        $Parser = new \CatDocXls\Parser;
        $Parser->xls('a.xls');
    }

    /**
     * @expectedException \CatDocXls\Exception
     */
    public function test_parse_unexisted_file_xlsx()
    {
        $Parser = new \CatDocXls\Parser;
        $Parser->xlsx('a.xls');
    }
}