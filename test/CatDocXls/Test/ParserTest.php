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
            array('multiline', array(array(array("line1: line 1\nLine2: line2")))),
        );
    }

    public function provider_parse_xls2()
    {
        return array(
            array('empty_1sheet', array()),
            array('1sheet', array(array('a', 'b'))),
            array('quotes', array(array('"double quoted"', "'single quoted'"))),
            array('semicolon', array(array('field with ; semicolon'))),
            array('empty_line', array(array('line 1'), array('line 3'))),
            array('multiline', array(array("line1: line 1\nLine2: line2"))),
        );
    }

    /**
     * @dataProvider provider_parse_common
     */
    public function test_parse_common_xls($filename, array $expected)
    {
        $Parser = new \CatDocXls\Parser;

        $result = $Parser->xls(__DIR__ . '/../../fixture/' . $filename . '.xls');
        $this->assertEquals($expected, $result, 'xls fail');
    }

    /**
     * @dataProvider provider_parse_xls2
     */
    public function test_parse_common_xls2($filename, array $expected)
    {
        $Parser = new \CatDocXls\Parser;

        $result = $Parser->xls2(__DIR__ . '/../../fixture/' . $filename . '.xls', 0);
        $this->assertEquals($expected, $result, 'xls fail');
    }

    /**
     * @dataProvider provider_parse_common
     */
    public function test_parse_common_xlsx($filename, array $expected)
    {
        $Parser = new \CatDocXls\Parser;

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

    /**
     * @expectedException \CatDocXls\Exception
     */
    public function test_parse_invalid_xls()
    {
        $Parser = new \CatDocXls\Parser;
        $Parser->xls(__DIR__ . '/../../fixture/invalid.xls');
    }
}
