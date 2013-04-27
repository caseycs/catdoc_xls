<?php
namespace CatDocXls\Test;

class ParserTest extends \PHPUnit_Framework_TestCase
{
    public function provider_parse()
    {
        return array(
            array('empty_1sheet.xls', array()),
            array('empty_2sheets.xls', array()),
            array('1sheet.xls', array(array(array('a', 'b')))),
            array('2sheets.xls', array(array(array('a', 'b')), array(array('c', 'd')))),
            array('quotes.xls', array(array(array('"double quoted"', "'single quoted'")))),
            array('semicolon.xls', array(array(array('field with ; semicolon')))),
        );
    }

    /**
     * @dataProvider provider_parse
     */
    public function test_parse($filename, array $expected)
    {
        $Parser = new \CatDocXls\Parser;
        $result = $Parser->parseToArray(__DIR__ . '/../../fixture/' . $filename);
        $this->assertEquals($expected, $result);
    }

    /**
     * @expectedException \CatDocXls\Exception
     */
    public function test_parse_unexisted_file()
    {
        $Parser = new \CatDocXls\Parser;
        $Parser->parseToArray('a.xls');
    }
}