#catdoc_xls
[![Build Status](https://travis-ci.org/caseycs/catdoc_xls.png?branch=master)](https://travis-ci.org/caseycs/catdoc_xls)

Excel files to PHP array convertor (xls/xlsx), wrapper on [catdoc](https://github.com/petewarden/catdoc),
[xls2csv](https://pypi.python.org/pypi/xls2csv) (with few modifications) and [xlsx2csv](https://github.com/dilshod/xlsx2csv).

##Usage example:

```php
$Parser = new \CatDocXls\Parser;
$result = $Parser->xls('path/to/file.xls');
print_r($result);

//some xsl files are not parsed via xls2csv binary correct, so you can try python script
$Parser = new \CatDocXls\Parser;
$result = $Parser->xls2('path/to/file.xls', 0);
print_r($result);

$Parser = new \CatDocXls\Parser;
$result = $Parser->xlsx('path/to/file.xlsx');
print_r($result);
```

See more examples in [ParserTest.php](test/CatDocXls/Test/ParserTest.php)

##Known issues

* Empty lines are ignored - this is hardcoded in xls2csv, and `--ignoreempty` is always passed to xlsx2csv
* Empty sheets are also ignored
* Xls2cvs always output date and datetime fields as days count, so is is passed as is - you should convert them manually, see http://www.linuxquestions.org/questions/red-hat-31/xls2csv-doesn-t-work-with-excel-date-format-703348/
