#catdoc_xls
[![Build Status](https://travis-ci.org/caseycs/catdoc_xls.png?branch=master)](https://travis-ci.org/caseycs/catdoc_xls)

PHP wrapper on [catdoc](https://github.com/petewarden/catdoc) util - xls files parser

Usage example:

```php
$Parser = new \CatDocXls\Parser;
$result = $Parser->parseToArray('path/to/file.xls');
print_r($result);
```

See more examples in [ParserTest.php](test/CatDocXls/Test/ParserTest.php)