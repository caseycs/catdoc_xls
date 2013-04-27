catdoc_xls
=========

PHP wrapper on [catdoc](https://github.com/petewarden/catdoc) util - xls files parser

Usage example:

```php
$Parser = new \CatDocXls\Parser;
$result = $Parser->parseToArray('path/to/file.xls');
print_r($result);
```

See more examples in [test/CatDocXls/Test/ParserTest.php](ParserTest.php)