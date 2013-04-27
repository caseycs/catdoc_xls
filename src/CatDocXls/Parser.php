<?php
namespace CatDocXls;

class Parser
{
    private $sheet_delimiter_default = '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~';

    public function parseToArray($xls_path, $encoding = 'utf-8', $sheet_delimiter = null)
    {
        $sheet_delimiter = $this->processPageDelimiter($sheet_delimiter);

        if (!is_file($xls_path)) {
            throw new Exception('file not found - ' . $xls_path);
        }

        if (!is_readable($xls_path)) {
            throw new Exception('file unreadable - ' . $xls_path);
        }

        $cmd = "xls2csv -d " . $encoding . " -c ';' -b \"" . $sheet_delimiter . "\n\" " . $xls_path;
        $output = array();
        $exit_code = null;
        exec($cmd, $output, $exit_code);

        if ($exit_code !== 0) {
            throw new Exception('xls2csv failed: ' . $cmd . ', exit code ' . $exit_code . ', output: ' . join("\n", $output));
        }

        if ($output === '') {
            throw new Exception('xls2csv output empty');
        }

        $sheets = $this->divideSheets($output, $sheet_delimiter);

        //process lines
        foreach ($sheets as &$sheet) {
            foreach ($sheet as &$line) {
                $line = str_getcsv($line, ';');
            }
        }

        return $sheets;
    }

    private function divideSheets(array $lines, $delimiter)
    {
        $sheets = array();
        do {
            $delimiter_pos = array_search($delimiter, $lines);
            if ($delimiter_pos === false) {
                if (!empty($lines)) {
                    $sheets[] = $lines;
                }
                break;
            }
            $sheets[] = array_splice($lines, 0, $delimiter_pos);
            array_splice($lines, 0, 1);
        } while (true);

        return $sheets;
    }

    private function processPageDelimiter($delimiter)
    {
        if (!$delimiter) {
            $delimiter = $this->sheet_delimiter_default;
        } elseif (strpos($delimiter, ' ') !== false) {
            throw new Exception('spaces in delimiter are not allowed');
        }

        return $delimiter;
    }
}