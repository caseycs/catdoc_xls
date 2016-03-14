<?php
namespace CatDocXls;

class Parser
{
    const SHEET_DELIMITER_DEFAULT = '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~';

    const XSL2CSV_ERROR_TAIL = 'is not OLE file or Error';

    public function xls($path, $sheet_delimiter = null)
    {
        $this->checkFile($path);

        $sheet_delimiter = $this->processPageDelimiter($sheet_delimiter);

        $cmd = "xls2csv " .
            "-d utf-8 " .
            "-c ';' " .
            "-b \"" . $sheet_delimiter . "\n\" " .
            $path . " 2>&1";
        $output = array();
        $exit_code = null;
        exec($cmd, $output, $exit_code);

        if ($exit_code !== 0) {
            throw new Exception('xls2csv failed: ' . $cmd . ', exit code ' . $exit_code . ', output: ' . join("\n", $output));
        }

        if ($output === '' || !is_array($output)) {
            throw new Exception('xls2csv output empty');
        }

        if (count($output) === 1 && strpos($output[0], self::XSL2CSV_ERROR_TAIL) === strlen($output[0]) - strlen(self::XSL2CSV_ERROR_TAIL)) {
            throw new Exception('xls2csv error ' . self::XSL2CSV_ERROR_TAIL);
        }

        $sheets = $this->divideSheets($output, $sheet_delimiter);
        $sheets = array_map(array($this, 'csvLinesToArray'), $sheets);

        return $sheets;
    }

    public function xls2($path, $sheet, $column_count = null)
    {
        $this->checkFile($path);

        $temp_file = tempnam(sys_get_temp_dir(), 'xls2csv_py');

        $cmd = __DIR__ . "/../../xls2csv.0.4.py " .
            "--input {$path} " .
            "--output {$temp_file} " .
            "--sheet {$sheet} " .
            "--enclose-text " .
            " 2>&1";

        $output = array();
        $exit_code = null;

        exec($cmd, $output, $exit_code);

        if ($exit_code !== 0) {
            unlink($temp_file);
            throw new Exception('xls2csv.0.4.py failed: ' . $cmd . ', exit code ' . $exit_code . ', output: ' . join("\n", $output));
        }

        if ($output === '') {
            unlink($temp_file);
            throw new Exception('xls2csv.0.4.py output empty');
        }

        $result = array();

        $handle = fopen($temp_file, 'r');
        while (($data = fgetcsv($handle, null, ';')) !== false) {
            if ($data === array('')) continue; //ignoring empty lines, same all all other
            if ($column_count !== null) {
                $data = array_slice($data, 0, $column_count);
            }
            $result[] = $data;
        }
        fclose($handle);

        unlink($temp_file);

        return $result;
    }

    public function xlsx($path, $sheet_delimiter = null)
    {
        $this->checkFile($path);

        $sheet_delimiter = $this->processPageDelimiter($sheet_delimiter);

        $cmd = "xlsx2csv " .
            "--ignoreempty " .
            "--dateformat '%Y-%m-%d %H:%M:%S' " .
            "--delimiter ';' " .
            "--sheet 0 " .
            "--sheetdelimiter \"" . $sheet_delimiter . "\n\" " .
            $path . " 2>&1";
        $output = array();
        $exit_code = null;

        exec($cmd, $output, $exit_code);

        if ($exit_code !== 0) {
            throw new Exception('xlsx2csv failed: ' . $cmd . ', exit code ' . $exit_code . ', output: ' . join("\n", $output));
        }

        if ($output === '') {
            throw new Exception('xlsx2csv output empty');
        }

        array_shift($output); //remove first sheet delimiter

        $sheets = $this->divideSheets($output, $sheet_delimiter);

        //process lines
        $result = array();
        foreach ($sheets as &$sheet_lines) {
            if (count($sheet_lines) === 1) {//ignoring empty sheet, same as xls2csv
                array_shift($sheet_lines);
                continue;
            }
            array_shift($sheet_lines); //remove sheet title, same as xls2csv
            $result[] = $this->csvLinesToArray($sheet_lines);
        }

        return $result;
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
            $delimiter = self::SHEET_DELIMITER_DEFAULT;
        } elseif (strpos($delimiter, ' ') !== false) {
            throw new Exception('spaces in delimiter are not allowed');
        }

        return $delimiter;
    }

    private function csvLinesToArray(array $csv_lines)
    {
        $result = array();

        //for correct handling multiline string values with fgetcsv we should pass csv through temporary file
        $handle = tmpfile();
        fwrite($handle, join(PHP_EOL, $csv_lines));
        fseek($handle, 0);

        while (($data = fgetcsv($handle, null, ';')) !== false) {
            $result[] = $data;
        }

        fclose($handle);

        return $result;
    }

    private function checkFile($path)
    {
        if (!is_file($path)) {
            throw new Exception('file not found - ' . $path);
        }

        if (!is_readable($path)) {
            throw new Exception('file unreadable - ' . $path);
        }
    }
}
