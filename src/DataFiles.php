<?php

namespace Tigress;

use Shuchkin\SimpleXLSXGen;

/**
 * Class Data Files (PHP version 8.4)
 *
 * @author Rudy Mas <rudy.mas@rudymas.be>
 * @copyright 2025, rudymas.be. (http://www.rudymas.be/)
 * @license https://opensource.org/licenses/GPL-3.0 GNU General Public License, version 3 (GPL-3.0)
 * @version 2025.01.16.1
 * @package Tigress\DataFiles
 */
class DataFiles
{
    private array $data = [];
    private array $footer = [];
    private array $header = [];
    private array $indexList = [];
    private array $mergeCells = [];

    /**
     * Get the version of the DataFiles
     *
     * @return string
     */
    public function version(): string
    {
        return '2025.01.16';
    }

    /**
     * Add a line to the data
     *
     * @param array $data
     * @return void
     */
    public function addLine(array $data): void
    {
        $this->data[] = $data;
    }

    /**
     * Add multiple lines to the data
     *
     * @param array $data
     * @return void
     */
    public function addLines(array $data): void
    {
        $this->data = array_merge($this->data, $data);
    }

    /**
     * Change a line in the data
     */
    public function changeLine(int $index, array $data): void
    {
        $this->data[$index] = $data;
    }

    /**
     * Change multiple lines in the data
     *
     * @param array $indexes
     * @param array $data
     * @return void
     */
    public function changeLines(array $indexes, array $data): void
    {
        foreach ($indexes as $index) {
            $this->data[$index] = $data;
        }
    }

    /**
     * Create an array with data from a database
     *
     * @param array $data
     * @return void
     */
    public function createArray(array $data): void
    {
        $this->data = [];
        foreach ($data as $valueRow) {
            $line = [];
            foreach ($valueRow as $value) {
                $line[] = $value;
            }
            $this->data[] = $line;
        }
    }

    /**
     * Create a CSV file
     *
     * @param string $filename
     * @param string $filepath
     * @param string $delimiter
     * @param string $enclosure
     * @param string $escape
     * @param bool $addIndexStart
     * @param bool $addIndexEnd
     * @param bool $addBOM
     * @return string
     */
    public function createCsvFile(
        string $filename,
        string $filepath = '',
        string $delimiter = ',',
        string $enclosure = '"',
        string $escape = '\\',
        bool   $addIndexStart = true,
        bool   $addIndexEnd = false,
        bool   $addBOM = false
    ): string
    {
        if (!str_ends_with($filename, '.csv')) {
            $filename = $filename . '.csv';
        }
        $lineArray = $this->createFileData($addIndexStart, $addIndexEnd);

        if ($filepath != '' && !file_exists($filepath)) {
            mkdir($filepath, 0777, true);
        }

        if ($filepath) {
            $fp = fopen($filepath . '/' . $filename, 'w');
        } else {
            $fp = fopen($filename, 'w');
        }

        // Add BOM to fix UTF-8 in Excel
        if ($addBOM) {
            fwrite($fp, "\xEF\xBB\xBF");
        }

        foreach ($lineArray as $line) {
            fputcsv($fp, $line, $delimiter, $enclosure, $escape);
        }

        fclose($fp);

        if ($filepath) {
            return $filepath . '/' . $filename;
        }
        return $filename;
    }

    /**
     * Create an Excel file
     *
     * @param string $filename
     * @param string $filepath
     * @param int $fontSize
     * @param bool $addIndexStart
     * @param bool $addIndexEnd
     * @return string
     */
    public function createExcel(
        string $filename,
        string $filepath = '',
        int $fontSize = 13,
        bool $addIndexStart = true,
        bool $addIndexEnd = false
    ): string
    {
        if (!str_ends_with($filename, '.xlsx')) {
            $filename = $filename . '.xlsx';
        }
        $lineArray = $this->createFileData($addIndexStart, $addIndexEnd);

        $xlsx = SimpleXLSXGen::fromArray($lineArray);
        $xlsx->setDefaultFontSize($fontSize);

        if ($filepath != '' && !file_exists($filepath)) {
            mkdir($filepath, 0777, true);
        }

        if (!empty($this->mergeCells)) {
            foreach ($this->mergeCells as $mergeCell) {
                $xlsx->mergeCells($mergeCell);
            }
        }

        if ($filepath) {
            $xlsx->saveAs($filepath . '/' . $filename);
            return $filepath . '/' . $filename;
        }
        $xlsx->saveAs($filename);
        return $filename;
    }

    public function createJsonFile(
        string $filename,
        string $filepath = '',
        bool $addIndexStart = true,
        bool $addIndexEnd = false
    ): string
    {
        if (!str_ends_with($filename, '.json')) {
            $filename = $filename . '.json';
        }
        $lineArray = $this->createFileData($addIndexStart, $addIndexEnd);

        if ($filepath != '' && !file_exists($filepath)) {
            mkdir($filepath, 0777, true);
        }

        if ($filepath) {
            file_put_contents($filepath . '/' . $filename, json_encode($lineArray, JSON_PRETTY_PRINT));
            return $filepath . '/' . $filename;
        }
        file_put_contents($filename, json_encode($lineArray, JSON_PRETTY_PRINT));
        return $filename;
    }

    /**
     * Remove a line from the data
     *
     * @param int $index
     * @return void
     */
    public function removeLine(int $index): void
    {
        unset($this->data[$index]);
    }

    /**
     * Remove multiple lines from the data
     *
     * @param array $indexes
     * @return void
     */
    public function removeLines(array $indexes): void
    {
        foreach ($indexes as $index) {
            unset($this->data[$index]);
        }
    }

    /**
     * Reset the data
     */
    public function reset(): void
    {
        $this->data = [];
    }

    /**
     * Get the data
     *
     * @return array
     */
    public function getData(): array
    {
        return $this->data;
    }

    /**
     * Get the index list
     *
     * @return array
     */
    public function getIndexList(): array
    {
        return $this->indexList;
    }

    /**
     * Set the index list
     *
     * @param array $indexList
     */
    public function setIndexList(array $indexList): void
    {
        $this->indexList = $indexList;
    }

    /**
     * Set the header
     *
     * @param array $header
     * @return void
     */
    public function setHeader(array $header): void
    {
        $this->header = $header;
    }

    /**
     * Set the footer
     *
     * @param array $footer
     * @return void
     */
    public function setFooter(array $footer): void
    {
        $this->footer = $footer;
    }

    /**
     * Set the merge cells
     *
     * @param array $mergeCells
     * @return void
     */
    public function setMergeCells(array $mergeCells): void
    {
        $this->mergeCells = $mergeCells;
    }

    /**
     * Create an array with indexes
     *
     * @param bool $addIndexStart
     * @param bool $addIndexEnd
     * @return array
     */
    private function createFileData(bool $addIndexStart, bool $addIndexEnd): array
    {
        if ($addIndexStart || $addIndexEnd) {
            $indexArray = [];
            foreach ($this->indexList as $index) {
                $indexArray[] = $index;
            }
        }

        $lineArray = [];

        if (!empty($this->header)) {
            foreach ($this->header as $header) {
                $lineArray[] = $this->html_decode($header);
            }
        }

        if ($addIndexStart) {
            $lineArray[] = $indexArray;
        }
        $lineArray = array_merge($lineArray, $this->data);
        if ($addIndexEnd) {
            $lineArray[] = $indexArray;
        }

        if (!empty($this->footer)) {
            foreach ($this->footer as $footer) {
                $lineArray[] = $this->html_decode($footer);
            }
        }

        return $lineArray;
    }

    /**
     * Decode HTML entities
     *
     * @param array $data
     * @return array
     */
    private function html_decode(array $data): array
    {
        foreach ($data as $key => $value) {
            $data[$key] = html_entity_decode($value, ENT_QUOTES, 'UTF-8');
        }
        return $data;
    }
}