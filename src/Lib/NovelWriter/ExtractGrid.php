<?php
/**
 * NovelWriterExtract, a tool for extracting metadata from a novelWriter project.
 *
 * Copyright 2025 Alan Langford. All rights reserved.
 *
 * Licensed under the General Public License, version 3 or higher. See the LICENSE
 * file in the root of this project for details.
 *
 */

namespace Lib\NovelWriter;

use Abivia\Criteria\Criteria;
use DateMalformedStringException;
use DateTimeImmutable;
use DateTimeZone;
use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Html as HtmlWriter;
use SimpleXMLElement;

class ExtractGrid
{
    const string STRUCTURE_KEYWORD = 'story';
    /**
     * @var array|mixed
     */
    protected array $cellStyle;
    protected array $commentBuffer;
    /**
     * @var int|mixed
     */
    protected int $contentLength;
    protected array $contentList;
    /**
     * @var array|mixed|string|string[]|null
     */
    protected string $contentString;
    static protected array $headings = [
        '_blank' => '',
        '_novel' => 'Novel',
        '_sequence' => '#',
        '_active' => 'Active',
        'name' => 'Scene',
        'words' => 'Words',
        '_status' => 'Status',
        'synopsis' => 'Synopsis',
        'value' => 'Value Shift',
        'polarity' => 'Polarity Shift',
        'purpose' => 'Purpose',
        'incite' => 'Inciting Incident',
        'goal' => 'Goal',
        'complication' => 'Complication(s)',
        'turning' => 'Turning Point',
        'crisis' => 'Crisis',
        'climax' => 'Climax',
        'resolution' => '(Non-)Resolution',
        'about' => 'What is this scene about?',
        'impact' => 'Impact of the scene',
        '@pov' => 'Point of View',
        '@plot' => 'Plot',
        'time' => 'Period/Time',
        'tod' => 'Time of Day',
        'duration' => 'Duration',
        '@location' => 'Location(s)',
        '@timeline' => 'Timelines',
        '@focus' => 'Focus Character',
        '@char' => 'Characters',
        'others' => 'Off-stage Characters',
        '@entity' => 'Entities',
        '@object' => 'Objects',
        '@custom' => 'Custom',
        '@mention' => 'Mentions',
        '@story' => 'References',
        'pace' => 'Pace',
        'weather' => 'What is the weather?',
        'appearance' => 'What does it look like?',
        'touch' => 'What do the materials feel like?',
        'aural' => 'What are the sounds?',
        'smell' => 'What are the smells?',
        'clothing' => 'What are the characters wearing?',
        'prose' => 'Quality/cadence in the prose',
        'emotions' => 'What are the characters feeling emotionally?',
        'comments' => 'Additional Notes',
    ];
    protected array $inUse = [];

    /**
     * @var false|mixed
     */
    protected bool $onFirst;
    protected SimpleXMLElement $project;
    protected array $sceneBuffer;
    protected array $sceneData = [];
    protected array $sceneFiles = [];
    protected array $seen = [];
    protected Worksheet $sheet;
    protected string $sourcePath;
    protected Spreadsheet $spreadsheet;
    /**
     * @var array|mixed
     */
    protected array $status;
    static protected array $styles = [
        '*' => ['align' => Alignment::HORIZONTAL_GENERAL, 'onFirst' => false, 'wrap' => true],
        '@' => ['align' => Alignment::HORIZONTAL_CENTER, 'onFirst' => true, 'wrap' => true],
        'comments' => ['align' => Alignment::HORIZONTAL_LEFT, 'wrap' => false],
        'duration' => ['align' => Alignment::HORIZONTAL_RIGHT],
        'time' => ['align' => Alignment::HORIZONTAL_LEFT],
        'words' => [
            'align' => Alignment::HORIZONTAL_RIGHT,
            'numberFormat' => '#,##0',
        ],
    ];
    /**
     * @var array|mixed
     */
    protected array $wordCountStyle;
    protected array $wordCounts;
    protected int $wordTotal;
    protected int $wrapSize = 40;

    public function checkProject(): bool
    {
        if (!isset($this->sourcePath)) {
            echo "Project path is not set.\n";
            return false;
        }
        if (!file_exists("$this->sourcePath/nwProject.nwx")) {
            echo "Path does not contain a novelWriter project.\n";
            return false;
        }
        return true;
    }

    public function checkOutputPath(string $path): array
    {
        $resolved = $this->parsePath($path);
        try {
            $this->getWriterType($resolved);
            $result = [true, $resolved];
        } catch (Exception) {
            $result = [false, "Unrecognized file extension."];
        }
        return $result;
    }

    /**
     * Get a word count by scene without counting headers, comments, etc.
     * @param array $markdown
     * @return int[]
     */
    private function countWords(array $markdown): array
    {
        $count = [-1 => 0];
        $scene = -1;
        $wordCount = 0;
        $hasWords = false;
        foreach ($markdown as $line) {
            $line = rtrim($line);
            if ($line === '') {
                continue;
            }
            if ($this->isScene($line)) {
                $count[$scene++] = $wordCount;
                $wordCount = 0;
                $hasWords = false;
                continue;
            }
            if (preg_match('!^[%@#[]!', $line)) {
                continue;
            }
            $newWords = str_word_count($line);
            $wordCount += $newWords;
            $hasWords = true;
        }
        if ($hasWords) {
            $count[$scene] = $wordCount;
        }
        $count[0] ??= 0;
        $count[0] += $count[-1];
        unset($count[-1]);

        return $count;
    }

    /**
     * Write the scene data to the specified path
     * @param string $path
     * @param string $format
     * @return void
     */
    public function export(string $path, string $format = ''): void
    {
        try {
            $this->loadProject();
        } catch (Exception $exception) {
            echo "Error loading project file: " . $exception->getMessage();
            return;
        }
        try {
            $this->loadScenes();
            $this->spreadsheet = new Spreadsheet();
            if ($format === '') {
                $this->prepareFullSheet();
            } else {
                $this->prepareSheet($format);
            }
            $typeMap = $this->getWriterType($path);
            $this->spreadsheet->setActiveSheetIndex(0);
            $writer = IOFactory::createWriter($this->spreadsheet, $typeMap);
            if ($writer instanceof HtmlWriter) {
                $writer->writeAllSheets();
            }
            $writer->save($this->parsePath($path));
            $this->spreadsheet->disconnectWorksheets();
            unset($this->spreadsheet);
        } catch (Exception $exception) {
            echo $exception->getMessage();
            return;
        }
    }

    private function formatCell(
        Worksheet $sheet,
        int $row,
        int $col,
        array $specs = []
    ): void
    {
        $style = [
            'alignment' => [
                'horizontal' => $specs['align'] ?? Alignment::HORIZONTAL_GENERAL,
                'vertical' => Alignment::VERTICAL_TOP,
                'wrapText' => $specs['wrap'] ?? true,
            ]
        ];
        if ($specs['bold'] ?? false) {
            $style['font'] = ['bold' => true];
        }
        if ($specs['numberFormat'] ?? false) {
            $style['numberFormat'] = ['formatCode' => $specs['numberFormat']];
        }
        $sheet->getStyle([$col, $row])->applyFromArray($style);
    }

    private function formatStyle(int $row, int $col, string $key): void
    {
        $this->formatCell($this->sheet, $row, $col, $this->getStyle($key));
        /*
        $style = $this->getStyle($key);
        $this->sheet->getStyle([$col, $row])->applyFromArray([
            'alignment' => [
                'horizontal' => $style['align'] ?? Alignment::HORIZONTAL_GENERAL,
                'vertical' => Alignment::VERTICAL_TOP,
                'wrapText' => $style['wrap'] ?? true,
            ],
        ]);
        */
    }

    /**
     * Get headings for each column that's in use.
     *
     * @return array|string[]
     */
    private function getHeaders(): array
    {
        $liveHeadings = self::$headings;
        foreach ($liveHeadings as $key => $heading) {
            if (!isset($this->inUse[$key])) {
                unset($liveHeadings[$key]);
            }
        }
        foreach (array_keys($this->inUse) as $key) {
            if (!isset($liveHeadings[$key])) {
                $liveHeadings[$key] = ucfirst($key);
            }
        }
        return $liveHeadings;
    }

    /**
     * Convert this scene data into a string, save the string in the contentString
     * and contentLength properties.
     * @param array $sceneData
     * @param string $column
     * @return void
     */
    private function getSceneData(array $sceneData, string $column): void
    {
        $data = $sceneData[$column] ?? '';
        if (is_array($data)) {
            $this->contentLength = 0;
            $this->contentList = $data;
            foreach ($data as $item) {
                $this->contentLength = max($this->contentLength, strlen($item));
            }
            $this->contentString = implode("\n", $data);
        } else {
            $this->contentString = preg_replace('!\s+!', ' ', $data);
            $this->contentLength = strlen($this->contentString);
        }
    }

    /**
     * Get a pre-defined or default style based on the column name.
     * @param string $key
     * @return array
     */
    private function getStyle(string $key): array
    {
        if (isset(self::$styles[$key])) {
            $style = self::$styles[$key];
        } elseif ($key[0] === '@') {
            $style = self::$styles['@'];
        } else {
            $style = self::$styles['*'];
        }
        return $style;
    }

    /**
     * @param string $path
     * @return string
     * @throws Exception
     */
    private function getWriterType(string $path): string
    {
        $ext = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        return match ($ext) {
            'csv' => IOFactory::WRITER_CSV,
            'html' => IOFactory::WRITER_HTML,
            'ods' => IOFactory::WRITER_ODS,
            'xlsx' => IOFactory::WRITER_XLSX,
            default => throw new Exception("Unsupported file type: $ext"),
        };
    }

    private function isScene(string $line): bool
    {
        return str_starts_with($line, '### ')
            || str_starts_with($line, '###! ');
    }

    /**
     * Extracts scene information from the project XML file.
     * @throws Exception
     */
    private function loadProject(): void
    {
        $this->sceneFiles = [];
        $this->project = new SimpleXMLElement(
            @file_get_contents("$this->sourcePath/nwProject.nwx")
        );
        $parent = '';
        $name = '';
        $this->loadStatus();
        foreach ($this->project->content->item as $item) {
            if ((string)$item['class'] !== 'NOVEL') {
                continue;
            }
            $itemType = (string)$item['type'];
            $handle = (string)$item['handle'];
            $root = (string)$item['root'];
            if ($itemType === 'ROOT') {
                $parent = $handle;
                $name = isset($item->name) ? (string)$item->name : '';
            } elseif ($itemType === 'FILE' && $root === $parent) {
                $statusKey = (string)$item->name['status'];
                $scene = [
                    'handle' => $handle,
                    '_active' => (string)$item->name['active'],
                    '_novel' => $name,
                    '_status' => $this->status[$statusKey] ?? 'Not Set',
                ];
                $scene['words'] = isset($item->meta['wordCount'])
                    ? (string)$item->meta['wordCount'] : '';

                $this->sceneFiles[] = $scene;
            }
        }
    }

    private function loadScenes(): void
    {
        $this->inUse = [
            '_status' => true,
            '_active' => true,
            'name' => true,
            'words' => true,
        ];
        $this->sceneData = [];
        $this->sceneBuffer = [];
        $this->commentBuffer = [];
        $this->wordCounts = [];
        $this->wordTotal = 0;
        // Track if we're in the scene header or the body, so we don't accumulate inline comments.
        $inHeader = true;
        foreach ($this->sceneFiles as $scene) {
            $markdown = explode(
                "\n",
                @file_get_contents("$this->sourcePath/content/{$scene['handle']}.nwd")
            );
            // Get word counts by scene.
            $wordCounts = $this->countWords($markdown);
            $sceneId = 0;
            foreach ($markdown as $line) {
                $line = rtrim($line);
                if ($line === '') {
                    continue;
                }
                if ($this->isScene($line)) {
                    // This is the start of a scene, save the preceding scene, if any.
                    if (count($this->sceneBuffer)) {
                        $this->sceneBuffer['comments'] = $this->commentBuffer;
                        $this->sceneData[] = $this->sceneBuffer;
                    }
                    // Reset the header flag, invalidate the word count, and clear the comment buffer
                    $inHeader = true;
                    $status = $scene['_status'];
                    $this->wordCounts[$status] ??= ['yes' => 0, 'no' => 0, '#yes' => 0, '#no' => 0];
                    $words = $wordCounts[$sceneId++] ?? 0;
                    $this->wordTotal += $words;
                    $active = $scene['_active'];
                    $this->wordCounts[$status][$active] += $words;
                    ++$this->wordCounts[$status]["#$active"];
                    $this->sceneBuffer = [
                        '_active' => $active,
                        '_novel' => $scene['_novel'],
                        '_status' => $status,
                        'name' => trim(substr($line, 4)),
                        'words' => $words,
                    ];
                    $this->commentBuffer = [];
                } elseif (str_starts_with($line, '%')) {
                    if (str_starts_with($line, '%%')) {
                        // We're in the header metadata
                        $inHeader = true;
                    } elseif ($inHeader) {
                        // Look for a story extension
                        $this->parseComment($line);
                    }
                } elseif (str_starts_with($line, '@')) {
                    $this->parseReference($line);
                } else {
                    $inHeader = false;
                }
            }
        }
        if (count($this->sceneBuffer)) {
            $this->sceneBuffer['comments'] = $this->commentBuffer;
            $this->sceneData[] = $this->sceneBuffer;
        }
    }

    /**
     * Build a table of status names so we can map keys to names in the output.
     * @return void
     */
    private function loadStatus(): void
    {
        $this->status = [];
        foreach ($this->project->settings->status->entry as $entry) {
            $this->status[(string)$entry['key']] = (string)$entry;
        }
    }

    /**
     * Examine the content of a comment and extract anything formatted as a story
     * @param string $line
     * @return void
     */
    private function parseComment(string $line): void
    {
        $parts = explode(':', $line, 2);
        $command = strtolower(trim(substr($parts[0], 1)));
        // Handle the two versions of "synopsis".
        if (
            str_starts_with($command, 'synopsis')
            || str_starts_with($command, 'short')
        ) {
            if (count($parts) > 1) {
                $this->sceneBuffer['synopsis'] = trim($parts[1]);
                $this->inUse['synopsis'] = true;
            }
        } elseif (str_starts_with($command, self::STRUCTURE_KEYWORD . '.')) {
            // This is a story attribute
            $subParts = explode('.', $command, 2);
            // Check for a story term and save the text
            if (count($subParts) > 1) {
                $term = trim($subParts[1]);
                $note = count($parts) > 1 ? trim($parts[1]) : '';
                if ($note !== '') {
                    $this->sceneBuffer[$term] = trim($parts[1]);
                    $this->inUse[$term] = true;
                }
            }
        } elseif (!str_starts_with($command, '~')) {
            // Just a regular comment (in the header)
            $this->commentBuffer[] = trim(substr($line, 1));
            $this->inUse['comments'] = true;
        }
    }

    /**
     * Support replacements in output paths:
     * The @d [php-format]@ command will inject the current date (default yyyy-mm-dd)
     * The @z {timezone}@ command selects the timezone (default is UTC). @z must precede @d to work.
     * @param string $path A string with optional commands
     * @return string The string after command processing.
     */
    private function parsePath(string $path): string
    {
        if (preg_match_all('/@[a-z][^@]*?@/i', $path, $matches, PREG_OFFSET_CAPTURE)) {
            $zone = null;
            $delta = 0;
            foreach ($matches[0] as $match) {
                $length = strlen($match[0]);
                $instruction = explode(' ', substr($match[0], 1, -1));
                $command = strtolower(array_shift($instruction));
                $inject = '';
                switch ($command) {
                    case 'd':
                        $format = count($instruction) ? implode(' ', $instruction) : 'Y-m-d';
                        try {
                            $inject = new DateTimeImmutable('now', $zone)
                                ->format($format);
                        } catch (DateMalformedStringException) {
                        }
                        break;
                    case 'z':
                        try {
                            $zone = new DateTimeZone($instruction[0]);
                        } catch (Exception) {
                        }
                        break;
                }
                $start = $match[1] + $delta;
                $path = substr($path, 0, $start) . $inject
                    . substr($path, $start + $length);
                $delta += strlen($inject) - $length;
            }
        }
        return $path;
    }

    /**
     * Parse an @reference in and save the value. If the value is a list, explode and trim it.
     * @param string $line
     * @return void
     */
    private function parseReference(string $line): void
    {
        $parts = explode(':', $line, 2);
        $command = strtolower(trim($parts[0]));
        // Ignore this if there is no value.
        if (count($parts) === 1 || trim($parts[1]) === '') {
            return;
        }
        $list = explode(',', $parts[1]);
        foreach ($list as $key => $item) {
            $item = trim($item);
            if ($item === '') {
                unset($list[$key]);
            } else {
                $list[$key] = $item;
            }
        }
        $this->sceneBuffer[$command] = $list;
        $this->inUse[$command] = true;
    }

    public function prepareFullSheet(): void
    {
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->sheet->setTitle('Scenes');

        // Add and style the headers
        $headers = $this->getHeaders();
        $headerKeys = array_keys($headers);
        $maxColChars = [];
        $col = 1;
        $wordCountCol = 0;
        foreach ($headers as $key => $header) {
            if ($key === 'words') {
                $wordCountCol = $col;
            }
            $this->setHeader($this->sheet, $col, $header);
            $maxColChars[$col] = ceil(1.4 * strlen($header));
            ++$col;
        }

        // Now add the data
        $row = 2;
        foreach ($this->sceneData as $sceneData) {
            $col = 1;
            foreach ($headerKeys as $key) {
                if (isset($sceneData[$key])) {
                    $this->getSceneData($sceneData, $key);
                    $this->sheet->setCellValue([$col, $row], $this->contentString);
                    $this->formatStyle($row, $col, $key);
                    $maxColChars[$col] = max($maxColChars[$col], $this->contentLength);
                }
                ++$col;
            }
            ++$row;
        }
        // Save the total word count
        $this->sheet->setCellValue([$wordCountCol, $row], $this->wordTotal);
        $this->formatStyle($row, $wordCountCol, 'words');

        // Set column widths, save for the last one
        for ($index = 1; $index < count($headers); ++$index) {
            if (isset($maxColChars[$index])) {
                $this->sheet->getColumnDimensionByColumn($index)->setWidth(
                    min($maxColChars[$index], $this->wrapSize)
                );
            }
        }
        $this->wordCountStyle = self::$styles['words'];
        $this->prepareWordCounts();
    }

    private function prepareSheet(string $formatPath): void
    {
        $formats = json_decode(@file_get_contents($formatPath), true);
        if (empty($formats)) {
            throw new Exception("Error reading format file $formatPath\n");
        }
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->sheet->setTitle('Scenes');

        // Add and style the headers
        $headers = ['col0'];
        foreach ($formats['columns'] as $column) {
            if (is_string($column)) {
                $headers[] = self::$headings[$column] ?? '???';
            } else {
                $headers[] = $column['heading'] ?? '???';
            }
        }
        $maxColChars = [];
        foreach ($headers as $col => $header) {
            if ($col === 0) {
                continue;
            }
            $this->setHeader($this->sheet, $col, $header);
            $maxColChars[$col] = ceil(1.4 * strlen($header));
        }

        // If there's a word count, determine which column and format it is in
        $wordCountCol = 0;
        $this->wordCountStyle = self::$styles['words'];
        foreach ($formats['columns'] as $index => $column) {
            if (is_string($column)) {
                if ($column === 'words') {
                    $wordCountCol = $index + 1;
                }
            } else {
                if (($column['key'] ?? false) === 'words') {
                    $wordCountCol = $index + 1;
                    if ($column['style'] ?? false) {
                        $this->wordCountStyle = $column['style'];
                    }
                }
            }
        }

        // Now add the data
        $this->seen = [];
        $criteria = new Criteria();
        $lastNovel = '';
        $sequence = 0;
        $row = 2;
        foreach ($this->sceneData as $sceneData) {
            if (($sceneData['_novel'] ?? '') !== $lastNovel) {
                $sequence = 0;
                $lastNovel = $sceneData['_novel'];
            }
            ++$sequence;
            $col = 1;
            foreach ($formats['columns'] as $column) {
                $seenKey = $headers[$col];
                $this->contentString = '';
                $this->contentList = [];
                $this->contentLength = 0;
                $this->cellStyle = [];
                $this->onFirst = false;
                if (is_string($column)) {
                    switch ($column) {
                        case '_blank':
                            break;
                        case '_sequence':
                            $this->contentString = (string)$sequence;
                            $this->contentLength = strlen($sequence);
                            $this->cellStyle['align'] = Alignment::HORIZONTAL_RIGHT;
                            break;
                        default:
                            $this->getSceneData($sceneData, $column);
                            break;
                    }
                    $this->setCellStyle(['key' => $column]);
                } elseif (isset($column['test'])) {
                    // Conditional data in this column
                    $included = $criteria->evaluate($column['test'], function ($key) use ($sceneData) {
                        return $sceneData[$key] ?? '';
                    });
                    if ($included) {
                        if (isset($column['result'])) {
                            // See if we need to pull data from a different column
                            if (str_starts_with($column['result'], '*')) {
                                $this->getSceneData(
                                    $sceneData, substr($column['result'], 1)
                                );
                            } else {
                                $this->contentString = $column['result'];
                                $this->contentLength = strlen($this->contentString);
                                $this->contentList = [$this->contentString];
                                $this->setCellStyle($column);
                            }
                        } elseif (isset($column['key'])) {
                            $this->getSceneData($sceneData, $column['key']);
                        } else {
                            $this->contentString = '*';
                            $this->contentLength = 1;
                            $this->contentList = ['*'];
                        }
                        $this->setCellStyle($column);
                    }
                } elseif (isset($column['key'])) {
                    // Renamed header and/or filtered data
                    if (($column['exclude'] ?? false) && ($sceneData[$column['key']] ?? false)) {
                        if (is_array($sceneData[$column['key']])) {
                            foreach ($sceneData[$column['key']] as $index => $value) {
                                if (in_array($value, $column['exclude'])) {
                                    unset($sceneData[$column['key']][$index]);
                                }
                            }
                        } elseif (in_array($sceneData[$column['key']], $column['exclude'])) {
                            $sceneData[$column['key']] = '';
                        }
                    }
                    $this->getSceneData($sceneData, $column['key']);
                    $this->setCellStyle($column);
                }
                ++$this->contentLength;
                if ($this->cellStyle['onFirst'] ?? false) {
                    $this->seen[$seenKey] ??= [];
                    foreach ($this->contentList as $newValue) {
                        if (!in_array($newValue, $this->seen[$seenKey])) {
                            $this->seen[$seenKey][] = $newValue;
                            $this->cellStyle['bold'] = true;
                            $this->contentLength = round($this->contentLength * 1.2);
                        }
                    }
                }
                $this->sheet->setCellValue([$col, $row], $this->contentString);
                $this->formatCell($this->sheet, $row, $col, $this->cellStyle);
                $maxColChars[$col] = max($maxColChars[$col], $this->contentLength);
                ++$col;
            }
            ++$row;
        }
        // Save the total word count, if there is one
        if ($wordCountCol) {
            $this->sheet->setCellValue([$wordCountCol, $row], $this->wordTotal);
            $this->formatCell($this->sheet, $row, $wordCountCol, $this->wordCountStyle);
        }


        // Set column widths
        for ($index = 1; $index <= count($headers); ++$index) {
            if (isset($maxColChars[$index])) {
                $this->sheet->getColumnDimensionByColumn($index)->setWidth(
                    min($maxColChars[$index], $formats['wrap'] ?? $this->wrapSize)
                );
            }
        }
        if ($formats['wordCounts'] ?? true) {
            $this->prepareWordCounts();
        }
    }

    private function prepareWordCounts(): void
    {
        $sheet = $this->spreadsheet->createSheet();
        $sheet->setTitle('Statistics');
        $maxStatusChars = 6;
        foreach (array_keys($this->wordCounts) as $status) {
            $maxStatusChars = max($maxStatusChars, strlen($status));
        }
        $sheet->getColumnDimensionByColumn(1)->setWidth(
            1.4 * $maxStatusChars
        );
        $this->setHeader($sheet, 2, 'Scenes');
        $this->setHeader($sheet, 5, 'Words');
        $headerLabels = [
            '', 'Status', 'Active', 'Inactive', 'Total', 'Active', 'Inactive', 'Total'
        ];
        foreach ($headerLabels as $col => $header) {
            if ($col === 0) {
                continue;
            }
            $this->setHeader($sheet, $col, $header, 2);
        }
        ksort($this->wordCounts);
        $this->wordCounts['Total'] = ['yes' => 0, 'no' => 0, '#yes' => 0, '#no' => 0];
        foreach ($this->wordCounts as $counts) {
            $this->wordCounts['Total']['yes'] += $counts['yes'];
            $this->wordCounts['Total']['no'] += $counts['no'];
            $this->wordCounts['Total']['#yes'] += $counts['#yes'];
            $this->wordCounts['Total']['#no'] += $counts['#no'];
        }
        $row = 3;
        $right = $this->wordCountStyle;
        $bold = $right;
        $bold['bold'] = true;
        foreach ($this->wordCounts as $status => $counts) {
            $col = 0;
            // Status
            $sheet->setCellValue([++$col, $row], $status);
            $this->formatCell($sheet, $row, $col);
            // Active scene count
            $sheet->setCellValue([++$col, $row], $counts['#yes']);
            $this->formatCell($sheet, $row, $col, $right);
            // Inactive scene count
            $sheet->setCellValue([++$col, $row], $counts['#no']);
            $this->formatCell($sheet, $row, $col, $right);
            // Total scene count
            $sheet->setCellValue([++$col, $row], $counts['#yes'] + $counts['#no']);
            $this->formatCell($sheet, $row, $col, $bold);
            // Active word count
            $sheet->setCellValue([++$col, $row], $counts['yes']);
            $this->formatCell($sheet, $row, $col, $right);
            // Inactive word count
            $sheet->setCellValue([++$col, $row], $counts['no']);
            $this->formatCell($sheet, $row, $col, $right);
            // Total word count
            $sheet->setCellValue([++$col, $row], $counts['yes'] + $counts['no']);
            $this->formatCell($sheet, $row, $col, $bold);
            ++$row;
        }
    }

    /**
     * Use a column definition to set attributes of the cell.
     * @param array $column
     * @return self
     */
    private function setCellStyle(array $column): self
    {
        if ($column['key'] ?? false) {
            $style = $this->getStyle($column['key']);
        } else {
            $style = $this->getStyle('*');
        }
        if ($column['style'] ?? false) {
            $style = array_merge($style, $column['style']);
        }
        $this->cellStyle = $style;

        return $this;
    }

    private function setHeader(Worksheet $sheet, int $col, string $header, $row = 1): void
    {
        $sheet->setCellValue([$col, $row], $header);
        $sheet->getStyle([$col, $row])->applyFromArray([
            'font' => ['bold' => true],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_TOP,
            ],
        ]);
    }

    /**
     * Specify where the NovelWriter project is.
     * @param string $path
     * @return self
     */
    public function setSourcePath(string $path): self
    {
        $this->sourcePath = $path;
        return $this;
    }

}
