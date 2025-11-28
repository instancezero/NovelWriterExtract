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
    /**
     * @var array|string[] Attributes associated with character nodes
     */
    static protected array $characterAttributes = [
        '_sequence',
        'name',
        '@tag',
        '_folder',
        'given',
        'surname',
        'pronouns',
        'fate',
        'synopsis',
    ];
    protected array $characterData = [];
    protected array $characterFiles = [];
    protected array $commentBuffer;
    protected array $contentList;
    /**
     * @var array|mixed|string|string[]|null
     */
    protected string $contentString;
    /**
     * @var float
     */
    protected float $contentWidth;
    protected mixed $formats;
    /**
     * @var array|string[] Default column headers for all fields.
     */
    static protected array $headerText = [
        '@char' => 'Characters',
        '@custom' => 'Custom',
        '@entity' => 'Entities',
        '@focus' => 'Focus Character',
        '@location' => 'Location(s)',
        '@mention' => 'Mentions',
        '@object' => 'Objects',
        '@plot' => 'Plot',
        '@pov' => 'Point of View',
        '@story' => 'References',
        '@tag' => 'Tag',
        '@timeline' => 'Timelines',
        '_active' => 'Active',
        '_folder' => 'Folder',
        '_novel' => 'Novel',
        '_sequence' => '#',
        '_status' => 'Status',
        'about' => 'This Scene is About',
        'appearance' => 'Appearance',
        'aural' => 'Environmental Sounds',
        'climax' => 'Climax',
        'clothing' => 'Clothing',
        'comments' => 'Additional Notes',
        'complication' => 'Complication(s)',
        'crisis' => 'Crisis',
        'duration' => 'Duration',
        'emotions' => 'Emotions',
        'fate' => 'Fate',
        'given' => 'Given',
        'goal' => 'Goal',
        'impact' => 'Impact of the scene',
        'incite' => 'Inciting Incident',
        'name' => 'Name',
        'others' => 'Off-stage Characters',
        'pace' => 'Pace',
        'polarity' => 'Polarity Shift',
        'pronouns' => 'Pronouns',
        'prose' => 'Prose Quality/Cadence',
        'purpose' => 'Purpose',
        'resolution' => '(Non-)Resolution',
        'smell' => 'Environmental Smells',
        'surname' => 'Surname',
        'synopsis' => 'Synopsis',
        'time' => 'Period/Time',
        'tod' => 'Time of Day',
        'touch' => 'Tactile',
        'turning' => 'Turning Point',
        'value' => 'Value Shift',
        'weather' => 'Weather',
        'words' => 'Words',
    ];
    protected array $inUse = [];
    /**
     * @var array|string[] Attributes associated with character nodes
     */
    static protected array $locationAttributes = [
        '_sequence',
        'name',
        '@tag',
        '_folder',
        'synopsis',
    ];
    protected array $locationData;
    protected array $locationFiles;
    /**
     * @var array|float[]
     */
    protected array $maxColChars;
    /**
     * @var false|mixed
     */
    protected bool $onFirst;
    protected SimpleXMLElement $project;
    /**
     * @var array|string[] Attributes associated with scene nodes
     */
    static protected array $sceneAttributes = [
        '_novel',
        '_sequence',
        '_active',
        'name',
        'words',
        '_status',
        'synopsis',
        'value',
        'polarity',
        'purpose',
        'incite',
        'goal',
        'complication',
        'turning',
        'crisis',
        'climax',
        'resolution',
        'about',
        'impact',
        '@pov',
        '@plot',
        'time',
        'tod',
        'duration',
        '@location',
        '@timeline',
        '@focus',
        '@char',
        'others',
        '@entity',
        '@object',
        '@custom',
        '@mention',
        '@story',
        'pace',
        'weather',
        'appearance',
        'touch',
        'aural',
        'smell',
        'clothing',
        'prose',
        'emotions',
        'comments',
    ];
    protected array $sceneBuffer;
    protected array $sceneData = [];
    protected array $sceneFiles = [];
    protected array $seen = [];
    private $sheetIndex = 0;
    protected string $sourcePath;
    protected Spreadsheet $spreadsheet;
    /**
     * @var array|mixed
     */
    protected array $status;
    static protected array $styles = [
        '*' => ['align' => Alignment::HORIZONTAL_GENERAL, 'onFirst' => false, 'wrap' => true],
        '@' => ['align' => Alignment::HORIZONTAL_CENTER, 'onFirst' => true, 'wrap' => true],
        '@tag' => ['align' => Alignment::HORIZONTAL_CENTER, 'onFirst' => false, 'wrap' => true],
        'comments' => ['align' => Alignment::HORIZONTAL_LEFT, 'wrap' => false],
        'duration' => ['align' => Alignment::HORIZONTAL_RIGHT],
        'pronouns' => ['align' => Alignment::HORIZONTAL_CENTER, 'onFirst' => false, 'wrap' => true],
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
     * Crudely estimate the column space required for a string.
     * @param string $text
     * @return float
     */
    private function estimateWidth(string $text): float
    {
        // Get rid of anything that's not a "wide" character.
        $wide = 0.6 * strlen(preg_replace('/[^mwA-HJ-LNP-VXZ0-9]/', '', $text));
        // Same with "wider"
        $wider = 0.8 * strlen(preg_replace('/[^MOQW]/', '', $text));
        // Same with "narrower"
        $narrower = 0.6 * strlen(preg_replace('/[^ilI|)(}{ !\'`]/', '', $text));
        return 1.05 * strlen($text) + $wide + $wider - $narrower;
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
            $this->loadCharacters();
            $this->loadLocations();
            $this->loadScenes();
            $this->spreadsheet = new Spreadsheet();
            $this->sheetIndex = 0;
            $this->prepareSheets($format);
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

    private function flagFirst(string $key)
    {
        if ($this->cellStyle['onFirst'] ?? false) {
            $this->seen[$key] ??= [];
            foreach ($this->contentList as $newValue) {
                if (!in_array($newValue, $this->seen[$key])) {
                    $this->seen[$key][] = $newValue;
                    $this->cellStyle['bold'] = true;
                    $this->contentWidth = $this->contentWidth * 1.2;
                }
            }
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

    /**
     * Prepare an array of header texts from a list of column keys (adds col0 so the indexed from 1)
     * @param array $columns
     * @return string[]
     */
    private function getHeaders(array $columns): array
    {
        $headers = [];
        foreach ($columns as $column) {
            if (is_array($column) && ($column['heading'] ?? false)) {
                $headers[] = $column['heading'];
            } elseif (is_string($column)) {
                if ((self::$headerText[$column] ?? false)) {
                    $headers[] = self::$headerText[$column];
                } else {
                    $headers[] = ucfirst($column);
                }
            } else {
                $headers[] = '????';
            }
        }
        return $headers;
    }

    /**
     * Convert this scene data into a string, save the string in the contentString
     * and contentWidth properties.
     * @param array $sceneData
     * @param string $column
     * @return void
     */
    private function getNodeData(array $sceneData, string $column): void
    {
        $node = $sceneData[$column] ?? '';
        if (is_array($node)) {
            $this->contentWidth = 0;
            $this->contentList = $node;
            foreach ($node as $item) {
                $this->contentWidth = max($this->contentWidth, $this->estimateWidth($item));
            }
            $delimiter = ($column[0] === '@') ? "\n" : "\n\n";
            $this->contentString = implode($delimiter, $node);
        } else {
            $this->contentString = preg_replace('!\s+!', ' ', $node);
            $this->contentWidth = $this->estimateWidth($this->contentString);
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

    private function loadCharacters(): void
    {
        $this->inUse = [
            'name' => true,
        ];
        $this->characterData = [];
        foreach ($this->characterFiles as $character) {
            $this->sceneBuffer = [
                '_folder' => $character['_folder'],
                'name' => $character['name']
            ];
            $this->commentBuffer = [];
            $markdown = explode(
                "\n",
                @file_get_contents("$this->sourcePath/content/{$character['handle']}.nwd")
            );
            foreach ($markdown as $line) {
                $line = rtrim($line);
                if ($line === '') {
                    continue;
                }
                if (str_starts_with($line, '%')) {
                    if (!str_starts_with($line, '%%')) {
                        // Look for a story extension
                        $this->parseComment($line);
                    }
                } elseif (str_starts_with($line, '@')) {
                    $this->parseReference($line);
                }
            }
            $this->sceneBuffer['comments'] = $this->commentBuffer;
            $this->characterData[] = $this->sceneBuffer;
        }
    }

    private function loadLocations()
    {
        $this->inUse = [
            'name' => true,
        ];
        $this->locationData = [];
        foreach ($this->locationFiles as $location) {
            $this->sceneBuffer = [
                '_folder' => $location['_folder'],
                'name' => $location['name']
            ];
            $this->commentBuffer = [];
            $markdown = explode(
                "\n",
                @file_get_contents("$this->sourcePath/content/{$location['handle']}.nwd")
            );
            foreach ($markdown as $line) {
                $line = rtrim($line);
                if ($line === '') {
                    continue;
                }
                if (str_starts_with($line, '%')) {
                    if (!str_starts_with($line, '%%')) {
                        // Look for a story extension
                        $this->parseComment($line);
                    }
                } elseif (str_starts_with($line, '@')) {
                    $this->parseReference($line);
                }
            }
            $this->sceneBuffer['comments'] = $this->commentBuffer;
            $this->locationData[] = $this->sceneBuffer;
        }
    }

    /**
     * Extracts scene information from the project XML file.
     * @throws Exception
     */
    private function loadProject(): void
    {
        $this->characterFiles = [];
        $folders = [];
        $this->locationFiles = [];
        $this->sceneFiles = [];
        $this->project = new SimpleXMLElement(
            @file_get_contents("$this->sourcePath/nwProject.nwx")
        );
        $novelParent = '';
        $name = '';
        $this->loadStatus();
        foreach ($this->project->content->item as $item) {
            $itemType = (string)$item['type'] ?? '';
            $handle = (string)$item['handle'] ?? '';
            $root = (string)$item['root'] ?? '';
            $nodeParent = (string)$item['parent'] ?? '';
            switch ((string)$item['class']) {
                case 'CHARACTER':
                    if ($itemType === 'FILE') {
                        $character = [
                            'handle' => $handle,
                            'name' => isset($item->name) ? (string)$item->name : '',
                            '_folder' => $folders[$nodeParent] ?? '',
                        ];
                        $this->characterFiles[] = $character;
                    } elseif ($itemType === 'ROOT') {
                        $folders[$handle] = '';
                    } elseif ($itemType === 'FOLDER') {
                        $folders[$handle] = isset($item->name) ? (string)$item->name : '';
                    }
                    break;
                case 'NOVEL':
                    if ($itemType === 'ROOT') {
                        $novelParent = $handle;
                        $name = isset($item->name) ? (string)$item->name : '';
                    } elseif ($itemType === 'FILE' && $root === $novelParent) {
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
                    break;
                case 'WORLD':
                    if ($itemType === 'FILE') {
                        $location = [
                            'handle' => $handle,
                            'name' => isset($item->name) ? (string)$item->name : '',
                            '_folder' => $folders[$nodeParent] ?? '',
                        ];
                        $this->locationFiles[] = $location;
                    } elseif ($itemType === 'ROOT') {
                        $folders[$handle] = '';
                    } elseif ($itemType === 'FOLDER') {
                        $folders[$handle] = isset($item->name) ? (string)$item->name : '';
                    }
                    break;
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
                $this->sceneBuffer['synopsis'] ??= [];
                $this->sceneBuffer['synopsis'][] = trim($parts[1]);
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
                    $this->sceneBuffer[$term] ??= [];
                    $this->sceneBuffer[$term][] = trim($parts[1]);
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

    private function prepareCharacters()
    {
        // Add and style the headers
        $columns = self::$characterAttributes;
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Characters');

        $headers = $this->getHeaders($columns);
        $this->maxColChars = [];
        foreach ($headers as $col0 => $header) {
            $col1 = $col0 + 1;
            $this->setHeader($sheet, $col1, $header);
            $this->maxColChars[$col1] = ceil(1.4 * strlen($header));
        }

        // Now add the data
        $this->prepareColumns($sheet, $this->characterData, $columns, $headers);

        $this->setColumnWidths($sheet, $headers);

        return true;
    }

    private function prepareColumn(array $column, array $sceneData): void
    {
        // Check for renamed header and/or filtered data
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
        $this->getNodeData($sceneData, $column['key']);
        $this->setCellStyle($column);

    }

    private function prepareColumnConditional(array $column, array $sceneData): void
    {
        if (isset($column['result'])) {
            // See if we need to pull data from a different column
            if (str_starts_with($column['result'], '*')) {
                $this->getNodeData(
                    $sceneData, substr($column['result'], 1)
                );
            } else {
                $this->contentString = $column['result'];
                $this->contentWidth = $this->estimateWidth($this->contentString);
                $this->contentList = [$this->contentString];
                $this->setCellStyle($column);
            }
        } elseif (isset($column['key'])) {
            $this->getNodeData($sceneData, $column['key']);
        } else {
            $this->contentString = '*';
            $this->contentWidth = 1;
            $this->contentList = ['*'];
        }
        $this->setCellStyle($column);
    }

    private function prepareColumns(
        Worksheet $sheet,
        array $nodes,
        array $columns,
        array $headers
    ): int
    {
        $this->seen = [];
        $criteria = new Criteria();
        $lastNovel = '';
        $sequence = 0;
        $row = 2;
        foreach ($nodes as $node) {
            if (($node['_novel'] ?? '') !== $lastNovel) {
                $sequence = 0;
                $lastNovel = $node['_novel'];
            }
            ++$sequence;
            $col0 = 0;
            foreach ($columns as $columnSpecification) {
                $col1 = $col0 + 1;
                $seenKey = $headers[$col0];
                $this->contentString = '';
                $this->contentList = [];
                $this->contentWidth = 0;
                $this->cellStyle = [];
                $this->onFirst = false;
                if (is_string($columnSpecification)) {
                    switch ($columnSpecification) {
                        case '_blank':
                            break;
                        case '_sequence':
                            $this->contentString = (string)$sequence;
                            $this->contentWidth = $this->estimateWidth($sequence);
                            $this->cellStyle['align'] = Alignment::HORIZONTAL_RIGHT;
                            break;
                        default:
                            $this->getNodeData($node, $columnSpecification);
                            break;
                    }
                    $this->setCellStyle(['key' => $columnSpecification]);
                } elseif (isset($columnSpecification['test'])) {
                    // Conditional data in this column
                    $included = $criteria->evaluate($columnSpecification['test'],
                        function ($key) use ($node) {
                            return $node[$key] ?? '';
                        }
                    );
                    if ($included) {
                        $this->prepareColumnConditional($columnSpecification, $node);
                    }
                } elseif (isset($columnSpecification['key'])) {
                    // Renamed header and/or filtered data
                    $this->prepareColumn($columnSpecification, $node);
                }
                ++$this->contentWidth;
                $this->flagFirst($seenKey);
                $sheet->setCellValue([$col1, $row], $this->contentString);
                $this->formatCell($sheet, $row, $col1, $this->cellStyle);
                $this->maxColChars[$col1] = max($this->maxColChars[$col1], $this->contentWidth);
                ++$col0;
            }
            ++$row;
        }

        return $row;
    }

    private function prepareLocations(): bool
    {
        // Add and style the headers
        $columns = self::$locationAttributes;
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Locations');

        $headers = $this->getHeaders($columns);
        $this->maxColChars = [];
        foreach ($headers as $col0 => $header) {
            $col1 = $col0 + 1;
            $this->setHeader($sheet, $col1, $header);
            $this->maxColChars[$col1] = ceil(1.4 * strlen($header));
        }

        // Now add the data
        $this->prepareColumns($sheet, $this->locationData, $columns, $headers);

        $this->setColumnWidths($sheet, $headers);

        return true;
    }

    private function prepareScenes(): bool
    {
        // Add and style the headers
        $columns = $this->formats['scenes'] ?? ($this->formats['columns'] ?? true);
        // If the scene column specification is 'true', then include all columns except the blank.
        if ($columns === true) {
            $columns = self::$sceneAttributes;
        } elseif ($columns === false) {
            return false;
        }
        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet->setTitle('Scenes');

        $headers = $this->getHeaders($columns);

        $this->maxColChars = [];
        foreach ($headers as $col0 => $header) {
            $col1 = $col0 + 1;
            $this->setHeader($sheet, $col1, $header);
            $this->maxColChars[$col1] = ceil(1.4 * strlen($header));
        }

        // If there's a word count, determine which column and format it is in
        $wordCountCol = 0;
        foreach ($columns as $col0 => $columnDefinition) {
            $col1 = $col0 + 1;
            if (is_string($columnDefinition)) {
                // No format override specified, just track the column.
                if ($columnDefinition === 'words') {
                    $wordCountCol = $col1;
                }
            } else {
                // Check for a style specification
                if (($columnDefinition['key'] ?? false) === 'words') {
                    $wordCountCol = $col1;
                    if ($columnDefinition['style'] ?? false) {
                        $this->wordCountStyle = $columnDefinition['style'];
                    }
                }
            }
        }
        // Now add the data
        $row = $this->prepareColumns($sheet, $this->sceneData, $columns, $headers);

        // Save the total word count, if there is one
        if ($wordCountCol) {
            $sheet->setCellValue([$wordCountCol, $row], $this->wordTotal);
            $this->formatCell($sheet, $row, $wordCountCol, $this->wordCountStyle);
        }

        $this->setColumnWidths($sheet, $headers);

        return true;
    }

    private function prepareSheets(string $formatPath): void
    {
        if ($formatPath === '') {
            $this->formats = [];
        } else {
            $this->formats = json_decode(@file_get_contents($formatPath), true);
            if (empty($this->formats)) {
                throw new Exception("Error reading format file $formatPath\n");
            }
        }
        $this->wordCountStyle = self::$styles['words'];
        $hadContent = $this->prepareScenes();
        if ($this->formats['wordCounts'] ?? true) {
            if ($hadContent) {
                $this->spreadsheet->createSheet();
                ++$this->sheetIndex;
                $this->spreadsheet->setActiveSheetIndex($this->sheetIndex);
            }
            $this->prepareWordCounts();
            $hadContent = true;
        }
        if ($this->formats['characters'] ?? true) {
            if ($hadContent) {
                $this->spreadsheet->createSheet();
                ++$this->sheetIndex;
            }
            $hadContent = $this->prepareCharacters();
        }
        if ($this->formats['locations'] ?? true) {
            if ($hadContent) {
                $this->spreadsheet->createSheet();
                ++$this->sheetIndex;
            }
            $hadContent = $this->prepareLocations();
        }
    }

    private function prepareWordCounts(): void
    {
        $sheet = $this->spreadsheet->getActiveSheet();
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

    private function setColumnWidths(Worksheet $sheet, array $headers)
    {
        // Set column widths
        for ($index = 1; $index <= count($headers); ++$index) {
            if (isset($this->maxColChars[$index])) {
                $sheet->getColumnDimensionByColumn($index)->setWidth(
                    min($this->maxColChars[$index], $this->formats['wrap'] ?? $this->wrapSize)
                );
            }
        }
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
