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
use Abivia\Criteria\LogicException;
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
     * @var array|string[] Predefined attributes associated with character nodes
     */
    static protected array $characterAttributeRef = [
        '_sequence',
        'name',
        '@tag',
        '_folder',
        'given',
        'surname',
        'pronouns',
        'age',
        'hair',
        'eyes',
        'skin',
        'build',
        'fate',
        'synopsis',
    ];
    protected array $characterAttributes = [];
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
     * @var array|string[] Override column headers for fields.
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
        '_sla' => 'Sentence Lengths',
        '_status' => 'Status',
        'about' => 'This NovelData is About',
        //'age' => 'Age',
        //'appearance' => 'Appearance',
        'aural' => 'Environmental Sounds',
        //'climax' => 'Climax',
        //'clothing' => 'Clothing',
        'comments' => 'Additional Notes',
        'complication' => 'Complication(s)',
        //'crisis' => 'Crisis',
        //'duration' => 'Duration',
        //'emotions' => 'Emotions',
        //'fate' => 'Fate',
        //'given' => 'Given',
        //'goal' => 'Goal',
        //'hair' => 'Hair',
        'impact' => 'Impact of the scene',
        'incite' => 'Inciting Incident',
        //'name' => 'Name',
        'others' => 'Off-stage Characters',
        //'pace' => 'Pace',
        'polarity' => 'Polarity Shift',
        //'pronouns' => 'Pronouns',
        'prose' => 'Prose Quality/Cadence',
        //'purpose' => 'Purpose',
        'resolution' => '(Non-)Resolution',
        'smell' => 'Environmental Smells',
        //'surname' => 'Surname',
        //'synopsis' => 'Synopsis',
        'time' => 'Period/Time',
        'tod' => 'Time of Day',
        'touch' => 'Tactile',
        'turning' => 'Turning Point',
        'value' => 'Value Shift',
        //'weather' => 'Weather',
        //'words' => 'Words',
    ];
    protected array $inUse = [];
    /**
     * @var array|string[] Attributes associated with character nodes
     */
    static protected array $locationAttributeRef = [
        '_sequence',
        'name',
        '@tag',
        '_folder',
        'synopsis',
    ];
    protected array $locationAttributes = [];
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
    /**
     * @var array|array[]
     */
    protected array $phrases;
    protected float $phraseColWidth;
    protected SimpleXMLElement $project;
    /**
     * @var array|string[] Attributes associated with scene nodes
     */
    static protected array $sceneAttributeRef = [
        '_novel',
        '_sequence',
        '_active',
        'name',
        'words',
        '_status',
        'synopsis',
        '_sla',
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
    protected array $sceneAttributes = [];
    protected array $sceneBuffer;
    /**
     * @var array[NovelData]
     */
    protected array $sceneData = [];
    protected array $sceneFiles = [];
    protected NovelWriterFileLoader $sceneLoader;
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
    protected bool $verbose = true;
    /**
     * @var array|mixed
     */
    protected array $wordCountStyle;
    protected array $wordDistribution;
    protected float $wordDistributionWidth;
    protected int $wordTotal;
    protected int $wrapSize = 40;

    /**
     * Get a list of attributes in use, ordered by the elements in attributeRef.
     * @param array $attributeRef
     * @return array
     */
    private function buildAttributes(array $attributeRef): array
    {
        $attributes = [];
        foreach ($attributeRef as $column) {
            if (isset($this->inUse[$column])) {
                $attributes[$column] = $column;
            }
        }
        foreach (array_keys($this->inUse) as $column) {
            if (!isset($attributes[$column])) {
                $attributes[$column] = $column;
            }
        }
        return array_values($attributes);
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
     * Get word counts by scene status, active state
     * @return array
     */
    private function countWords(): array
    {
        $counts = [];
        foreach ($this->sceneData as $scene) {
            $status = $scene['_status'];
            $active = $scene['_active'];
            $counts[$status] ??= [
                'scenes.yes' => 0, 'scenes.no' => 0, 'scenes.total' => 0,
                'words.yes' => 0, 'words.no' => 0, 'words.total' => 0,
            ];
            $counts[$status]["scenes.$active"]++;
            $counts[$status]['scenes.total']++;
            $counts[$status]["words.$active"] += $scene['words'];
            $counts[$status]['words.total'] += $scene['words'];
        }

        return $counts;
    }

    private function estimateListWidth(): float
    {
        $width = 0.0;
        foreach ($this->contentList as $item) {
            $width = max($this->contentWidth, $this->estimateWidth($item));
        }
        return $width;
    }

    /**
     * Crudely estimate the column space required for a string.
     * @param string $text
     * @param bool $bold
     * @return float
     */
    private function estimateWidth(string $text, bool $bold = false): float
    {
        $width = 0.0;
        $lines = explode("\n", $text);
        foreach ($lines as $line) {
            // Filter anything that's not a "wide" character.
            $wide = 0.6 * strlen(preg_replace('/[^mwA-HJ-LNP-VXZ0-9]/', '', $line));
            // Same with "wider"
            $wider = 0.8 * strlen(preg_replace('/[^MOQW]/', '', $line));
            // Same with "narrower"
            $narrower = 0.6 * strlen(preg_replace('/[^iltI|)(}{ !\'.;:`]/', '', $line));
            $width = max($width, 1.05 * strlen($text) + $wide + $wider - $narrower);
        }
        if ($bold) {
            $width *= self::BOLD_FACTOR;
        }
        return $width;
    }

    /**
     * Write the scene data to the specified path
     * @param string $path
     * @param string $format
     * @return void
     */
    public function export(string $path, string $format = ''): void
    {
        if ($this->verbose) {
            echo "Loading Project\n";
        }
        try {
            $this->loadProject();
        } catch (Exception $exception) {
            echo "Error loading project file: " . $exception->getMessage();
            return;
        }
        try {
            if ($this->verbose) {
                echo "Loading Characters\n";
            }
            $this->loadCharacters();
            if ($this->verbose) {
                echo "Loading Locations\n";
            }
            $this->loadLocations();
            if ($this->verbose) {
                echo "Loading Scenes\n";
            }
            $this->loadScenes();
            $this->spreadsheet = new Spreadsheet();
            $this->sheetIndex = 0;
            $this->prepareSheets($format);
            $typeMap = $this->getWriterType($path);
            $this->spreadsheet->setActiveSheetIndex(0);
            if ($this->verbose) {
                echo "Writing\n";
            }
            $writer = IOFactory::createWriter($this->spreadsheet, $typeMap);
            if ($writer instanceof HtmlWriter) {
                $writer->writeAllSheets();
            }
            $writer->save($this->parsePath($path));
            $this->spreadsheet->disconnectWorksheets();
            unset($this->spreadsheet);
        } catch (Exception $exception) {
            echo 'Exception: ' . $exception->getMessage();
            return;
        }
        if ($this->verbose) {
            echo "Done\n";
        }
    }

    /**
     * If the cell has the onFirst attribute, bold the first occurrence.
     * @param string $key
     * @return void
     */
    private function flagFirst(string $key): void
    {
        if ($this->cellStyle['onFirst'] ?? false) {
            $this->seen[$key] ??= [];
            $hasNewValue = false;
            $newItems = [];
            foreach ($this->contentList as $slot => $newValue) {
                if (!in_array($newValue, $this->seen[$key])) {
                    $this->seen[$key][] = $newValue;
                    $newItems[] = $slot;
                    $hasNewValue = true;
                }
            }
            if (count($newItems) && count($newItems) != count($this->contentList)) {
                foreach ($newItems as $slot) {
                    $this->contentList[$slot] = "*{$this->contentList[$slot]}*";
                }
                $this->estimateListWidth();
                $this->contentString = implode("\n", $this->contentList);
            }
            if ($hasNewValue) {
                $this->cellStyle['bold'] = true;
                $this->contentWidth = $this->contentWidth * self::BOLD_FACTOR;
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

    private function getColumns(mixed $option, array $default): array|false
    {
        // If the column specification is 'true', then include all columns.
        if ($option === true) {
            $option = $default;
        }
        return $option;
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
                    $words = explode('_', $column);
                    $words = array_map(fn($word):string => ucfirst($word), $words);
                    $headers[] = implode(' ', $words);
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
     * @param NovelData $node
     * @param string $column
     * @return void
     */
    private function getNodeData(NovelData $node, string $column): void
    {
        $data = $node[$column] ?? '';
        if (is_array($data)) {
            $this->contentWidth = 0;
            $this->contentList = $data;
            $this->contentWidth = $this->estimateListWidth();
            $delimiter = ($column[0] === '@') ? "\n" : "\n\n";
            $this->contentString = implode($delimiter, $data);
        } else {
            $this->contentList = [$data];
            $this->contentString = preg_replace('! +!', ' ', $data);
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
            '_folder' => true,
        ];
        $this->characterData = [];
        foreach ($this->characterFiles as $file) {
            $this->characterData[] = $this->loadFile($file);
        }
        $this->characterAttributes = $this->buildAttributes(self::$characterAttributeRef);
    }

    /**
     * @param array $file
     * @return void
     */
    private function loadFile(array $file): NovelData
    {
        $markdown = explode(
            "\n",
            @file_get_contents("$this->sourcePath/content/{$file['handle']}.nwd")
        );
        $loader = new NovelWriterFileLoader();
        return $loader->loadFile($file, $markdown, $this->inUse);
    }

    private function loadLocations(): void
    {
        $this->inUse = [
            'name' => true,
        ];
        $this->locationData = [];
        foreach ($this->locationFiles as $location) {
            $this->locationData[] = $this->loadFile($location);
        }
        $this->locationAttributes = $this->buildAttributes(self::$locationAttributeRef);
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
        $this->sceneLoader = new NovelWriterFileLoader();
        // Track if we're in the scene header or the body, so we don't accumulate inline comments.
        foreach ($this->sceneFiles as $sceneNode) {
            $markdown = explode(
                "\n",
                @file_get_contents("$this->sourcePath/content/{$sceneNode['handle']}.nwd")
            );
            $this->sceneLoader->loadScene($sceneNode, $markdown, $this->sceneData);
        }
        $this->profileScenes();
        foreach (array_keys($this->sceneLoader->inUse) as $key) {
            $this->inUse[$key] = true;
        }
        $this->sceneAttributes = $this->buildAttributes(self::$sceneAttributeRef);
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

    private function prepareAnalysis(): bool
    {
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Analysis');

        $fw = $this->estimateWidth('Frequency');
        $this->setHeader($sheet, 1, 'Word', width: 6.0);
        $this->setHeader($sheet, 2, 'Frequency', width: $fw);
        $this->setHeader($sheet, 4, 'Phrase', width: $this->phraseColWidth);
        $this->setHeader($sheet, 5, 'Frequency', width: $fw);

        $sheet->getColumnDimensionByColumn(1)->setWidth($this->wordDistributionWidth);
        $sheet->getColumnDimensionByColumn(4)->setWidth($this->phraseColWidth);

        $row = 2;
        $right = $this->wordCountStyle;
        // Word frequency
        foreach ($this->wordDistribution as $word => $count) {
            // Status
            $sheet->setCellValue([1, $row], $word);
            $sheet->setCellValue([2, $row], $count);
            $this->formatCell($sheet, $row, 2, $right);
            ++$row;
        }
        $row = 2;
        // Status
        foreach ($this->phrases as $word => $count) {
            // Status
            $sheet->setCellValue([4, $row], $word);
            $sheet->setCellValue([5, $row], $count);
            $this->formatCell($sheet, $row, 5, $right);
            ++$row;
        }
        return true;
    }

    private function prepareCharacters()
    {
        $columns = $this->formats['characters'] ?? true;
        if ($columns === false) {
            return false;
        }
        $columns = $this->getColumns($columns, $this->characterAttributes);
        // Add and style the headers
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Characters');

        $headers = $this->setHeaders($sheet, $columns);

        // Now add the data
        $this->prepareColumns($sheet, $this->characterData, $columns, $headers);

        $this->setColumnWidths($sheet, $headers);

        return true;
    }

    private function prepareColumn(array $column, NovelData $sceneData): void
    {
        // Check for renamed header and/or filtered data
        if (($column['exclude'] ?? false) && ($sceneData[$column['key']] ?? false)) {
            if (is_array($sceneData[$column['key']])) {
                foreach ($sceneData[$column['key']] as $index => $value) {
                    if (in_array($value, $column['exclude'])) {
                        $sceneData->unset($column['key'], $index);
                    }
                }
            } elseif (in_array($sceneData[$column['key']], $column['exclude'])) {
                $sceneData[$column['key']] = '';
            }
        }
        $this->getNodeData($sceneData, $column['key']);
        $this->setCellStyle($column);

    }

    private function prepareColumnConditional(array $column, NovelData $sceneData): void
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

    /**
     * @param Worksheet $sheet
     * @param NovelData[] $nodes
     * @param array $columns
     * @param array $headers
     * @return int
     * @throws \Abivia\Criteria\LogicException
     */
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
        $columns = $this->formats['locations'] ?? true;
        if ($columns === false) {
            return false;
        }
        $columns = $this->getColumns($columns, $this->locationAttributes);
        // Add and style the headers
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Locations');

        $headers = $this->setHeaders($sheet, $columns);

        // Now add the data
        $this->prepareColumns($sheet, $this->locationData, $columns, $headers);

        $this->setColumnWidths($sheet, $headers);

        return true;
    }

    /**
     * @throws LogicException
     */
    private function prepareScenes(): bool
    {
        $columns = $this->formats['scenes'] ?? ($this->formats['columns'] ?? true);
        if ($columns === false) {
            return false;
        }
        $columns = $this->getColumns($columns, $this->sceneAttributes);
        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet->setTitle('Scenes');

        $headers = $this->setHeaders($sheet, $columns);

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
        if ($this->verbose) {
            echo "Preparing Scenes\n";
        }
        $hadContent = $this->prepareScenes();
        if ($this->formats['wordCounts'] ?? true) {
            if ($hadContent) {
                $this->spreadsheet->createSheet();
                ++$this->sheetIndex;
                $this->spreadsheet->setActiveSheetIndex($this->sheetIndex);
            }
            if ($this->verbose) {
                echo "Preparing Statistics\n";
            }
            $this->prepareWordCounts();
            $hadContent = true;
        }
        if ($this->formats['characters'] ?? true) {
            if ($hadContent) {
                $this->spreadsheet->createSheet();
                ++$this->sheetIndex;
            }
            if ($this->verbose) {
                echo "Preparing Characters\n";
            }
            $hadContent = $this->prepareCharacters();
        }
        if ($this->formats['locations'] ?? true) {
            if ($hadContent) {
                $this->spreadsheet->createSheet();
                ++$this->sheetIndex;
            }
            if ($this->verbose) {
                echo "Preparing Locations\n";
            }
            $hadContent = $this->prepareLocations();
        }
        if ($this->formats['analysis'] ?? true) {
            if ($hadContent) {
                $this->spreadsheet->createSheet();
                ++$this->sheetIndex;
            }
            if ($this->verbose) {
                echo "Preparing Analysis\n";
            }
            $hadContent = $this->prepareAnalysis();
        }
    }

    private function prepareWordCounts(): void
    {
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Statistics');
        $maxStatusChars = 6;
        $statusList = array_keys($this->sceneLoader->sceneStatusList);
        sort($statusList);
        foreach ($statusList as $status) {
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
        $stats = $this->countWords();
        $columns = array_keys($stats[$statusList[0]]);
        $totals = array_fill_keys($columns, 0);

        $row = 3;
        $right = $this->wordCountStyle;
        $bold = $right;
        $bold['bold'] = true;
        foreach ($statusList as $status) {
            $col = 0;
            $statData = $stats[$status];
            // Status
            $sheet->setCellValue([++$col, $row], $status);
            $this->formatCell($sheet, $row, $col);
            foreach ($columns as $column) {
                $sheet->setCellValue([++$col, $row], $statData[$column]);
                $this->formatCell($sheet, $row, $col, $right);
                $totals[$column] += $statData[$column];
            }
            ++$row;
        }
        $col = 0;
        // Status
        $sheet->setCellValue([++$col, $row], 'Total');
        $this->formatCell($sheet, $row, $col);
        foreach ($columns as $column) {
            $sheet->setCellValue([++$col, $row], $totals[$column]);
            $this->formatCell($sheet, $row, $col, $right);
        }
    }

    private function profileScenes()
    {
        $this->phrases = [];
        $this->phraseColWidth = 5;
        $this->wordDistribution = [];
        $this->wordDistributionWidth = 0;
        $this->wordTotal = 0;
        foreach ($this->sceneData as $scene) {
            $this->wordTotal += $scene->words;
            foreach ($scene->distribution as $word => $frequency) {
                $this->wordDistribution[$word] ??= 0;
                $this->wordDistribution[$word] += $frequency;
            }
            foreach ($scene->phrases as $phrase => $frequency) {
                $this->phrases[$phrase] ??= 0;
                $this->phrases[$phrase] += $scene->phrases[$phrase];
            }
        }
        foreach ($this->wordDistribution as $word => $frequency) {
            if ($frequency < 10) {
                unset($this->wordDistribution[$word]);
                continue;
            }
            $this->wordDistributionWidth = max(
                $this->wordDistributionWidth, $this->estimateWidth($word)
            );
        }
        arsort($this->wordDistribution);
        foreach ($this->phrases as $word => $frequency) {
            if ($frequency < 10) {
                unset($this->phrases[$word]);
                continue;
            }
            $this->phraseColWidth = max($this->phraseColWidth, $this->estimateWidth($word));
        }
        arsort($this->phrases);
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

    private function setHeader(
        Worksheet $sheet,
        int $col,
        string $header,
        int $row = 1,
        float $width = 0.0
    ): void
    {
        $sheet->setCellValue([$col, $row], $header);
        $sheet->getStyle([$col, $row])->applyFromArray([
            'font' => ['bold' => true],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_TOP,
            ],
        ]);
        if ($width != 0.0) {
            $sheet->getColumnDimensionByColumn($col)->setWidth(1.4 * strlen($header));
        }
    }

    private function setHeaders(Worksheet $sheet, array $columns): array
    {
        // Add and style the headers
        $headers = $this->getHeaders($columns);

        $this->maxColChars = [];
        foreach ($headers as $col0 => $header) {
            $col1 = $col0 + 1;
            $this->setHeader($sheet, $col1, $header);
            $this->maxColChars[$col1] = ceil(1.4 * strlen($header));
        }
        $sheet->freezePane('A2');

        return $headers;
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
