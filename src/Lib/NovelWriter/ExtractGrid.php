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
    const float BOLD_FACTOR = 1.2;
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
    /**
     * @var array|float[]
     */
    protected array $columnWidth;
    protected array $contentList;
    /**
     * @var array|mixed|string|string[]|null
     */
    protected string $contentString;
    /**
     * @var float
     */
    protected float $contentWidth;
    protected bool $cronModeFixed = true;
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
        '_cron' => 'Cron Time',
        '_folder' => 'Folder',
        '_novel' => 'Novel',
        '_sequence' => '#',
        '_sla' => 'Sentence Lengths',
        '_status' => 'Status',
        'about' => 'This NovelData is About',
        'aural' => 'Environmental Sounds',
        'comments' => 'Additional Notes',
        'complication' => 'Complication(s)',
        'impact' => 'Impact of the scene',
        'incite' => 'Inciting Incident',
        'others' => 'Off-stage Characters',
        'polarity' => 'Polarity Shift',
        'prose' => 'Prose Quality/Cadence',
        'resolution' => '(Non-)Resolution',
        'smell' => 'Environmental Smells',
        'time' => 'Period/Time',
        'tod' => 'Time of Day',
        'touch' => 'Tactile',
        'turning' => 'Turning Point',
        'value' => 'Value Shift',
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
     * @var false|mixed
     */
    protected bool $onFirst;
    protected float $phraseColWidth;
    /**
     * @var array|array[]
     */
    protected array $phrases;
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
        '_cron',
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
    /**
     * @var array[NovelData]
     */
    protected array $sceneData = [];
    protected array $sceneFiles = [];
    protected NovelWriterFileLoader $sceneLoader;
    protected array $seen = [];
    protected array $seenGlobal = [];
    private int $sheetIndex = 0;
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

    /**
     * Make sure we have a valid output type.
     * @param string $path
     * @return array
     */
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
     * Check the path to ensure a novelWriter project exists there.
     * @return bool
     */
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
     * Reset the variables associated with building a cell.
     * @return void
     */
    private function clearCellData(): void
    {
        $this->contentString = '';
        $this->contentList = [];
        $this->contentWidth = 0.0;
        $this->cellStyle = [];
        $this->onFirst = false;
    }

    private function computeCron()
    {
        $timeFormat = $this->formats['time'] ?? [];
        if (!is_array($timeFormat)) {
            return;
        }
        NovelData::configureCron($timeFormat);
        /** @var NovelData $scene */
        foreach ($this->sceneData as $scene) {
            $scene->computeCron();
        }
    }

    /**
     * Get word counts by scene status, active state.
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

            // Same with "wider" (and "much wider")
            $wider = 0.8 * strlen(preg_replace('/[^MOQW]/', '', $line))
                + 1.1 * strlen(preg_replace('/[^*]/', '', $line));

            // Same with "narrower" which gets subtracted.
            $narrower = 0.7 * strlen(preg_replace('/[^iltI|)(}{ !\-\'.;:`]/', '', $line));
            $width = max($width, strlen($line) + $wide + $wider - $narrower);
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
            $this->write($path);
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
     * @param bool $useGlobal
     * @return void
     */
    private function flagFirst(string $key, bool $useGlobal = false): void
    {
        if ($this->cellStyle['onFirst'] ?? false) {
            $this->seen[$key] ??= [];
            $this->seenGlobal ??= [];
            $hasNewValue = false;
            $newItems = [];
            $oldItems = [];
            foreach ($this->contentList as $slot => $newValue) {
                $isNew = false;
                if ($useGlobal) {
                    if (!in_array($newValue, $this->seenGlobal)) {
                        $isNew = true;
                    }
                } elseif (!in_array($newValue, $this->seen[$key])) {
                    $isNew = true;
                }
                if ($isNew) {
                    $newItems[] = $newValue;
                    $this->seen[$key][] = $newValue;
                    $this->seenGlobal[] = $newValue;
                    $hasNewValue = true;
                } else {
                    $oldItems[] = "($newValue)";
                }
            }
            if (count($oldItems) !== count($this->contentList)) {
                $this->contentList = array_merge($newItems, $oldItems);
                $this->contentString = implode("\n", $this->contentList);
                $this->contentWidth = $this->estimateWidth($this->contentString, $hasNewValue);
            }
            if ($hasNewValue) {
                $this->cellStyle['bold'] = true;
            }
        }
    }

    /**
     * Set the formatting for a cell.
     * @param Worksheet $sheet
     * @param int $row
     * @param int $col
     * @param array $specs
     * @return void
     */
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
     * Get a list of scenes that reference the named character.
     * @param string $character
     * @return array
     */
    private function getCharacterScenes(string $character): array
    {
        $character = strtolower($character);
        $nodeList = [];
        foreach ($this->sceneData as $node) {
            // if the node references the character...
            $sceneCharacters = $node['@char'];
            if ($sceneCharacters) {
                foreach ($sceneCharacters as $sceneCharacter) {
                    if ($character === strtolower($sceneCharacter)) {
                        $nodeList[] = $node;
                        break;
                    }
                }
            }
        }
        return $nodeList;
    }

    /**
     * Get a user provided column list, the default list, or false for no columns.
     * @param mixed $option
     * @param array $default
     * @return array|false
     */
    private function getColumns(mixed $option, array $default): array|false
    {
        // If the column specification is 'true', then include all columns.
        if ($option === true) {
            $option = $default;
        }
        return $option;
    }

    /**
     * Prepare an array of header texts from a list of column keys
     * @param array $columns
     * @param array $overrides
     * @return string[]
     */
    private function getHeaders(array $columns, array $overrides = []): array
    {
        $headers = [];
        foreach ($columns as $column) {
            if (is_array($column) && ($column['heading'] ?? false)) {
                $headers[] = $column['heading'];
            } elseif (is_string($column)) {
                if ($overrides[$column] ?? false) {
                    $headers[] = $overrides[$column];
                } else {
                    if ((self::$headerText[$column] ?? false)) {
                        $headers[] = self::$headerText[$column];
                    } else {
                        $words = explode('_', $column);
                        $words = array_map(fn($word): string => ucfirst($word), $words);
                        $headers[] = implode(' ', $words);
                    }
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
            $data = array_unique($data);
            sort($data);
            $this->contentWidth = 0;
            $this->contentList = $data;
            $delimiter = ($column[0] === '@') ? "\n" : "\n\n";
            $this->contentString = implode($delimiter, $data);
        } else {
            $this->contentList = [$data];
            $this->contentString = preg_replace('! +!', ' ', $data);
        }
        $this->contentWidth = $this->estimateWidth($this->contentString);
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
     * Use the output path to determine the type of writer to use.
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
     * Get the contents of a novelWriter data file as a NodeData object.
     * @param array $file
     * @return NovelData
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

    /**
     * Load a novelWriter file in the locations tree.
     * @return void
     */
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

    /**
     * Load the novelWriter scene files.
     * @return void
     */
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

    /**
     * Generate the story word frequency analysis sheet.
     * @param bool $newSheet
     * @return bool
     */
    private function prepareAnalysis(bool $newSheet): bool
    {
        if (count($this->wordDistribution) === 0 && count($this->phrases) === 0) {
            return $newSheet;
        }
        if ($this->verbose) {
            echo "Preparing Analysis\n";
        }
        if ($newSheet) {
            $this->spreadsheet->createSheet();
            ++$this->sheetIndex;
        }
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Analysis');

        $fw = $this->estimateWidth('Frequency');
        $this->setHeader($sheet, 1, 'Word', width: 6.0);
        $this->setHeader($sheet, 2, 'Frequency', width: $fw);
        $this->setHeader($sheet, 3, 'Clumpiness', width: 8.0);
        $this->setHeader($sheet, 4, 'Avg. Clump', width: 8.0);
        $this->setHeader($sheet, 5, '', width: 1.0);
        $this->setHeader($sheet, 6, 'Phrase', width: $this->phraseColWidth);
        $this->setHeader($sheet, 7, 'Frequency', width: $fw);
        $this->setHeader($sheet, 8, 'Clumpiness', width: 8.0);
        $this->setHeader($sheet, 9, 'Avg. Clump', width: 8.0);

        $sheet->getColumnDimensionByColumn(1)->setWidth($this->wordDistributionWidth);
        $sheet->getColumnDimensionByColumn(5)->setWidth($this->phraseColWidth);

        $row = 2;
        $right = $this->wordCountStyle;
        // Word frequency
        foreach ($this->wordDistribution as $word => $stats) {
            // Status
            $sheet->setCellValue([1, $row], $word);
            $sheet->setCellValue([2, $row], $stats['count']);
            $this->formatCell($sheet, $row, 2, $right);
            $sheet->setCellValue([3, $row], $stats['clumpiness']);
            $this->formatCell($sheet, $row, 3, $right);
            $sheet->setCellValue([4, $row], round($stats['clumpiness']/$stats['count'], 2));
            $this->formatCell($sheet, $row, 4, $right);
            ++$row;
        }
        $row = 2;
        // Status
        foreach ($this->phrases as $word => $stats) {
            // Status
            $sheet->setCellValue([6, $row], $word);
            $sheet->setCellValue([7, $row], $stats['count']);
            $this->formatCell($sheet, $row, 7, $right);
            $sheet->setCellValue([8, $row], $stats['clumpiness']);
            $this->formatCell($sheet, $row, 8, $right);
            $sheet->setCellValue([9, $row], round($stats['clumpiness']/$stats['count'], 2));
            $this->formatCell($sheet, $row, 9, $right);
            ++$row;
        }
        $sheet->setSelectedCell('A2');

        return true;
    }

    /**
     * Generate the character sheet.
     * @param bool $newSheet
     * @return bool
     * @throws LogicException
     */
    private function prepareCharacters(bool $newSheet): bool
    {
        $columns = $this->formats['characters'] ?? true;
        if ($columns === false || count($this->characterData) === 0) {
            return $newSheet;
        }
        if ($this->verbose) {
            echo "Preparing Characters\n";
        }
        if ($newSheet) {
            $this->spreadsheet->createSheet();
            ++$this->sheetIndex;
            $this->columnWidth = [];
        }
        $columns = $this->getColumns($columns, $this->characterAttributes);
        // Add and style the headers
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Characters');

        $headers = $this->setHeaders($sheet, $columns);

        // Now add the data
        $this->prepareColumns($sheet, $this->characterData, $columns, $headers);

        $this->setColumnWidths($sheet, $headers);
        $sheet->setSelectedCell('A2');

        return true;
    }

    /**
     * Get a column with customization
     * @param array $column
     * @param NovelData $node
     * @return void
     */
    private function prepareColumnCustom(array $column, NovelData $node): void
    {
        $nodeCopy = clone $node;
        // Check for renamed header and/or filtered data
        if (($column['exclude'] ?? false) && ($nodeCopy[$column['key']] ?? false)) {
            if (is_array($nodeCopy[$column['key']])) {
                foreach ($nodeCopy[$column['key']] as $index => $value) {
                    if (in_array($value, $column['exclude'])) {
                        $nodeCopy->unset($column['key'], $index);
                    }
                }
            } elseif (in_array($nodeCopy[$column['key']], $column['exclude'])) {
                $nodeCopy[$column['key']] = '';
            }
        }
        $this->getNodeData($nodeCopy, $column['key']);
        $this->setCellStyle($column);
    }

    /**
     * Get a string column with a fallback to a second column if the primary is absent.
     * @param NovelData $node
     * @param int $sequence
     * @param string $columnName
     * @param string $fallbackName
     * @return void
     */
    private function prepareColumnFallback(
        NovelData $node,
        int $sequence,
        string $columnName,
        string $fallbackName
    ): void
    {
        switch ($columnName) {
            case '_blank':
                break;
            case '_sequence':
                $this->contentString = (string)$sequence;
                $this->contentWidth = $this->estimateWidth($sequence);
                $this->cellStyle['align'] = Alignment::HORIZONTAL_RIGHT;
                break;
            default:
                if (isset($node[$columnName])) {
                    $this->getNodeData($node, $columnName);
                } else {
                    $this->getNodeData($node, $fallbackName);
                }
                break;
        }
        $this->setCellStyle(['key' => $columnName]);
    }

    /**
     * Get a column with a conditional/indirection.
     * @param array $column
     * @param NovelData $sceneData
     * @return void
     */
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
     * Get a column when the column specification is just a string.
     * @param NovelData $node
     * @param int $sequence
     * @param string $columnName
     * @return void
     */
    private function prepareColumnSimple(NovelData $node, int $sequence, string $columnName): void
    {
        switch ($columnName) {
            case '_blank':
                break;
            case '_sequence':
                $this->contentString = (string)$sequence;
                $this->contentWidth = $this->estimateWidth($sequence);
                $this->cellStyle['align'] = Alignment::HORIZONTAL_RIGHT;
                break;
            default:
                $this->getNodeData($node, $columnName);
                break;
        }
        $this->setCellStyle(['key' => $columnName]);
    }

    /**
     * @param Worksheet $sheet
     * @param NovelData[] $nodes
     * @param array $columns
     * @param array $headers
     * @return int
     * @throws LogicException
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
                $this->clearCellData();
                if (is_string($columnSpecification)) {
                    $this->prepareColumnSimple($node, $sequence, $columnSpecification);
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
                    $this->prepareColumnCustom($columnSpecification, $node);
                }
                $this->flagFirst($seenKey, $columnSpecification === '@mention');
                //++$this->contentWidth;
                $sheet->setCellValue([$col1, $row], $this->contentString);
                $this->formatCell($sheet, $row, $col1, $this->cellStyle);
                $this->columnWidth[$col1] = max($this->columnWidth[$col1], $this->contentWidth);
                ++$col0;
            }
            ++$row;
        }

        return $row;
    }

    /**
     * Generate the locations sheet.
     * @param bool $newSheet
     * @return bool
     * @throws LogicException
     */
    private function prepareLocations(bool $newSheet): bool
    {
        $columns = $this->formats['locations'] ?? true;
        if ($columns === false || count($this->locationData) === 0) {
            return $newSheet;
        }
        if ($this->verbose) {
            echo "Preparing Locations\n";
        }
        if ($newSheet) {
            $this->spreadsheet->createSheet();
            ++$this->sheetIndex;
            $this->columnWidth = [];
        }
        $columns = $this->getColumns($columns, $this->locationAttributes);
        // Add and style the headers
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Locations');

        $headers = $this->setHeaders($sheet, $columns);

        // Now add the data
        $this->prepareColumns($sheet, $this->locationData, $columns, $headers);

        $this->setColumnWidths($sheet, $headers);
        $sheet->setSelectedCell('A2');

        return true;
    }

    /**
     * Generate a scenes sheet
     * @throws LogicException
     */
    private function prepareScenes(): bool
    {
        $columns = $this->formats['scenes'] ?? ($this->formats['columns'] ?? true);
        if ($columns === false || count($this->sceneData) === 0) {
            return false;
        }
        if ($this->verbose) {
            echo "Preparing Scenes\n";
        }
        $this->columnWidth = [];
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
        $sheet->setSelectedCell('A2');

        return true;
    }

    /**
     * Generate all sheets.
     * @param string $formatPath
     * @return void
     * @throws LogicException
     * @throws Exception
     */
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
        $this->computeCron();
        $this->wordCountStyle = self::$styles['words'];
        $hadContent = $this->prepareScenes();
        if ($this->formats['wordCounts'] ?? true) {
            $hadContent = $this->prepareWordCounts($hadContent);
        }
        if ($this->formats['characters'] ?? true) {
            $hadContent = $this->prepareCharacters($hadContent);
        }
        if ($this->formats['locations'] ?? true) {
            $hadContent = $this->prepareLocations($hadContent);
        }
        $timelines = $this->formats['timelines'] ?? false;
        if ($timelines !== false) {
            $hadContent = $this->prepareTimelines($hadContent, $timelines);
        }
        if ($this->formats['analysis'] ?? true) {
            $hadContent = $this->prepareAnalysis($hadContent);
        }
    }

    /**
     * Generate a timeline sheet for the named character.
     * @param bool $newSheet
     * @param array $nodeList
     * @param string $character
     * @return void
     */
    private function prepareTimeline(
        bool $newSheet,
        array $nodeList,
        string $character,
        array $optionalColumns
    ): void
    {
        if ($newSheet) {
            $this->spreadsheet->createSheet();
            ++$this->sheetIndex;
            $this->columnWidth = [];
        }
        if ($this->verbose) {
            echo "Preparing timeline for $character\n";
        }
        $columns = array_merge(['_sequence', 'name'], $optionalColumns, ['synopsis']);
        $columns = $this->getColumns($columns, $this->characterAttributes);
        // Add and style the headers
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle($character);

        $headers = $this->setHeaders(
            $sheet,
            $columns,
            ['time' => 'Time','name' => 'Scene', 'synopsis' => 'Storyline/Synopsis']
        );
        $storyLineKey = strtolower("of_$character");

        // Now add the data
        $lastNovel = '';
        $sequence = 0;
        $row = 2;
        foreach ($nodeList as $node) {
            if (($node['_novel'] ?? '') !== $lastNovel) {
                $sequence = 0;
                $lastNovel = $node['_novel'];
            }
            ++$sequence;
            $col0 = 0;
            foreach ($columns as $columnSpecification) {
                $col1 = $col0 + 1;
                $this->clearCellData();
                if ($columnSpecification === 'synopsis') {
                    $this->prepareColumnFallback($node, $sequence, $storyLineKey, 'synopsis');
                } else {
                    $this->prepareColumnSimple($node, $sequence, $columnSpecification);
                }
                ++$this->contentWidth;
                $sheet->setCellValue([$col1, $row], $this->contentString);
                $this->formatCell($sheet, $row, $col1, $this->cellStyle);
                $this->columnWidth[$col1] = max($this->columnWidth[$col1], $this->contentWidth);
                ++$col0;
            }
            ++$row;
        }

        $this->setColumnWidths($sheet, $headers);
        $sheet->setSelectedCell('A2');
    }

    /**
     * Determine which timelines to prepare and generate them.
     * @param bool $hadContent
     * @param mixed $timelines
     * @return bool
     *
     *  Timelines:
     *  false: disabled.
     *  int: number of scenes the character has to be in to get a report
     *  array:
     *  [chars] => list of characters to generate timeline for
     *  [minimum] Minimum number of scenes required to generate a timeline
     *  [show] array for optional columns:
     *      [time] => report the time from the scene
     *      [_cron] report the "cron time" for the scene
     */
    private function prepareTimelines(bool $hadContent, mixed $timelines): bool
    {
        $timelineChars = null;
        $optionalColumns = [];
        if (is_array($timelines)) {
            $timeLineFloor = $timelines['minimum'] ?? 1;
            if (isset($timelines['chars'])) {
                $timelineChars = $timelines['chars'];
            }
            if (isset($timelines['show'])) {
                if (isset($timelines['show']['time']) && $timelines['show']['time']) {
                    $optionalColumns[] = 'time';
                }
                if (isset($timelines['show']['_cron']) && $timelines['show']['_cron']) {
                    $optionalColumns[] = '_cron';
                }
            } else {
                $optionalColumns = ['time', '_cron'];
            }

        } else {
            // We just expect a scene threshold.
            $timeLineFloor = is_numeric($timelines) ? $timelines : 1;
        }
        // If no characters were specified, get all characters
        if ($timelineChars === null) {
            $timelineChars = [];
            foreach ($this->characterData as $characterNode) {
                if ($characterNode['@tag']) {
                    $timelineChars[] = $characterNode['@tag'][0];
                }
            }
        }
        sort($timelineChars);
        $newHadContent = false;
        $skipped = [];
        foreach ($timelineChars as $character) {
            $nodeList = $this->getCharacterScenes($character);
            if (count($nodeList) >= $timeLineFloor) {
                $this->prepareTimeline($hadContent, $nodeList, $character, $optionalColumns);
                $hadContent = true;
                $newHadContent = true;
            } else {
                $skipped[] = $character;
            }
        }
        if (count($skipped) > 0) {
            echo "Skip timeline for < $timeLineFloor scenes: " . implode(', ', $skipped) . "\n";
        }
        return $newHadContent;
    }

    /**
     * Generate the statistics sheet.
     * @param bool $newSheet
     * @return bool
     */
    private function prepareWordCounts(bool $newSheet): bool
    {
        if (count($this->sceneData) === 0) {
            return $newSheet;
        }
        if ($this->verbose) {
            echo "Preparing Statistics\n";
        }
        if ($newSheet) {
            $this->spreadsheet->createSheet();
            ++$this->sheetIndex;
            $this->columnWidth = [];
        }
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex);
        $sheet->setTitle('Statistics');
        $statusWidth = $this->estimateWidth('Status');
        $statusList = array_keys($this->sceneLoader->sceneStatusList);
        sort($statusList);
        foreach ($statusList as $status) {
            $statusWidth = max($statusWidth, $this->estimateWidth($status));
        }
        $sheet->getColumnDimensionByColumn(1)->setWidth($statusWidth);
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
        $sheet->setSelectedCell('A2');

        return true;
    }

    /**
     * Profile the scenes to get word and phrase repetition counts.
     * @return void
     */
    private function profileScenes(): void
    {
        $this->phrases = [];
        $this->phraseColWidth = 5;
        $this->wordDistribution = [];
        $this->wordDistributionWidth = 0;
        $this->wordTotal = 0;
        foreach ($this->sceneData as $scene) {
            $this->wordTotal += $scene->words;
            foreach ($scene->distribution as $word => $frequency) {
                $this->wordDistribution[$word] ??= ['count' => 0, 'clumpiness' => 0];
                $this->wordDistribution[$word]['count'] += $frequency['count'];
                $this->wordDistribution[$word]['clumpiness'] = max(
                    $this->wordDistribution[$word]['clumpiness'], $frequency['clumpiness']
                );
            }
            foreach ($scene->phrases as $phrase => $frequency) {
                $this->phrases[$phrase] ??= ['count' => 0, 'clumpiness' => 0];
                $this->phrases[$phrase]['count'] += $frequency['count'];
                $this->phrases[$phrase]['clumpiness'] = max(
                    $this->phrases[$phrase]['clumpiness'], $frequency['clumpiness']
                );
            }
        }
        foreach ($this->wordDistribution as $word => $frequency) {
            if ($frequency['count'] < 10 && $frequency['clumpiness'] < 20.0) {
                unset($this->wordDistribution[$word]);
                continue;
            }
            $this->wordDistributionWidth = max(
                $this->wordDistributionWidth, $this->estimateWidth($word)
            );
        }
        arsort($this->wordDistribution);
        foreach ($this->phrases as $word => $frequency) {
            if ($frequency['count'] < 10 && $frequency['clumpiness'] < 10.0) {
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
     * @return void
     */
    private function setCellStyle(array $column): void
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
    }

    /**
     * Set the width of a column.
     * @param Worksheet $sheet
     * @param int $col
     * @param float $width
     * @return void
     */
    private function setColumnWidth(Worksheet $sheet, int $col, float $width): void
    {
        $sheet->getColumnDimensionByColumn($col)->setWidth($width);
    }

    private function setColumnWidths(Worksheet $sheet, array $headers): void
    {
        // Set column widths
        for ($index = 1; $index <= count($headers); ++$index) {
            if (isset($this->columnWidth[$index])) {
                $this->setColumnWidth(
                    $sheet,
                    $index,
                    min($this->columnWidth[$index], $this->formats['wrap'] ?? $this->wrapSize)
                );
            }
        }
    }

    /**
     * Set a header cell.
     * @param Worksheet $sheet
     * @param int $col
     * @param string $header
     * @param int $row
     * @param float $width
     * @return void
     */
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
            $this->setColumnWidth($sheet, $col, $this->estimateWidth($header, true));
        }
    }

    /**
     * Set headers for a list of columns.
     * @param Worksheet $sheet
     * @param array $columns
     * @param array $overrides
     * @return string[]
     */
    private function setHeaders(Worksheet $sheet, array $columns, array $overrides = []): array
    {
        // Add and style the headers
        $headers = $this->getHeaders($columns, $overrides);

        $this->columnWidth = [];
        foreach ($headers as $col0 => $header) {
            $col1 = $col0 + 1;
            $this->setHeader($sheet, $col1, $header);
            $this->columnWidth[$col1] = $this->estimateWidth($header, true);
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

    /**
     * Write all sheets to the output path.
     * @param string $path
     * @return void
     * @throws Exception
     */
    private function write(string $path): void
    {
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
    }

}
