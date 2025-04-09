<?php

namespace Lib\NovelWriter;

use Abivia\Criteria\Criteria;
use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
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
        'name' => 'Scene',
        'words' => 'Words',
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
    static protected array $styles = [
        '*' => ['align' => Alignment::HORIZONTAL_GENERAL, 'onFirst' => false, 'wrap' => true],
        '@' => ['align' => Alignment::HORIZONTAL_CENTER, 'onFirst' => true, 'wrap' => true],
        'comments' => ['align' => Alignment::HORIZONTAL_LEFT, 'wrap' => false],
        'duration' => ['align' => Alignment::HORIZONTAL_RIGHT],
        'time' => ['align' => Alignment::HORIZONTAL_LEFT],
        'words' => ['align' => Alignment::HORIZONTAL_RIGHT],
    ];
    protected int $wrapSize = 40;

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
            $ext = strtolower(pathinfo($path, PATHINFO_EXTENSION));
            $typeMap =match ($ext) {
                'csv' => IOFactory::WRITER_CSV,
                'html' => IOFactory::WRITER_HTML,
                'ods' => IOFactory::WRITER_ODS,
                'xlsx' => IOFactory::WRITER_XLSX,
                default => throw new Exception("Unsupported file type: $ext"),
            };
            $writer = IOFactory::createWriter($this->spreadsheet, $typeMap);
            $writer->save($path);
            $this->spreadsheet->disconnectWorksheets();
            unset($this->spreadsheet);
        } catch (Exception $exception) {
            echo $exception->getMessage();
            return;
        }
    }

    private function formatCell(int $row, int $col, array $specs): void
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
        $this->sheet->getStyle([$col, $row])->applyFromArray($style);
    }

    private function formatStyle(int $row, int $col, string $key): void
    {
        $style = $this->getStyle($key);
        $this->sheet->getStyle([$col, $row])->applyFromArray([
            'alignment' => [
                'horizontal' => $style['align'] ?? Alignment::HORIZONTAL_GENERAL,
                'vertical' => Alignment::VERTICAL_TOP,
                'wrapText' => $style['wrap'] ?? true,
            ],
        ]);
    }

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
                $scene = ['handle' => $handle, '_novel' => $name];
                $scene['words'] = isset($item->meta['wordCount'])
                    ? (string)$item->meta['wordCount'] : '';

                $this->sceneFiles[] = $scene;
            }
        }
    }

    private function loadScenes(): void
    {
        $this->inUse = ['name' => true, 'words' => true];
        $this->sceneData = [];
        $this->sceneBuffer = [];
        $this->commentBuffer = [];
        // Track if we're in the scene header or the body, so we don't accumulate inline comments.
        $inHeader = true;
        foreach ($this->sceneFiles as $scene) {
            $markdown = explode(
                "\n",
                @file_get_contents("$this->sourcePath/content/{$scene['handle']}.nwd")
            );
            // The word count is only useful in for the file, so place it in the first scene.
            $words = $scene['words'];
            foreach ($markdown as $line) {
                $line = rtrim($line);
                if (str_starts_with($line, '### ')) {
                    // This is the start of a scene, save the preceding scene, if any.
                    if (count($this->sceneBuffer)) {
                        $this->sceneBuffer['comments'] = $this->commentBuffer;
                        $this->sceneData[] = $this->sceneBuffer;
                    }
                    // Reset the header flag, invalidate the word count, and clear the comment buffer
                    $inHeader = true;
                    $this->sceneBuffer = [
                        '_novel' => $scene['_novel'],
                        'name' => trim(substr($line, 4)),
                        'words' => $words,
                    ];
                    $words = '';
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
                } elseif ($line !== '') {
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

        // Add and style the headers
        $headers = $this->getHeaders();
        $headerKeys = array_keys($headers);
        $maxColChars = [];
        $col = 1;
        foreach ($headers as $header) {
            $this->sheet->setCellValue([$col, 1], $header);
            $this->sheet->getStyle([$col, 1])->applyFromArray([
                'font' => ['bold' => true],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_TOP,
                ],
            ]);
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

        // Set column widths, save for the last one
        for ($index = 1; $index < count($headers); ++$index) {
            if (isset($maxColChars[$index])) {
                $this->sheet->getColumnDimensionByColumn($index)->setWidth(
                    min($maxColChars[$index], $this->wrapSize)
                );
            }
        }
    }

    private function prepareSheet(string $formatPath): void
    {
        $formats = json_decode(@file_get_contents($formatPath), true);
        if (empty($formats)) {
            throw new Exception("Error reading format file $formatPath\n");
        }
        $this->sheet = $this->spreadsheet->getActiveSheet();

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
            $this->sheet->setCellValue([$col, 1], $header);
            $this->sheet->getStyle([$col, 1])->applyFromArray([
                'font' => ['bold' => true],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_TOP,
                ],
            ]);
            $maxColChars[$col] = ceil(1.4 * strlen($header));
        }

        // Now add the data
        $this->seen = [];
        $criteria = new Criteria();
        $lastNovel = '';
        $sequence = 0;
        $row = 2;
        foreach ($this->sceneData as $sceneData) {
            if ($sceneData['_novel'] !== $lastNovel) {
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
                $this->formatCell($row, $col, $this->cellStyle);
                $maxColChars[$col] = max($maxColChars[$col], $this->contentLength);
                ++$col;
            }
            ++$row;
        }

        // Set column widths
        for ($index = 1; $index <= count($headers); ++$index) {
            if (isset($maxColChars[$index])) {
                $this->sheet->getColumnDimensionByColumn($index)->setWidth(
                    min($maxColChars[$index], $formats['wrap'] ?? $this->wrapSize)
                );
            }
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
