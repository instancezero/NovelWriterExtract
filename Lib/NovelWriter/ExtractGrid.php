<?php

namespace Lib\NovelWriter;

use \Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use SimpleXMLElement;

class ExtractGrid
{
    const STRUCTURE_KEYWORD = 'story';

    static protected array $alignment = [
        'words' => Alignment::HORIZONTAL_RIGHT,
        'time' => Alignment::HORIZONTAL_RIGHT,
        'duration' => Alignment::HORIZONTAL_RIGHT,
        '@location' => Alignment::HORIZONTAL_CENTER,
        '@char' => Alignment::HORIZONTAL_CENTER,
    ];
    protected array $commentBuffer;
    static protected array $headings = [
        'name' => 'Scene',
        'words' => 'Word Count',
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
        'pov' => 'Point of View',
        'time' => 'Period/Time',
        'duration' => 'Duration',
        '@location' => 'Location(s)',
        '@char' => 'Characters',
        'others' => 'Off-stage Characters',
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
    protected SimpleXMLElement $project;
    protected array $sceneBuffer;
    protected array $sceneData = [];
    protected array $sceneFiles = [];
    protected string $sourcePath;
    protected Spreadsheet $spreadsheet;
    protected int $wrapSize = 40;

    public function export(string $path)
    {
        $ext = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        switch ($ext) {
            case 'csv':
                $this->exportCsv($path);
                break;
            case 'ods':
                $this->exportOds($path);
                break;
            case 'xlsx':
                $this->exportXlsx($path);
                break;
            default:
                throw new Exception("Unsupported file type: $ext");
        }
    }

    public function exportCsv(string $csvPath)
    {
        $this->loadProject();
        $this->loadScenes();
        $csvFile = fopen($csvPath, 'w');
        if ($csvFile === false) {
            throw new Exception("$csvPath failed to open.");
        }
        $headers = $this->getHeaders();
        $headerKeys = array_keys($headers);
        fputcsv($csvFile, $headers);
        foreach ($this->sceneData as $sceneData) {
            $buffer = [];
            foreach ($headerKeys as $key) {
                if (isset($sceneData[$key])) {
                    if (is_array($sceneData[$key])) {
                        $buffer[$key] = implode("\n", $sceneData[$key]);
                    } else {
                        $normalized = preg_replace('!\s+!', ' ', $sceneData[$key]);
                        $buffer[$key] = $normalized;
                    }
                } else {
                    $buffer[$key] = '';
                }
            }
            fputcsv($csvFile, $buffer);
        }
        fclose($csvFile);
    }

    private function exportOds(string $odsPath)
    {
        $this->prepareSheet($odsPath);
        $writer = IOFactory::createWriter($this->spreadsheet, IOFactory::WRITER_ODS);
        $writer->save($odsPath);
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);
    }

    private function exportXlsx(string $xlPath)
    {
        $this->prepareSheet($xlPath);
        $writer = IOFactory::createWriter($this->spreadsheet, IOFactory::WRITER_XLSX);
        $writer->save($xlPath);
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);
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

    private function loadProject()
    {
        $this->sceneFiles = [];
        $this->project = new SimpleXMLElement(file_get_contents("$this->sourcePath/nwProject.nwx"));
        $parent = '';
        foreach ($this->project->content->item as $item) {
            if ((string)$item['class'] !== 'NOVEL') {
                continue;
            }
            $itemType = (string)$item['type'];
            $handle = (string)$item['handle'];
            $root = (string)$item['root'];
            if ($itemType === 'ROOT') {
                $parent = $handle;
            } elseif ($itemType === 'FILE' && $root === $parent) {
                $this->sceneFiles[] = ['handle' => $handle, 'words' => (string)$item->meta['wordCount']];
            }
        }
    }

    private function loadScenes()
    {
        $this->inUse = ['name' => true, 'words' => true];
        $this->sceneData = [];
        $this->sceneBuffer = [];
        $this->commentBuffer = [];
        // Track if we're in the scene header or the body, so we don't accumulate inline comments.
        $inHeader = true;
        foreach ($this->sceneFiles as $scene) {
            $markdown = explode("\n", file_get_contents("$this->sourcePath/content/{$scene['handle']}.nwd"));
            $words = $scene['words'];
            foreach ($markdown as $line) {
                $line = rtrim($line);
                if (str_starts_with($line, '### ')) {
                    if (count($this->sceneBuffer)) {
                        $this->sceneBuffer['comments'] = $this->commentBuffer;
                        $this->sceneData[] = $this->sceneBuffer;
                    }
                    $inHeader = true;
                    $this->sceneBuffer = [
                        'name' => trim(substr($line, 4)),
                        'words' => $words,
                    ];
                    $words = '';
                    $this->commentBuffer = [];
                } elseif (str_starts_with($line, '%')) {
                    if (str_starts_with($line, '%%')) {
                        $inHeader = true;
                    } elseif ($inHeader) {
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

    private function parseComment(string $line)
    {
        $parts = explode(':', $line, 2);
        $command = strtolower(trim(substr($parts[0], 1)));
        if (str_starts_with($command, 'synopsis')) {
            $this->sceneBuffer['synopsis'] = trim($parts[1]);
            $this->inUse['synopsis'] = true;
        } elseif (str_starts_with($command, self::STRUCTURE_KEYWORD . '.')) {
            $term = trim(substr($command, 6));
            $note = count($parts) > 1 ? trim($parts[1]) : '';
            if ($note !== '') {
                $this->sceneBuffer[$term] = trim($parts[1]);
                $this->inUse[$term] = true;
            }
        } else {
            $this->commentBuffer[] = trim(substr($line, 1));
            $this->inUse['comments'] = true;
        }
    }

    private function parseReference(string $line)
    {
        $parts = explode(':', $line, 2);
        $command = strtolower(trim($parts[0]));
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

    public function prepareSheet(string $path)
    {
        $this->loadProject();
        $this->loadScenes();
        $this->spreadsheet = new Spreadsheet();
        $sheet = $this->spreadsheet->getActiveSheet();

        // Add and style the headers
        $headers = $this->getHeaders();
        $headerKeys = array_keys($headers);
        $maxColChars = [];
        $col = 1;
        foreach ($headers as $header) {
            $sheet->setCellValue([$col, 1], $header);
            $sheet->getStyle([$col, 1])->applyFromArray([
                'font' => ['bold' => true],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_TOP,
                ],
            ]);
            $maxColChars[$col] = ceil(1.4 *strlen($header));
            ++$col;
        }

        // Now add the data
        $row = 2;
        foreach ($this->sceneData as $sceneData) {
            $col = 1;
            foreach ($headerKeys as $key) {
                if (isset($sceneData[$key])) {
                    if (is_array($sceneData[$key])) {
                        $length = 0;
                        foreach ($sceneData[$key] as $item) {
                            $length = max($length, strlen($item));
                        }
                        $contents = implode("\n", $sceneData[$key]);
                    } else {
                        $contents = preg_replace('!\s+!', ' ', $sceneData[$key]);
                        $length = strlen($sceneData[$key]);
                    }
                    $sheet->setCellValue([$col, $row], $contents);
                    $wrap = $key !== 'comments';
                    $sheet->getStyle([$col, $row])->applyFromArray([
                        'alignment' => [
                            'horizontal' => self::$alignment[$key] ?? Alignment::HORIZONTAL_GENERAL,
                            'vertical' => Alignment::VERTICAL_TOP,
                            'wrapText' => $wrap,
                        ],
                    ]);
                    $maxColChars[$col] = max($maxColChars[$col], $length);
                }
                ++$col;
            }
            ++$row;
        }

        // Set column widths, save for the last one
        for ($index = 1; $index < count($headers); ++$index) {
            if (isset($maxColChars[$index])) {
                $sheet->getColumnDimensionByColumn($index)->setWidth(min($maxColChars[$index], $this->wrapSize));
            }
        }
    }

    public function setSourcePath(string $path)
    {
        $this->sourcePath = $path;
    }

}
