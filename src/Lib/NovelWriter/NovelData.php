<?php

namespace Lib\NovelWriter;

use ArrayAccess;
use Carbon\Carbon;
use Carbon\Exceptions\InvalidFormatException;

class NovelData implements ArrayAccess
{
    const string STRUCTURE_KEYWORD = 'story';

    public array $comments = [];
    protected static string $chronMode;
    protected static array $chronUnits =[
        '' => 1,
        'm' => 1,
        'h' => '60m',
        'd' => '24h',
        'w' => '7d',
        'mo' => '30d',
        'y' => '12mo',
    ];
    protected static array $chronVars = [];
    /**
     * @var mixed|string
     */
    protected static mixed $chronZone;
    /**
     * @var array|mixed
     */
    public array $distribution;
    public string $name {
        set {
            $this->name = trim(preg_replace('/^###!?\s+/', '', $value));
        }
    }
    /**
     * @var array|array[]|mixed
     */
    public array $phrases;
    public array $synopsis = [];
    public array $terms = [];
    private array $text = [];
    public int $words;

    public function __construct(public array $node)
    {
    }

    public function addText(string $line): void
    {
        $last = array_key_last($this->text);
        if ($line === '' && ($last === null || $this->text[$last] === '')) {
            return;
        }
        $this->text[] = $line;
    }

    public function computeChron(): void
    {
        $time = trim($this->terms['time'][0] ?? '');
        switch (self::$chronMode) {
            case 'fixed':
                try {
                    $this->node['_chron'] = Carbon::parse($time, self::$chronZone)->toIso8601String();
                } catch (InvalidFormatException) {
                    $this->node['_chron'] = '';
                }
                break;
            case 'relative':
                if ($time === '') {
                    $this->node['_chron'] = '';
                } elseif (str_contains($time, '=')) {
                    // We have a definition.
                    $parts = explode('=', $time);
                    if (!isset(self::$chronVars[$parts[0]])) {
                        self::$chronVars[$parts[0]] = self::cronAbsolute($parts[1]);
                    }
                    $this->node['_chron'] = self::$chronVars[$parts[0]];
                } else {
                    $this->node['_chron'] = self::cronAbsolute($time);
                }
                break;
            default:
                $this->node['_chron'] = '';
                break;
        }
    }

    public static function configureCron(array $cronSettings): void
    {
        self::$chronMode = strtolower($cronSettings['mode'] ?? 'off');
        switch (self::$chronMode) {
            case 'fixed':
                self::$chronZone = $cronSettings['zone'] ?? 'UTC';
                break;
            case 'relative':
                if (isset($cronSettings['units']) && is_array($cronSettings['units'])) {
                    $units = $cronSettings['units'];
                    self::$chronUnits[''] = 1;
                    foreach ($units as $unit => $interval) {
                        if (is_numeric($interval)) {
                            self::$chronUnits[$unit] = $interval;
                        } else {
                            self::$chronUnits[$unit] = self::cronAbsolute($interval);
                        }
                    }
                }
                break;
            default:
                self::$chronMode = 'off';
        }
    }

    /**
     * Convert a relative time expression to an absolute time.
     * @param string $relative
     * @return float
     */
    private static function cronAbsolute(string $relative): float
    {
        $relative = preg_replace('/\s/', '', $relative);
        $parts = preg_split('/([+\-])/', $relative, flags: PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
        $result = 0.0;
        $sign = 1;
        foreach ($parts as $part) {
            if ($part === '+') {
                $sign = 1;
            } elseif ($part === '-') {
                $sign = -1;
            } elseif (isset(self::$chronVars[$part])) {
                $result += $sign * self::$chronVars[$part];
            } else {
                $subParts = preg_split('/([0-9.]+)/', $part, flags: PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
                $increment = floatval($subParts[0]);
                $result += $sign * $increment * (self::$chronUnits[$subParts[1] ?? ''] ?? 1);
            }
        }
        return $result;
    }


    public function offsetExists(mixed $offset): bool
    {
        if (str_starts_with($offset, '_') && isset($this->node[$offset])) {
            return isset($this->node[$offset]);
        }
        if (isset($this->{$offset})) {
            return true;
        }
        if (isset($this->terms[$offset])) {
            return true;
        }
        return false;
    }

    public function offsetGet(mixed $offset): mixed
    {
        if (str_starts_with($offset, '_') && isset($this->node[$offset])) {
            return $this->node[$offset] ?? null;
        }
        if (isset($this->{$offset})) {
            return $this->{$offset};
        }
        if (isset($this->terms[$offset])) {
            return $this->terms[$offset];
        }
        return null;
    }

    public function offsetSet(mixed $offset, mixed $value): void
    {
        // TODO: Implement offsetSet() method.
    }

    public function offsetUnset(mixed $offset): void
    {
        // TODO: Implement offsetUnset() method.
    }

    /**
     * Examine the content of a comment and extract anything formatted
     * as a story into an array of terms.
     * @param string $line
     * @param array $inUse
     * @return void
     */
    public function parseComment(string $line, array &$inUse): void
    {
        $parts = explode(':', $line, 2);
        $command = strtolower(trim(substr($parts[0], 1)));
        // Handle the two versions of "synopsis".
        if (
            str_starts_with($command, 'synopsis')
            || str_starts_with($command, 'short')
        ) {
            if (count($parts) > 1) {
                $synopsis = trim($parts[1]);
                if ($synopsis !== '') {
                    $this->synopsis[] = $synopsis;
                    $inUse['synopsis'] = true;
                }
            }
        } elseif (str_starts_with($command, self::STRUCTURE_KEYWORD . '.')) {
            // This is a story attribute
            $subParts = explode('.', $command, 2);
            // Check for a story term and save the text
            if (count($subParts) > 1) {
                $term = trim($subParts[1]);
                $note = count($parts) > 1 ? trim($parts[1]) : '';
                if ($note !== '') {
                    $this->terms[$term] ??= [];
                    $this->terms[$term][] = trim($parts[1]);
                    $inUse[$term] = true;
                }
            }
        } elseif (!str_starts_with($command, '~')) {
            // Just a regular comment (in the header)
            $this->comments[] = trim(substr($line, 1));
            $inUse['comments'] = true;
        }
    }

    /**
     * Parse an at-reference in and save the value. If the value is a list, explode and trim it.
     * @param string $line
     * @param array $inUse
     * @return void
     */
    public function parseReference(string $line, array &$inUse): void
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
        $this->terms[$command] = $list;
        $inUse[$command] = true;
    }

    private function profileSentences(): void
    {
        $paragraphs = [''];
        $pNum = 0;
        foreach ($this->text as $line) {
            if ($line === '') {
                if ($paragraphs[$pNum] !== '') {
                    $paragraphs[++$pNum] = '';
                }
                continue;
            }
            $paragraphs[$pNum] .= " $line";
        }
        $pElements = [];
        $slg = '';
        $slgRef = ['▁', '▂', '▃', '▄', '▅', '▆', '▇', '█'];
        $totalSentences = 0;
        foreach ($paragraphs as $paragraph) {
            if ($paragraph === '') {
                continue;
            }
            $paragraph = preg_replace('!(dr|mr|mrs|prof)\.!i', '$1', $paragraph);
            $sentences = explode('.', $paragraph);
            $totalSentences += count($sentences);
            $elements = [];
            $last = '';
            $run = 0;
            foreach ($sentences as $sentence) {
                $words = str_word_count($sentence);
                $slg .= $slgRef[min(floor($words / 3) + 1, 7)];
                $size = match (true) {
                    $words < 5 => 's',
                    $words >= 20 => 'l',
                    default => 'm'
                };
                if ($size !== $last) {
                    if ($run) {
                        if ($run > 10) {
                            $last = strtoupper($last);
                        }
                        $elements[] = (($run > 1) ? $run : '') . $last;
                    }
                    $last = $size;
                    $run = 1;
                } else {
                    ++$run;
                }
            }
            $slg .= ' ';
            $elements[] = (($run > 1) ? $run : '') . $last;
            $pElements[]= 'P' . count($sentences) . ':' . implode('.', $elements);
        }
        $avg = count($pElements)
            ? count($pElements) . '@'
            . round($totalSentences/count($pElements), 1) . ': '
            : '';
        $this->node['_sla'] = $avg . implode(', ', $pElements);
        $this->node['_slg'] = $slg;
    }

    /**
     * Compute the frequency of words and two or three character phrases in the scene.
     * @return void
     */
    public function profileText(): void
    {
        $this->distribution = [];
        $this->phrases = [];
        $phraseTwo = [];
        $phraseThree = [];
        $text = implode(' ', $this->text);
        $text = str_replace(["'", "’"], '', $text);
        $words = str_word_count($text, 1, '-');
        $this->words = count($words);
        foreach ($words as $offset => $word) {
            $word = strtolower($word);

            // Single word count
            $this->distribution[$word] ??= ['count' => 0, 'offsets' => []];
            ++$this->distribution[$word]['count'];
            $this->distribution[$word]['offsets'][] = $offset;

            // Two-word phrase count
            $phraseTwo[] = $word;
            if (count($phraseTwo) > 2) {
                array_shift($phraseTwo);
                $phrase = implode(' ', $phraseTwo);
                $this->phrases[$phrase] ??= ['count' => 0, 'offsets' => []];
                ++$this->phrases[$phrase]['count'];
                $this->phrases[$phrase]['offsets'][] = $offset;
            }
            $phraseThree[] = $word;
            if (count($phraseThree) > 3) {
                array_shift($phraseThree);
                $phrase = implode(' ', $phraseThree);
                $this->phrases[$phrase] ??= ['count' => 0, 'offsets' => []];
                ++$this->phrases[$phrase]['count'];
                $this->phrases[$phrase]['offsets'][] = $offset;
            }
        }
        $this->computeClumpiness($this->distribution);
        $this->computeClumpiness($this->phrases);
        $this->profileSentences();
    }

    private function computeClumpiness(array &$distribution): void
    {
        foreach ($distribution as &$info) {
            $histogram = [];
            for ($index = 0; $index < count($info['offsets']) - 1; ++$index) {
                $delta = $info['offsets'][$index + 1] - $info['offsets'][$index];
                $histogram[$delta] ??= 0;
                ++$histogram[$delta];
            }
            $clumpiness = 0.0;
            foreach ($histogram as $delta => $count) {
                $clumpiness += $count / (1.0 + log($delta));
            }
            $info['clumpiness'] = round(100.0 * $clumpiness, 1);
            unset($info['offsets']);
        }
    }

    public function unset(string $key, mixed $index): void
    {
        unset($this->terms[$key][$index]);
    }

}
