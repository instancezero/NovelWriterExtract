<?php

namespace Lib\NovelWriter;

use ArrayAccess;

class NovelData implements ArrayAccess
{
    const string STRUCTURE_KEYWORD = 'story';

    public array $comments = [];
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
     * Examine the content of a comment and extract anything formatted as a story
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

    public function profileText()
    {
        $this->distribution = [];
        $this->phrases = [];
        $phraseTwo = [];
        $phraseThree = [];
        $text = implode(' ', $this->text);
        $text = str_replace(["'", "’"], '', $text);
        $words = str_word_count($text, 1);
        $this->words = count($words);
        foreach ($words as $word) {
            $word = strtolower($word);
            $this->distribution[$word] ??= 0;
            ++$this->distribution[$word];
            $phraseTwo[] = $word;
            if (count($phraseTwo) > 2) {
                array_shift($phraseTwo);
                $phrase = implode(' ', $phraseTwo);
                $this->phrases[$phrase] ??= 0;
                ++$this->phrases[$phrase];
            }
            $phraseThree[] = $word;
            if (count($phraseThree) > 3) {
                array_shift($phraseThree);
                $phrase = implode(' ', $phraseThree);
                $this->phrases[$phrase] ??= 0;
                ++$this->phrases[$phrase];
            }
        }
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
        $totalSentences = 0;
        foreach ($paragraphs as $paragraph) {
            if ($paragraph === '') {
                continue;
            }
            $paragraph = preg_replace('!(dr|mr|mrs)\.!i', '$1', $paragraph);
            $sentences = explode('.', $paragraph);
            $totalSentences += count($sentences);
            $elements = [];
            $last = '';
            $run = 0;
            foreach ($sentences as $sentence) {
                $words = str_word_count($sentence);
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
            $elements[] = (($run > 1) ? $run : '') . $last;
            $pElements[]= 'P' . count($sentences) . ':' . implode('.', $elements);
        }
        $avg = count($pElements)
            ? count($pElements) . '@'
            . round($totalSentences/count($pElements), 1) . ': '
            : '';
        $this->node['_sla'] = $avg . implode(', ', $pElements);
    }

    public function unset(string $key, mixed $index)
    {
        unset($this->terms[$key][$index]);
    }

}
