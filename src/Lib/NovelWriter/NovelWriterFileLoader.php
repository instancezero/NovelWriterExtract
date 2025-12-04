<?php

namespace Lib\NovelWriter;

class NovelWriterFileLoader
{
    public array $inUse = [];
    public array $sceneStatusList = [];

    private function isSceneHeader(string $line): bool
    {
        return str_starts_with($line, '### ')
            || str_starts_with($line, '###! ');
    }

    /**
     * Parse the lines in a novelWriter file into a note node.
     *
     * @param array $node
     * @param array $markdown
     * @param array $inUse
     * @return NovelData
     */
    public function loadFile(array $node, array $markdown, array &$inUse): NovelData
    {
        $data = new NovelData($node);
        $data->name = $node['name'];

        foreach ($markdown as $line) {
            $line = rtrim($line);
            if ($line === '') {
                continue;
            }
            if (str_starts_with($line, '%')) {
                if (!str_starts_with($line, '%%')) {
                    // Look for a story extension
                    $data->parseComment($line, $inUse);
                }
            } elseif (str_starts_with($line, '@')) {
                $data->parseReference($line, $inUse);
            }
        }

        return $data;
    }

    /**
     * Parse the lines in a novelWriter file into Scenes.
     *
     * @param array $node
     * @param array $markdown
     * @param array $sceneData
     * @return void
     */
    public function loadScene(array $node, array $markdown, array &$sceneData): void
    {
        $this->inUse['_sl'] = true;
        $inHeader = true;
        foreach ($markdown as $line) {
            $line = rtrim($line);
            if ($line === '') {
                //continue;
            }
            if ($this->isSceneHeader($line)) {
                // This is the start of a scene, save the preceding scene, if any.
                if (isset($scene)) {
                    $scene->profileText();
                    $sceneData[] = $scene;
                }
                // Reset the header flag, invalidate the word count, and clear the comment buffer
                $inHeader = true;
                $scene = new NovelData($node);
                $this->sceneStatusList[$node['_status']] = true;
                $scene->name = $line;
            } elseif (isset($scene)) {
                if (str_starts_with($line, '%')) {
                    if (str_starts_with($line, '%%')) {
                        // We're in the header metadata
                        $inHeader = true;
                    } elseif ($inHeader) {
                        // Look for a story extension
                        $scene->parseComment($line, $this->inUse);
                    }
                } elseif (str_starts_with($line, '@')) {
                    $scene->parseReference($line, $this->inUse);
                } else {
                    $inHeader = $line === '';
                    if (preg_match('!^[%@#[]!', $line)) {
                        continue;
                    }
                    $scene->addText($line);
                }
            }
        }
        if (isset($scene)) {
            $scene->profileText();
            $sceneData[] = $scene;
        }
    }
}
