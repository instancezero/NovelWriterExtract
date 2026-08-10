<?php

namespace Unit;

use Lib\NovelWriter\NovelData;
use Lib\NovelWriter\NovelWriterFileLoader;
use PHPUnit\Framework\TestCase;

class NovelWriterFileLoaderTest extends TestCase
{
    public function testLoadSceneFormat1()
    {
        $lines = explode("\n", file_get_contents(__DIR__ . '/../testData/bc0cbd2a407f3.nwd'));
        $loader = new NovelWriterFileLoader();
        $scenes = [];
        $loader->loadScene(['_status' => 'foo'], $lines, $scenes);
        $this->assertCount(3, $scenes);
        /** @var NovelData $scene */
        foreach ($scenes as $scene) {
            $this->assertEquals('Another Scene', $scene['meta:name']);
            $this->assertEquals('2025-10-23 12:31:45', $scene['meta:updatedDate']);
        }
        $this->assertEquals('John', $scenes[0]['@pov'][0]);
        $this->assertEquals('John', $scenes[1]['@focus'][0]);
        $this->assertNull($scenes[2]['@pov']);
    }

    public function testLoadSceneFormat2()
    {
        $lines = explode("\n", file_get_contents(__DIR__ . '/../testData/bc0cbd2a407f3.md'));
        $loader = new NovelWriterFileLoader();
        $scenes = [];
        $loader->loadScene(['_status' => 'foo'], $lines, $scenes);
        $this->assertCount(3, $scenes);
        /** @var NovelData $scene */
        foreach ($scenes as $scene) {
            $this->assertEquals('Another Scene', $scene['meta:name']);
            $this->assertEquals('2024-03-11 22:56:28', $scene['meta:updatedDate']);
        }
        $this->assertEquals('John', $scenes[0]['@pov'][0]);
        $this->assertEquals('John', $scenes[1]['@focus'][0]);
        $this->assertNull($scenes[2]['@pov']);
    }

}
