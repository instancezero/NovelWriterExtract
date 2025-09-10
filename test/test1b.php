<?php

use Lib\NovelWriter\ExtractGrid;

require_once __DIR__ . '/../vendor/autoload.php';

$extract = new ExtractGrid();
$extract->setSourcePath(__DIR__ . '/projects/test1/ExtractTest');
$extract->export(__DIR__ . '/test1.ods', __DIR__ . '/test1b.json');
$extract->export(__DIR__ . '/test1.csv', __DIR__ . '/test1b.json');
$extract->export(__DIR__ . '/test1.html', __DIR__ . '/test1b.json');
$extract->export(__DIR__ . '/test1.xlsx', __DIR__ . '/test1b.json');
