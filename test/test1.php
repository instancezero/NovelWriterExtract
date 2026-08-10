<?php

use Lib\NovelWriter\ExtractGrid;

require_once __DIR__ . '/../vendor/autoload.php';

$extract = new ExtractGrid();
$extract->setSourcePath(__DIR__ . '/testData/test1/ExtractTest');
$extract->export(__DIR__ . '/results/test1.ods');
$extract->export(__DIR__ . '/results/test1.csv');
$extract->export(__DIR__ . '/results/test1.html');
$extract->export(__DIR__ . '/results/test1.xlsx');
