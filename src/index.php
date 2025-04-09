<?php

use Lib\NovelWriter\ExtractGrid;

require_once __DIR__ . '/../vendor/autoload.php';

$args = $argv;
array_shift($args);
if (count($args) < 2) {
    echo "novelWriterExtract version {{{v}}}\n\n"
        . "Usage: php " . __FILE__ . " project_folder output_file [format_file.json]\n"
        . "eg: novelWriterExtract ~/nwsample nwsample_meta.ods\n"
        . "Supported output file types: CSV, HTML, ODS, XLSX\n";
} else {
    try {
        $extract = new ExtractGrid();
        $extract->setSourcePath($args[0]);
        $extract->export($args[1], $args[2] ?? '');
    } catch (Exception $exception) {
        echo "Extract failed: " . $exception->getMessage();
    }
}
