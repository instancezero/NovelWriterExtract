<?php

use Lib\NovelWriter\ExtractGrid;

require_once './vendor/autoload.php';

$args = $argv;
array_shift($args);
if (count($args) < 2) {
    echo "Usage: php " . __FILE__ . " project_folder output_file\n"
        . "eg: php " . __FILE__ . "~/nwsample nwsample_meta.ods\n"
        . "Supported output file types: CSV, ODS, XLSX\n";
} else {
    try {
        $extract = new ExtractGrid();
        $extract->setSourcePath($args[0]);
        $extract->export($args[1]);
    } catch (Exception $exception) {
        echo "Extract failed: " . $exception->getMessage();
    }
}
