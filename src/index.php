<?php

use Lib\NovelWriter\ExtractGrid;

require_once __DIR__ . '/../vendor/autoload.php';

$args = $argv;
array_shift($args);

$extract = new ExtractGrid();
if (count($args) === 1 && $args[0] === '*') {
    $project = readline("Path to your novelWriter project: ");
    if ($project === false) {
        return;
    }
    $extract->setSourcePath(realpath($project));
    if (!$extract->checkProject()) {
        echo "\n";
        exit;
    }
    $output = readline("Output file: ");
    if ($output === false) {
        return;
    }
    $valid = $extract->checkOutputPath($output);
    if (!$valid[0]) {
        echo "Invalid file format $valid[1]\n";
        exit;
    }
    echo "Output file is $valid[1]\n(if present, timestamp may change)\n";
    $good = strtolower(readline("Is this correct? [Y/n]: "));
    if ($good !== '' && $good !== 'y' && $good !== 'yes') {
        echo "\n";
        exit;
    }
    $format = readline("Format file [return for none]: ");
} elseif (count($args) < 2) {
    echo "novelWriterExtract version {{{v}}}\n\n"
        . "Usage: novelWriterExtract project_folder output_file [format_file.json]\n"
        . "e.g.: novelWriterExtract ~/nwsample nwsample_meta@d@.ods myformat.json\n"
        . "To be prompted for arguments use novelWriterExtract *\n\n"
        . "Supported output file types: CSV, HTML, ODS, XLSX.\n"
        . "Full details in the README.md file at https://github.com/instancezero/NovelWriterExtract\n\n"
    ;
    exit;
} else {
    $extract->setSourcePath($args[0]);
    $output = $args[1];
    $format = $args[2] ?? '';
}
try {
    $extract->export($output, $format);
} catch (Exception $exception) {
    echo "Extract failed: " . $exception->getMessage();
}
