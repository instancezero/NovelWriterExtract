<?php
$composer = json_decode(file_get_contents(__DIR__ . '/composer.json'), true);
$version = $composer['version'];
$base = __DIR__ . '/bin';
rename("$base/linux/linux-arm", "$base/linux/novelWriterExtract-$version-arm");
rename("$base/linux/linux-x64", "$base/linux/novelWriterExtract-$version-x64");
rename("$base/mac/mac-arm", "$base/mac/novelWriterExtract-$version-arm");
rename("$base/mac/mac-x64", "$base/mac/novelWriterExtract-$version-x64");
rename("$base/windows/windows-x64.exe", "$base/windows/novelWriterExtract-$version-x64.exe");
