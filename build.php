<?php
// Build the PHAR
echo "Building PHAR.\n";
exec('C:\tools\php84\php.exe compile.php');

// Wipe out previous files
echo "Removing old binaries.\n";
$base = __DIR__ . '\bin';
foreach (['linux', 'mac', 'windows'] as $os) {
    $list = scandir("$base\\$os");
    foreach ($list as $file) {
        if ($file[0] === '.') {
            continue;
        }
        unlink("$base\\$os\\$file");
    }
}

// Build the executables
echo "Building binaries.\n";
exec('.\vendor\bin\phpacker build --src=.\bin\novelWriterExtract.phar --dest=.\bin');

// Rename the executables
echo "Renaming binaries.\n";
$composer = json_decode(file_get_contents(__DIR__ . '\composer.json'), true);
$version = $composer['version'];
$base = __DIR__ . '/bin';
if (file_exists("$base/linux/linux-arm")) {
    rename("$base/linux/linux-arm", "$base/linux/novelWriterExtract-$version-arm");
    rename("$base/linux/linux-x64", "$base/linux/novelWriterExtract-$version-x64");
    rename("$base/mac/mac-arm", "$base/mac/novelWriterExtract-$version-arm");
    rename("$base/mac/mac-x64", "$base/mac/novelWriterExtract-$version-x64");
    rename("$base/windows/windows-x64.exe", "$base/windows/novelWriterExtract-$version-x64.exe");
} else {
    echo "No binaries found.\n";
}
