<?php
declare(strict_types=1);

/**
 * Compile to phar
 * @note php.ini setting phar.readonly must be set to false
 * lifted from https://github.com/8ctopus/webp8/blob/master/src/Compiler.php
 * parts taken from composer compiler https://github.com/composer/composer/blob/master/src/Composer/Compiler.php
 */

use Symfony\Component\Finder\Finder;

require(__DIR__ .'/vendor/autoload.php');

function buildPhar($filename, $stageTo) {
    // create phar
    $phar = new Phar($filename);

    $phar->setSignatureAlgorithm(\Phar::SHA512);

    // start buffering, mandatory to modify stub
    $phar->startBuffering();

    // add src files
    $finder = new Finder();

    $finder->files()
        ->ignoreVCS(true)
        ->name('*.php')
        ->in($stageTo . '/src');
    echo "adding src/";
    foreach ($finder as $file) {
        echo '.';
        $phar->addFile($file->getRealPath(), getRelativeFilePath($file, $stageTo));
    }
    echo "\n";

    // add vendor files
    $finder = new Finder();

    $finder->files()
        ->ignoreVCS(true)
        ->name('*.php')
        ->exclude('Tests')
        ->exclude('tests')
        ->exclude('docs')
        ->in($stageTo .'/vendor/');

    echo "Adding /vendor";
    foreach ($finder as $file) {
        echo '.';
        $phar->addFile($file->getRealPath(), getRelativeFilePath($file, $stageTo));
    }
    echo "\n";

    // entry point
    $file = './src/index.php';

    // create default "boot" loader
    $boot_loader = $phar->createDefaultStub($file);

    // add shebang to bootloader
    $stub = "#!/usr/bin/env php\n";

    $boot_loader = $stub . $boot_loader;

    // set bootloader
    $phar->setStub($boot_loader);

    $phar->stopBuffering();

    // compress to gzip
    //$phar->compress(Phar::GZ);

    echo "Create phar - OK\n";
}

/**
 * Get file relative path
 * @param  \SplFileInfo $file
 * @return string
 */
function getRelativeFilePath(SplFileInfo $file, $baseDir): string
{
    $realPath = $file->getRealPath();
    $pathPrefix = realpath($baseDir) . DIRECTORY_SEPARATOR;

    $pos = strpos($realPath, $pathPrefix);
    $relativePath = $pos !== false
        ? substr_replace($realPath, '', $pos, strlen($pathPrefix))
        : $realPath;

    return strtr($relativePath, '\\', '/');
}

/**
 * Remove the staging folder
 */
function stageDown($toDir)
{
    $fs = new Symfony\Component\Filesystem\Filesystem();
    $fs->remove($toDir);
}

/**
 * Copy required files into a staging directory and run composer in non-dev mode.
 */
function stageUp($toDir) {
    $fs = new Symfony\Component\Filesystem\Filesystem();
    $fs->mkdir($toDir);
    $fs->copy(
        __DIR__ . '/composer.json',
        $toDir . '/composer.json',
        true
    );
    $fs->mirror(__DIR__ . '/src', $toDir . '/src');
    updateVersion($toDir);
    chdir($toDir);
    exec('composer install --no-dev');
    chdir(__DIR__);
}

function updateVersion(string $toDir)
{
    $composer = json_decode(file_get_contents("$toDir/composer.json"), true);
    $version = $composer['version'];
    $entry = file_get_contents("$toDir/src/index.php");
    $entry = str_replace('{{{v}}}', $version, $entry);
    file_put_contents("$toDir/src/index.php", $entry);
}

// Build a version with no require-dev packages
$stageTo = __DIR__ . '/local/stage';
stageUp($stageTo);

$filename = './bin/novelWriterExtract.phar';
// Remove a preexisting phar.
if (file_exists($filename)) {
    unlink($filename);
}
buildPhar($filename, $stageTo);

// Clean up
stageDown($stageTo);

