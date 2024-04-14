# novelWriter MetaData Extraction Tool

This is a PHP script that extracts metadata from a novelWriter project https://novelwriter.io 
that follows the syntax outlined in https://github.com/vkbo/novelWriter/discussions/1769 into
a data file suitable for additional analysis. It also extracts data stored in comments, @char,
and @location constructs. Additional support is likely, but the tool was built primarily for my
own purposes, so that may take a while or not happen at all.

The data can be exported as a comma-separated (CSV), OpenDocument Spreadsheet (ODS), or
Excel (XLXS) format.

Basic usage is `php novelWriterExtract.php nw_project_folder output_file`

The code was written for PHP 8.3 but will probably run just fine in lesser versions.

This is proof-of-concept code. Use at your own risk, although at this point the novelWriter files
are only read, not written. I expect that novelWriter will implement the `%story.term`
constructs in the not too distant future and may support data export from within the application
directly, thus making this tool somewhat redundant. Also, the final implementation might vary
depending on how `vkbo` (novelWriter's author and maintainer) approaches it, so this code might
just break entirely. No warranties are explicit or implied, yada yada yada.

Like novelWriter, the extraction tool supports multiple scenes per document. However, the "Word
Count" column applies to the entire file, so only the first scene in the file will show a count.

There is an extremely small chance that I'll enhance this code to perform the reverse operation,
namely to extract updates from a spreadsheet and post to novelWriter files with the changes.
While that would be really cool, it's an approach that could be fraught with problems that test
your backups.

