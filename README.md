# A novelWriter Metadata Extraction Tool

This program extracts metadata from a novelWriter project https://novelwriter.io, including data 
that follows the syntax outlined in https://github.com/vkbo/novelWriter/discussions/1769 into
a data file suitable for additional analysis.
It also extracts data stored in comments, and tag references like @char and @location.

The data can be exported as a comma-separated (CSV), OpenDocument Spreadsheet (ODS),
Hypertext (HTML), or Excel (XLXS) format.
The format is determined by the extension of the output file.

Binaries for linux, Mac, and Windows can be found in the bin/ folder.

Basic usage is `novelWriterExtract nw_project_folder output_file [format_file]`

If you're not running a binary, the code was written for PHP 8.3
but will probably run just fine in lesser versions.

All novelWriter files are only read, not written.
I expect that novelWriter will implement the `%story.term` constructs in the near future
and may support data export from within the application directly,
thus making this tool somewhat redundant.
Also, the final implementation might vary depending on how `vkbo`
(novelWriter's author and maintainer) approaches it, so this code might just break entirely.
No warranties are explicit or implied, yada yada yada.

Like novelWriter, the extraction tool supports multiple scenes per document.
However, the "Word Count" column applies to the entire file,
so only the first scene in the file will show a count.

## Formats

Version 1.0.0 introduces the capability to specify which terms should be extracted, 
along with some other formatting options.
If no format file is specified, all terms are extracted from the project.
The format is defined in JSON (there's a highly specific sample in the `formats` folder).

The overall format syntax is:

```json
{
  "columns": [
    "column1","column2","..."
  ],
  "wrap": integer
}
```

A column definition can be either the name of a %story term or the name of an @ reference
in NovelWriter.
Additional column names are:

* _blank: an empty column.
* _novel: the name from the novel this scene is in.
* _sequence: a sequential scene number in the novel.
* words: The number of words in the scene.

The "wrap" setting is not used for the CSV output file format.
It specifies the maximum width of a column in characters. The default is 40.

A simple format file could look like this:

```json
{
  "columns": [
    "_sequence", "name", "@location", "@char", "synopsis", "words"
  ]
}
```
But columns can do much more. You can change the column heading from the default:

```json
{
  "columns": [
    {
      "key": "@custom",
      "heading": "Additional References"
    }
  ]
}
```

For formats other than CSV, you can change the column alignment, 
and highlight cells that contain the first mention of a value with the "onFirst" attribute:
```json
{
  "columns": [
    {
      "key": "@char",
      "style": {
        "align": "center",
        "onFirst": true
      }
    }
  ]
}
```
You can break any attribute with a specific value into a new column, 
and highlight the first time the value appears, as in this example with locations:
```json
{
  "columns": [
    {
      "heading": "Europe",
      "key": "@location",
      "test": [
        {
          "arg": "@location",
          "op": "has",
          "value": ["France", "Germany", "Italy"]
        }
      ]
    },
    {
      "heading": "N. America",
      "key": "@location",
      "test": [
        {
          "arg": "@location",
          "op": "has",
          "value": ["Canada", "USA", "Mexico"]
        }
      ]
    }  
  ]
}
```

You can break your main characters into separate columns
and create a column for secondary characters:
```json
{
  "columns": [
    {
      "heading": "Hero",
      "test": [
        {
          "arg": "@char",
          "op": "includes",
          "value": "Suzie"
        }
      ],
      "result": "Sue"
    },
    {
      "heading": "Ally",
      "test": [
        {
          "arg": "@char",
          "op": "includes",
          "value": "Mark"
        }
      ],
      "result": "Mark"
    },
    {
      "heading": "Villain",
      "test": [
        {
          "arg": "@char",
          "op": "includes",
          "value": "Darth"
        }
      ],
      "result": "Darth"
    },
    {
      "key": "@char",
      "heading": "Others",
      "exclude": [
        "Sue",
        "Mark",
        "Darth"
      ]
    }
  ]
}
```
