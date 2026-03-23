# A Metadata Extraction Tool for novelWriter

This program extracts metadata from a novelWriter project https://novelwriter.io, including data 
that follows the syntax outlined in https://github.com/vkbo/novelWriter/discussions/1769 into
a data file suitable for additional analysis.

→ Version 1.1 adds several new features.
Check the release notes at the bottom of this file for the details.

It also extracts data stored in comments, and tag references like @char and @location.

The data can be exported as a comma-separated (CSV), OpenDocument Spreadsheet (ODS),
Hypertext (HTML), or Excel (XLXS) format.
The format is determined by the extension of the output file.

**Note**: if the output file format is CSV, 
only scenes will be written since CSV files don't support multiple sheets.

Stand-alone binaries with no dependencies for Linux, Mac,
and Windows can be found in the bin/ folder.

Basic usage is `novelWriterExtract nw_project_folder output_file [format_file]`

Starting with version 1.2, `novelWriterExtract *` will prompt you for the additional arguments.

Starting with v1.1 The output file supports two formatting commands:
@z timezone_identifier@ and @d [php-date-format]@

The date specification in its simplest form of @d@ will become the current date
in the format yyyy-mm-dd. The optional format can be any valid PHP date/time string.

The timezone identifier (@z) is anything recognized by PHP, for example America/Toronto. 
If no timezone is specified then UTC is used.
The timezone specification must precede the use of @d@ or it will have no effect. 

If you're not running a binary, the code was written for PHP 8.4
but will probably run just fine in lesser versions.

Like novelWriter, the extraction tool supports multiple scenes per document.

All novelWriter files are only read, never written to.

Starting with version 2.7, novelWriter has implemented the `%story.term` constructs
and supports data export from within the application directly,
however NovelWriterExtract offers several filtering and format options not present in novelWriter.

Future changes to novelWriter might break this tool.
Please open an issue if that happens.
No warranties are explicit or implied, yada yada yada.

If you find NovelWriterExtract to be useful, please tell your fellow authors, editors, friends,
co-workers, grocery store cashiers, and random people walking down the street.
After all, if you're an author, they probably already think you're crazy.

**If you want to support my work you can 
[buy me a coffee](https://buymeacoffee.com/alanlangford).
Every little bit helps and is greatly appreciated!**

## Formats

Version 1.0.0 introduces the capability to specify which terms should be extracted, 
along with some other formatting options.
If no format file is specified, all terms are extracted from the project.
The format is defined in JSON (there's a highly specific sample in the `formats` folder).

The overall format syntax is:

```json
{
  "characters": true, // Column array or boolean
  "locations": true,  // Column array or boolean
  "scenes": [
    "column1","column2","..."
  ],
  "wordCounts": true, // Boolean
  "wrap": 40          // Integer
}
```

A column definition can be either the name of a %story term
or the name of an @ reference in NovelWriter.

### Characters

Default character columns are:

* _sequence: a sequential character number.
* name: The name of the note that contains the character information.
* tag: text from the @tag directive.
* _folder: the name of the folder the character is located in.
* given, surname, pronouns, age, hair, eyes, skin, build, fate: text from the related %story directives.
* synopsis: text from the character's %synopsis or %short directive.

Custom %story columns can also be defined.

### Locations

Default location columns are:

* _sequence: a sequential character number.
* name: The name of the note that contains the character information.
* tag: text from the @tag directive.
* _folder: the name of the folder this location is contained in.
* synopsis: text from the character's %synopsis or %story directive.

Custom %story columns can also be displayed.

### Scenes

Besides the @ tags and %story terms, these column names are available:

* _active: The value of the active column in the document tree (yes/no).
* _blank: an empty column.
* _novel: the name from the novel this scene is in.
* _sequence: a sequential scene number in the novel.
* _status: The text value associated with the status icon in the document tree.
* words: The number of words in the scene.

The "wrap" setting is not used for the CSV output file format.
It specifies the maximum width of a column in characters. The default is 40.

A simple format file could look like this:

```json
{
  "scenes": [
    "_sequence", "name", "@location", "@char", "synopsis", "words"
  ]
}
```
But columns can do much more. You can change the column heading from the default:

```json
{
  "scenes": [
    {
      "key": "@custom",
      "heading": "Additional References"
    }
  ]
}
```

For formats other than CSV, you can change the column alignment, number format,
and highlight cells that contain the first mention of a value with the "onFirst" attribute:
```json
{
  "scenes": [
    {
      "key": "@char",
      "style": {
        "align": "center",
        "onFirst": true
      }
    },
    {
      "key": "words",
      "style": {
        "numberFormat": "#,##0."
      }
    }
  ]
}
```
Unfortunately, HTML and CSV output formats aren't language sensitive,
so it's not possible to get the European number style like 1.000,00.
However, the extracts to ODS and XSLX should convert automatically.

You can break any attribute with a specific value into a new column, 
and highlight the first time the value appears, as in this example with locations:
```json
{
  "scenes": [
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
  "scenes": [
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

### Timelines

The timelines section will list the scenes that a named character appears in. By default,
the scene synopsis is listed, but this can be overriden by supplying a %story.of_{character_name}
line within the scene in novelWriter.

The timeline specification is just a list of character names.
```json
{
  "timelines": [
    "Jane", "Mark"
  ]
}
```

### Word Counts

the ```wordCounts``` flag produces a sheet with statistics on the novel's scenes.
The sheet columns tally word counts and lists the number of scenes,
broken down by active, inactive, and total.
The rows list this data by scene status with totals at the bottom.

If you don't want to see the word/scene count statistics,
they can be disabled in the JSON format specification.
```json
{
  "wordCounts": false
}

```

## Release Notes

### 1.3.1 xxxx-xx-xx

- Fixed a bug where custom story attributes weren't being reported.
- Improved default headers for custom story attributes.
For example, %story:my_thing will use "My Thing" as the header instead of "My_thing".
- Column headers are now frozen so they don't scroll off the sheet.

### 1.3.0 2025-11-28

- Added capability to extract characters and locations
- Improved column width estimation
- Major code re-work under the hood.
- "columns" element renamed to "scenes". "columns" is still recognized for backwards compatibility.

### 1.2.2 2025-11-02

- Repeated references were also being separated by double line feeds. 
This update uses a single line feed for references, two for synopsis/story comments.

### 1.2.1 2025-10-23

- novelWriter 2.8 will support repeated named comments (e.g. synopsis and the story tags),
separating multiple occurrences with two line feeds. This update supports that behaviour
  (previously a repeated comment would overwrite any earlier ones.)

### 1.2.0 2025-09-20

- Improved the output when there are insufficient arguments on the command line.
- If * is supplied as the first argument, the program will prompt for arguments. 
- Fixed a bug that was generating messy warnings.

### 1.1.0 2025-09-09

Added:

- Ability to embed date/time in output filename using @d@; set timezone with @z@ in output path.
- Improved word counts. Counts now exclude those in novelWriter commands.
- Better handling of files with multiple scenes, with individual word counts per scene.
- A new _active column lists the scene active state.
- A new _status column shows the scene status (using the text labels, not icons).
- A new statistics table counts scene and word totals, broken out by status and active state.
- It is now possible to set number formatting by column with the numberFormat style setting.

### 1.0.0 2025-04-09

Initial release with format feature to customize extracted data.
