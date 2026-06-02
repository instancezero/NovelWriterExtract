# A Metadata Extraction Tool for novelWriter

NovelWriterExtract is a cross-platform command-line tool using a JSON configuration file that 
extracts and processed metadata from a novelWriter project https://novelwriter.io,
including the story metadata
(as outlined in https://github.com/vkbo/novelWriter/discussions/1769.)
The extracted data can be exported to a series of tables in OpenDocument Spreadsheet,
CSV, HTML, or Microsoft Excel formats.

Check the release notes at the bottom of this file for information on updates.

It also extracts data stored in comments, and tag references like @char and @location.

The output format is determined by the extension of the filename specifed on the command line..

**Note**: if the output file format is CSV, 
only scenes will be written since CSV files don't support multiple sheets.

Stand-alone binaries with no dependencies for Linux, Mac,
and Windows can be found in the bin/ folder.

## Main Features
* Extracts scene metadata with the ability to specify which columns appear in the output.
Since story metadata can have arbitrary identifiers, this allows you to extract different views
of your novel.
* Generates scene-by-scene timelines for selected characters.
* Pulls data from the locations and character sections of the manuscript, including any story
metadata associated with those nodes, into separate character and location reference sheets.
* Supports a flexible method of calculating relative timelines, even if your novel doesn't use
Earth units for describing time.
If your story has multiple timelines, this facilitates doing a chronological sort on scenes
and character timelines.
* Can provide statistics on scene and word counts, broken down by scene status.
* Has a built-in word and phrase frequency analysis tool, 
designed to help you spot places where a word or phrase is overused.

## Usage

Basic usage is `novelWriterExtract nw_project_folder output_file [format_file]`

Starting with version 1.2, supplying an asterisk on the command line (`novelWriterExtract *`)
will cause the application to prompt you for the additional arguments.

Starting with v1.1 The output file supports two formatting commands:
@z timezone_identifier@ and @d [php-date-format]@

The date specification in its simplest form of @d@ will become the current date
in the format yyyy-mm-dd. The optional format can be any valid PHP date/time string.

The timezone identifier (@z) is anything recognized by PHP, for example America/Toronto. 
If no timezone is specified then UTC is used.
The timezone specification must precede the use of @d@ or it will have no effect. 

If you're not running a binary, the code was written for PHP 8.4
but will probably run just fine in lesser versions.

## Operation

Like novelWriter, the extraction tool supports multiple scenes per document.

All novelWriter files are only read, never written to.

Starting with version 2.7, novelWriter has implemented constructs of the form `%story.term`,
supporting data export from within the application directly,
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

Version 1.0.0 introduces the option to specify which terms should be extracted, 
along with some other formatting options.
If no format file is specified, all terms are extracted from the project.
The format is defined in JSON (there's a highly specific sample in the `formats` folder).

The overall syntax is (each section is detailed below):

```json lines
{
  "characters": true, // Array of columns or boolean
  "locations": true,  // Array of columns or boolean
  "scenes": [
    // Column specifications can be just the column name or a more complex expression.
    "column1","column2","..." 
  ],
  "time": {},         // Time unit specification. Details below.
  "timelines": {},    // Setting related to character timeline generation.
  "wordCounts": true, // Boolean
  "wrap": 40          // Integer, the number of characters to wrap multi-line columns at.
}
```

A column definition can be either the name of a %story term
or the name of an @ reference in NovelWriter.

### Characters

If the `characters` attribute is true (which is the default), 
NovelWriterExtract will generate a sheet that lists all the characters in the novelWriter project.
If it is set to false, no sheet will be produced.

The default columns in the character sheet are:

* _sequence: a sequential character number.
* name: The name of the note that contains the character information.
* @tag: text from the @tag directive.
* _folder: the name of the sub-folder the character is located in.
* (Any character attributes from the related %story directives, sorted alphabetically by name.)
* synopsis: text from the character's %synopsis or %short directive.

If an array is specified, it is a list of story attributes, 
which are included if they are used in the manuscript.  
To illustrate, if you have defined attributes for a character's nickname and age with 
constructs like this:

```
%story.age: 30
%story.nickname: The Claw
%story.build: thin
```
Then the default output columns, with `characters` set to true will be:
* _sequence
* name
* tag
* folder (only if there are character sub-folders)
* age
* build
* nickname
* synopsis

If `characters` is \[nickname, age], then the columns will be:
* _sequence
* name
* tag
* folder (only if there are character sub-folders)
* nickname
* age
* synopsis

Where the build column is omitted and the column order has changed.


### Locations

Default location columns are:

* _sequence: a sequential character number.
* name: The name of the note that contains the character information.
* tag: text from the @tag directive.
* _folder: the name of the folder this location is contained in.
* synopsis: text from the character's %synopsis or %story directive.

Custom %story columns can be displayed in the same way as in the character section.

### Scenes

Besides the @ tags and %story terms, these column names are available:

* _active: The value of the active column in the document tree (yes/no).
* _blank: an empty column.
* _chron: A relative time (details below).
* _novel: the name from the novel this scene is in.
* _sequence: a sequential scene number in the novel.
* _sla: A sentence length analysis (details below).
* _slg: A sentence length graph. (details below).
* _status: The text value associated with the status icon in the document tree.
* words: The number of words in the scene.

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

#### Relative time

Sometimes when a story has multiple timelines, it's useful to be able to look at the
story structure on an ascending timeline. The application provides two ways of doing this
through the `%story.time` data.

Fixed time mode will attempt to parse a human-readable date/time string 
and convert it to a sortable ISO8601 value. For example, "%story.time: March 5, 2001 9:15pm"
should result in a _chron value of 2001-03-05T22:15"

Relative time is expressed in time units. These default to Earth units but can be customized.

The time mode is specified in the "time" section of the format file:

```json lines
{
  "time": {
    "mode": "fixed|relative|off" // Any value other than fixed or relative will be interpreted as off.
  }
}

```

Preset relative Earth units are:
* No units or 'm': minutes.
* 'h': hours of 60 minutes.
* 'd': days of 24 hours.
* 'w': weeks of 7 days.
* 'mo': months of 30 days.
* 'y': years of 12 months.

"%story.time: 15" would represent 15 minutes into the start of the story.
"%story.time: 3mo" represents three months in (a default month is fixed at 30 days.)

It is possible to set a base time and then use it in time expressions. 
In one scene you can define a base time: "%story.time: prolog=-4.5y" 
Sets "prolog" to 4.5 years before "time zero."
You can then use "prolog" as the basis for other times,
so "%story.time: prolog+6mo" is six months after five years in the past, or -4 years.
The only rule is that the base time must be defined before it is used.

If your story uses its own time system, you can accommodate this with custom units. 
Times with no unit specification will be taken as unit time.
Everything else is a multiple of that or another defined unit. 
Units are specified as part of the time configuration: 

```json lines
{
  "time": {
    "mode": "relative",
    "units": {
      "zip": 1,
      "blarg": "16zip",
      "snarf": "128blarg"
    }
  }
}

```
In this time system, the base unit is a zip. A "blarg" is 16 zips, and a "snarf" is 128 "blargs",
or 2048 zips.


#### Sentence Length Analysis

The `_sla` column produces a compressed representation of sentence lengths in the scene. 
The first element is the number of sentences in the scene and the average number of sentences
per paragraph, for example "102@3.5:" means there's 102 sentences in the scene 
and the average paragraph is 3.5 sentences long.

Following the first element, there is one comma-separated string per paragraph.
The string starts with a P and the number of sentences in the paragraph and a colon.
The rest of the string characterizes the sentences in the paragraph by length.
Each sentence is assigned an s if it contains less than five words,
an m if it contains five to nineteen, and an l if it has 20 or more words.
Groups of sentences with the same length are assigned a multiplier.

For example, the string "P9:2l.6m.s" means the paragraph has nine sentences, two long,
followed by six medium and one short. Expanded, this would be "P9:llmmmmmms".

While complex, this is designed to make it easier to detect sequences 
of paragraphs with the same length, like the six medium-length sentences in the example.

With this release, the criteris for sentence length is pre-set and fixed.
I'll look at ways to change that in future releases.

#### Sentence Length Graph

The `_slg` column is intended to provide the same kind of information 
as the sentence length analysis, but in a more visual way. 
Each sentence is represented by a vertical stack with eight possible levels:

* One bar: 1–2 words.
* Two bars: 3–5 words.
* Three bars: 6–8 words.
* Four bars: 9–11 words.
* Five bars: 12–14 words.
* Six bars: 15–17 words.
* Seven bars: 18–20 words.
* Eight bars: 21+ words.

A space separates each paragraph. This results in output like this:

▆█▄▄▇▆▄▆▄▆▂ ▃▄▇▆▂


### Timelines

The timelines section lists the scenes that a named character appears in. By default,
the scene synopsis is listed, but this can be overridden by supplying a %story.of_{character_tag}
line within the scene in novelWriter. The "of_" construct allows the author to relate the
scene ffrom the perspective of the named character.

The timeline specification can limit characters by the number of scenes they appear in:
```json lines
{
  "timelines": {
    "minimum": 4    // Characters appearing in less than four scenes will not be generated
  }
}
```
The default minimum is zero, which will generate a sheet for every character.

You can also specify which characters to generate (with or without the minimum):
```json lines
{
  "timelines": {
    "chars": ["Bob", "Shivanna"],
    "minimum": 4
  }
}
```
This will only produce sheets for the two named characters if they appear in four or more scenes.

By default, character sheets include the %story.time and (if enabled) the relative time columns.
You can change this with the "show" option:
```json lines
{
  "timelines": {
    "chars": ["Bob", "Shivanna"],
    "minimum": 4,
    "show": ["time", "_chron"]    // Only output the named columns. Use an empty array for none.
  }
}
```

It is also possible to redirect timelines to a Markdown document 
(which recent versions of LibreOffice can read) by setting the `markdown` flag true.
In this case, the data will not be added to the output sheets, 
but instead directed to a file with the same base name as the sheets and a ".md" extension.
```json lines
{
  "timelines": {
    "markdown": true
  }
}
```


### Word and Phrase Use and Clustering Analysis

The `analysis` flag is a boolean true or false (default). When enabled, 
the program will generate two tables on an Analysis sheet.

Note that the analysis process is compute-intensive and will take some time to process.

The tables present frequency, "Clumpiness", and "Average Clumpiness". 
The first reports on individual words; the second reports on phrases of two or three words.

"Clumpiness" is a metric that is higher when occurrences of the word/phrase 
are closer to each other in a scene. The higher this number is, 
the more likely that the word or phrase is repeated more than once in close proximity.

### Word Counts

The ```wordCounts``` flag produces a sheet with statistics on the novel's scenes.
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

### Wrap

The "wrap" setting specifies the maximum width of a column in characters. The default is 40.
This does not apply to the CSV output file format

## Release Notes

### 1.4.1 2026-05-11

Fixed:
- A renaming issue caused the _chron column to come up blank.

### 1.4.0 2026-05-08

Fixed:
- A bug where custom story attributes weren't being reported.

Added:
- Relative time calculations
- Character timelines
- Analysis tools

Changed:
- Improved default headers for custom story attributes.
For example, %story:my_thing will use "My Thing" as the header instead of "My_thing".
- Column headers are now frozen so they don't scroll off the sheet.

### 1.3.0 2025-11-28

- Added the capability to extract characters and locations
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
