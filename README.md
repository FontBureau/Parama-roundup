# Parama-roundup

### Resource Summary
* [Example var repo](https://github.com/typenetwork/amstelvar/)
* [Example designspace file](https://github.com/TypeNetwork/Amstelvar/blob/master/sources/Amstelvar-NewCharset/Amstelvar-Roman-010.designspace)
* [Example input sheet](https://docs.google.com/spreadsheets/d/1L1Cy2Y1JFOl32nuevTkzcxjNsSTlLGn_oiRLkMnyonY)

### Specification

Description of the functionality of a python program

Overview - the parametric axes values of variable fonts need to gathered from ufos and instances via a design space file and written to a google spreadsheet

A source google sheet name and location will supply a design space file name and location. A column ID, containing a list of axes names, each with a glyph id and one or more glyph index points, and a units per em?

An axis with one index point value is to be measured in y from the index point to the origin. Positive and negative values should be returned.
An axis with two index point values measures the x or y distance as indicated by the first character of the axis name.

A target google sheet name accepts a list of values for each ufo and instance in the design space, the values are written in columns with the style names to match the original axes list.

What can be used from there, currently consists of QA, writing true values for parametric axes to JavaScript, and for use by web programs.

Exporting, sheet to Designspace file, will be useful and specified, if roundup is comepleted.

### Instructions

1. Run `quickstart.py` on the command line, following instructions to log in to your Google account. [Visit the APIv4 website for further instructions.](https://developers.google.com/sheets/api/quickstart/python))

```$ python quickstart.py```

You only have to do this the first time to set up your tokens.

2. Create a Google spreadsheet with a sheet called `Index Refs` modeled on the [example sheet](https://docs.google.com/spreadsheets/d/1L1Cy2Y1JFOl32nuevTkzcxjNsSTlLGn_oiRLkMnyonY)

3. Click on “Share” in the upper right, and set spreadsheet permissions to “Anyone with the link can edit”

![Sharing settings](assets/sharing.png)

4. Record the Google Sheet ID for the sheet, which is the long sequence of letters and numbers found in the URL:

<pre>https://docs.google.com/spreadsheets/d/<strong><span style="color: red">1L1Cy2Y1JFOl32nuevTkzcxjNsSTlLGn_oiRLkMnyonY</span></strong>/edit?usp=sharing</pre>

5. Run `buildSpreadsheet.py` on the command line 
<pre>$ cd /path/to/my/repository/sources/
$ python3 /path/to/parama_roundup/buildSpreadsheet.py -d <strong>MyDesignspace.designspace</strong> -i <strong>INPUT_GOOGLE_SHEET_ID</strong> -o <strong>INPUT_GOOGLE_SHEET_ID</strong></pre>

| Arg | Description                                                           |
|----|------------------------------------------------------------------------|
| -d | Designspace filename                                                   |
| -i | Spreadsheet ID for Input sheet (must contain “Index Refs” sheet        |
| -o | Spreadsheet ID for Output sheet (if omitted, input sheet will be used) |

6. Check the output spreadsheet for `Axes`, `Measurements`, and `Widths` sheets with the resulting data