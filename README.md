# Parama-roundup

### Resource Summary
* [Example var repo](https://github.com/typenetwork/amstelvar/)
* [Example designspace file](https://github.com/TypeNetwork/Amstelvar/blob/master/sources/Amstelvar-NewCharset/Amstelvar-Roman-009.designspace)
* [Example output sheet](https://docs.google.com/spreadsheets/d/1r0oqR3ic8qQJycGgmZW-kqKEzgoPIIiKxt4MhOJFu7U)

### Specification

Description of the functionality of a python program

Overview - the parametric axes values of variable fonts need to gathered from ufos and instances via a design space file and written to a google spreadsheet

A source google sheet name and location will supply a design space file name and location. A column ID, containing a list of axes names, each with a glyph id and one or more glyph index points, and a units per em?

An axis with one index point value is to be measured in y from the index point to the origin. Positive and negative values should be returned.
An axis with two index point values measures the x or y distance as indicated by the first character of the axis name.

A target google sheet name accepts a list of values for each ufo and instance in the design space, the values are written in columns with the style names to match the original axes list.

What can be used from there, currently consists of QA, writing true values for parametric axes to JavaScript, and for use by web programs.

Exporting, sheet to Designspace file, will be useful and specified, if roundup is comepleted.

