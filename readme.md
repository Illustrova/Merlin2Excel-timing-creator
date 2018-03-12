# Merlin to Excel Timetable Converter

This is an in-house tool which helps to convert xml exported from Merlin project to predefined excel template. It's probably useless for anyone else then my current colleagues, due to very specific kind of data and output template, although it will (should) work properly.

## Getting Started

1. Clone the repository  on your machine and install dependencies.
2. Run the conversion: copy your xml file into the source folder and run command:
```
node index.js
```

or just run:

```
node index.js path/to/file.xml
```

3. Your excel will be placed in root folder.

## Source file

Export source file from Merlin project as xml file.

__Important!__ When exporting, make sure that option *"include calculated values"* is checked (it is unchecked by default).
2. The result will be a file with `.mprojectx` extension. On Mac, right-click it and press "Show package contents" to access the xml. On Windows .mprojectx will be displayed as folder. Get the `content.xml` and place it into source folder/or just copy the path to it.

## Options

Options are defined in `config.js` in root folder.

###### companyName ######
Company name will be used in excel filename.
##### publicHolidays #####
Array of date strings, representing public holidays/days off. They will be marked with color and text in the .xlsx file. Strings can be in various formats, but suitable for parsing by Moment.js library. See the [docs](https://momentjs.com/docs/#/parsing/string/) for more info.
##### colors.match #####
An object where the key is the color name, and value is a string or array of strings containing regular expressions to search for match. If type of activity is defined, it will be used for match; otherwise the title of activity will be tested.
##### colors.reserve #####
If no matches with test strings found, one of these colors will be used.

## Try it

The ```/examples``` folder contains few sample xml files, try to run it according the instructions above. You also may test with your personal files, but color settings will not work, unless you edit config.js according your data. An example of external data for test is available at `examples/input/sample_otherContent.xml`.

## Built With

* [xml2js](https://github.com/Leonidas-from-XIV/node-xml2js) - XML parser
* [Excel4node](https://github.com/natergj/excel4node) - Excel builder
* [Moment.js](https://momentjs.com/) - Date/time management library

## Development notes
This is an occasional small project created while I was working in the company which runs CG and VFX jobs for commercial production. In our daily routine we were using [Merlin project](https://www.projectwizards.net/en/products/merlin-project/what-is) for resources and timetable management, however for presentational purposes we used to manually move data to existing excel template.

The main challenge of this project was to transform data that way that it matches the template exactly. Besides of moving the data accordingly, it also has to be positioned nicely (each work process should keep same line, and as close to the center as possible, depending on the number of other activities). Also each type of activity should preserve it's color for visual purposes.

## Authors

* **Irina Illustrova** - [Illustrova](https://github.com/Illustrova)