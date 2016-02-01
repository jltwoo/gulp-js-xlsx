# gulp-js-xlsx

[![Build Status](https://travis-ci.org/DataGarage/gulp-xlsx2json.png?branch=master)](https://travis-ci.org/DataGarage/gulp-xlsx2json)

Gulp plugin that converts xlxs to JSON using [js-xlsx](https://github.com/SheetJS/js-xlsx).

This was originally gulp plugin that can converted my translation tables stored in the 
`./test/sample_dictionary.xlsx` format to json used by [angular-translate](https://github.com/angular-translate/angular-translate)

## Install

Install with [npm](https://npmjs.org/package/gulp-js-xlsx)

```
npm install --save-dev gulp-js-xlsx
```


## Example
Reads in `src/**/*.xlsx` files and outputs the json files to `./dist`
```js
var gulp = require('gulp');
var gulpXlsx = require('gulp-js-xlsx');
var rename = require('gulp-rename');

gulp.task('default', function () {
	gulp.src('src/**/*.xlsx')
		.pipe(gulpXlsx.run({
			parseWorksheet: 'tree'
		}))
		.pipe(rename({extname: '.json'}))
		.pipe(gulp.dest('dist'));
});
```


## Options
##### parseWorksheet
(default: `"row_array"`) `String|function()`  
Specifies the parsing method for each workSheet  

`"row_array"`  
			
| Name	|	Gender|	City |	Country |
| ------|------ |------| -----|
| John	|	Male  |	New  York |	United States |
| Henry	|	Male  |	Toronto |	Canada |
| Katie	|	Female|	Beijing |	China |
| Ken	  |	Male  |	Paris |	France |
```
 {
  "Sheet1": [
    {
      "Name": "John",
      "Gender": "Male",
      "City": "New York",
      "Country": "United States"
    },
    {
      "Name": "Henry",
      "Gender": "Male",
      "City": "Toronto",
      "Country": "Canada"
    }, ...
  ]
}
```

`"tree"`
| NodeKey	|	NodeValue |	ParentKey |
| ------|------ |------|
|1 |	1	|
|2 |	1.1	|1
|3 |	2	|
|4 |	2.1	|3
|5 |	2.1.1	|4
|6 |	1.2	1|
|7 |	1.2.1	|6
|8 |	2.2	3|
|9 |	1.2.2|	6
|10 |	2.3	3|

```
{
	"Sheet1": {
		"value": "root",
		"children": [
			{
				"value": "1",
				"children": [
					{ "value": "1.1", "children": [] },
					{ "value": "1.2",
						"children": [
							{ "value": "1.2.1", "children": [] },
							{ "value": "1.2.2", "children": [] }
						]
					}
				]
			},
			{
				"value": "2", 
				"children": [
					{ "value": "2.1",  children": 
						[
							{ "value": "2.1.1", "children": [] }
						]
					},
					{ "value": "2.2", "children": [] },
					{ "value": "2.3", "children": [] }
				]
			}
		]
	}
}
```

`"dictionary"`  
|KEY |	EN_US  | EN_UK |
| ------|------ |------|
|FAVOURITE |	favorite |	favourite |
|COLOR |	color |	colour |
|ALUMINIUM |	aluminum |	aluminium |
|AESTHETIC |	esthetic |	aesthetic |
|SOCIALIZE |	socialise |	socialize |
|MISSING_US |	|	britain  |
|MISSING_UK |	american |	 |  
```
{
	"Sheet1": {
		"EN_US": {
			"FAVOURITE": "favorite",
			"COLOR": "color",
			"ALUMINIUM": "aluminum",
			"AESTHETIC": "esthetic",
			"SOCIALIZE": "socialise",
			"MISSING_UK": "american"
		},
		"EN_UK": {
			"FAVOURITE": "favourite",
			"COLOR": "colour",
			"ALUMINIUM": "aluminium",
			"AESTHETIC": "aesthetic",
			"SOCIALIZE": "socialize",
			"MISSING_US": "britain"
		}
	}
}
```

`function(worksheet)`  
Every xlsx can be different. So if necessary, perform your own parsing of the worksheet object supplied by [js-xlsx](https://github.com/SheetJS/js-xlsx).

```js
// From js-xlsx module
var xlsx = require('xlsx');

gulp.task('default', function () {
	gulp.src('src/**/*.xlsx')
		.pipe(gulpXlsx.run({
			parseWorksheet: function(worksheet){
				var array = xlsx.utils.sheet_to_row_object_array(worksheet);

				// Perform your own custom parsing here...  
				// if nothing is done here, it's same as 'row_array'

				return 	array;
			}
		}))
		.pipe(rename({extname: '.json'}))
		.pipe(gulp.dest('dist'));
});
```



## Related

Similar module: [gulp-xlsx2json](https://github.com/tcarlsen/gulp-xlsx2json) I need a version that allowed me to customize how to parse the xlsx tables  
Node version: [node-xlsx](https://github.com/mgcrea/node-xlsx)

## License

MIT [@jltwoo](http://github.com/jltwoo)