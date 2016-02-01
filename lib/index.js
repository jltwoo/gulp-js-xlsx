'user strict';

var path = require('path');
var xlsx = require('xlsx');
var extend = require('util')._extend;
var map = require('map-stream');
var gUtil = require('gulp-util');

module.exports.util = xlsx.utils;

var parsers = {};
parsers.sheetToRowObjectArray = function (worksheet){
	var json = xlsx.utils.sheet_to_row_object_array(worksheet);
	return json;
};


// This function assumes the table follows the format
// specified in sample_dictionary.xlsx so each column represents
// a new dictionary
parsers.sheetToDictionary = function(worksheet){
	var array = xlsx.utils.sheet_to_row_object_array(worksheet);
	var dict = {};
	for (var i = 0; i < array.length; i++) {
		var row = array[i];
		if(!row.KEY){
			throw 'No key specified for value';
		}
		for(var dictName in row){
			if(row.hasOwnProperty(dictName)){
				if(dictName !== "KEY"){
					dict[dictName] = dict[dictName] || {};
					dict[dictName][row.KEY] = row[dictName];
				}
			}
		}
	}
	return dict;
}



// This function assumes the table follows this format:
// NodeKey,NodeValue,ParentKey
// 1, First Node, 
// 2, Child Node, 1
// first line is treated as header values

parsers.sheetToTree = function(worksheet, opts){
	var array = xlsx.utils.sheet_to_row_object_array(worksheet);
	var nodes = {};
	var root = {
		value: 'root',
		children: []
	};
	// Ignores the first line as header
	for (var i = 0; i < array.length; i++) {
		var row = array[i];
		if(!nodes[row.NodeKey]){
			var newNode = {
				value: row.NodeValue,
				children: []
			}
			nodes[row.NodeKey] = newNode;

			var parent = root;
			if(row.ParentKey){ // has parent
				if(nodes[row.ParentKey]){
					parent = nodes[row.ParentKey];
				} else { // parent doesn't exist yet, may be further down the list
					parent = { value: null };
					nodes[row.ParentKey] = parent;
				}
			}
			parent.children.push(newNode);
		}
		else{
			gUtil.log('gulp-js-xlsx: duplicated id detected: ' + gUtil.colors.red(row[0]));
		} 
	};
	return root;
};

module.exports.parsers = parsers;
module.exports.run = function (options) {

  var opts = extend({}, options);

	// should return a json object
	function parseWorkbook(workbook){
		var sheet_name_list = workbook.SheetNames;
		var root = {};
		sheet_name_list.forEach(function(sheetName) { /* iterate through sheets */
		  var worksheet = workbook.Sheets[sheetName];
		  var fnParseWorksheet = parsers.sheetToRowObjectArray;
		  if(opts.parseWorkSheet){
  		  if(typeof opts.parseWorkSheet === 'function'){
  		  	fnParseWorksheet = opts.parseWorkSheet;
  		  } 
  		  else if(typeof opts.parseWorkSheet=== 'string'){
  		  	switch(opts.parseWorkSheet){
  		  		case "row_array":
  		  				fnParseWorksheet = parsers.sheetToRowObjectArray;
  		  			break;
  		  		case "dictionary":
  		  				fnParseWorksheet = parsers.sheetToDictionary;
  		  			break;
  		  		case "tree":
  		  				fnParseWorksheet = parsers.sheetToTree;
  		  			break;
  		  		default:
  						throw 'Unsupported parser specified: '+opts.parseWorkSheet;
  		  		break;
  		  	}
  		  } 
  		  else {
  				throw 'Unsupported value type specified for parser: '+opts.parseWorkSheet;
  		  }
  		}

		  var json = fnParseWorksheet(worksheet);
		  return root[sheetName]= json;
		});
		return JSON.stringify(root);
	};

	return map(function (file, cb) {
		if (file.isNull()) { return cb(null, file); }

		if (file.isStream()) {
			return cb(new gUtil.PluginError('gulp-js-xlsx', 'Streaming not supported'));
		}

		if (['.xls','.xlsx','.ods','.xlsm'].indexOf(path.extname(file.path)) === -1) {
			gUtil.log('gulp-js-xlsx: Skipping unsupported format ' + gUtil.colors.red(file.relative));
			return cb(null, file);
		}

		if(file.isBuffer()){
      try {
      	var fnParse = opts.parse || parseWorkbook;
				var output = fnParse(
					xlsx.readFile(file.path)
				);

        file.contents = new Buffer(output);
				file.path = gUtil.replaceExtension(file.path, '.json');

      } catch(e) {
        return cb(new gUtil.PluginError('gulp-js-xlsx', e));
      }
    }
    cb(null, file);

	});
};