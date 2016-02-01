'use strict';
var fs = require('fs');
var gulp = require('gulp');
var assert = require('assert');
var map = require('map-stream');
var rename = require('gulp-rename');
var gulpXlsx = require('../lib');

function expectStream(done){
	return map(function(file,cb){
		var json = JSON.parse(String(file.contents));
		cb();
		done();
	})
}


it('should convert xlsx to json', function (cb) {
	gulp.src(__dirname + '/sample.xlsx')
	.pipe(gulpXlsx.run())
  .pipe(expectStream(cb));
});

it('should convert xlsx tree format to json', function (cb) {
	gulp.src(__dirname + '/sample_tree.xlsx')
	.pipe(gulpXlsx.run({
		parseWorkSheet: 'tree'
	}))
  .pipe(expectStream(cb));
});

it('should convert xlsx key-maps to json', function (cb) {
	gulp.src(__dirname + '/sample_dictionary.xlsx')
	.pipe(gulpXlsx.run({
		parseWorkSheet: 'dictionary'
	}))
  .pipe(expectStream(cb));
});
