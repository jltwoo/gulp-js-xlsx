var gulp = require('gulp');
var xlsx2json = require('./');
var rename = require('gulp-rename');

gulp.task('default', function () {
    gulp.src('./sample/*.xlsx')
        .pipe(xlsx2json())
				.pipe(rename({extname: '.json'}))
        .pipe(gulp.dest('./dist'));
});
