/**
 * INSTALL
 *
 * This requires nodejs and npm to be installed:
 *   - https://nodejs.org/en/download/package-manager/
 *
 * Install gulp:
 *   - npm install -g gulp
 *
 * Install required gulp modules:
 *   - npm install gulp-util gulp-uglify gulp-rename gulp-minify-css gulp-concat
 *
 * 
 * RUN
 *
 * Create minified JS:
 *   - gulp scripts
 *
 * Create minified CSS:
 *   - gulp styles
 */

/* Needed gulp config */
var gulp = require('gulp');  
var uglify = require('gulp-uglify');
var rename = require('gulp-rename');
var minifyCss = require('gulp-minify-css');
var concat = require('gulp-concat');

/* scripts task */
gulp.task('scripts', function() {
  return gulp.src([
    /* Add your JS files here, they will be combined in this order */
    'socialcalcconstants.js',
    'socialcalctableeditor.js',
    'formatnumber2.js',
    'formula1.js',
    'socialcalcpopup.js',
    'socialcalcspreadsheetcontrol.js',
    'socialcalcviewer.js'
    ])
    .pipe(concat('socialcalc.js'))
    .pipe(gulp.dest(''))
    .pipe(rename({suffix: '.min'}))
    .pipe(uglify())
    .pipe(gulp.dest(''))
});

/* styles task */
gulp.task('styles', function() {
  return gulp.src([
    /* Add your CSS files here, they will be combined in this order */
    'socialcalc.css',
    ])
    .pipe(rename({suffix: '.min'}))
    .pipe(minifyCss())
    .pipe(gulp.dest(''))
});

// The default task (called when you run `gulp` from cli)
gulp.task('default', ['scripts', 'styles']);
