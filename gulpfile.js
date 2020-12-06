'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
var path = require('path');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

gulp.task('update-version', function () {

  const gutil = require('gulp-util');
  const fs = require('fs');

  var src = require('./package.json');
  var version = src.version.split('-')[0];

  gutil.log('Setting version to:\t' + version);

  var packageSolutionJson = require('./config/package-solution.json');
  packageSolutionJson.solution.version = version + '.0';
  fs.writeFileSync('./config/package-solution.json', JSON.stringify(packageSolutionJson, null, 2));

  var webPartManfiestJson = require('./src/WebPart/WebPart.manifest.json');
  webPartManfiestJson.version = version;
  fs.writeFileSync('./src/WebPart/WebPart.manifest.json', JSON.stringify(webPartManfiestJson, null, 2));
});

build.initialize(gulp);

build.mergeConfig({
  showToast: false
});
