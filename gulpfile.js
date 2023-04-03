'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const del = require("del");
var path = require('path');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

gulp.task('full-clean', function (cb) {
  del(["lib", "dist", "package", "release", "sharepoint", "temp"]).then(() => cb(), err => cb(err));
});

gulp.task('update-version', function (cb) {

  const gutil = require('gulp-util');
  const fs = require('fs');

  var src = require('./package.json');
  var version = src.version.split('-')[0];

  gutil.log('Setting version to:\t' + version);

  var packageSolutionJson = require('./config/package-solution.json');
  packageSolutionJson.solution.version = version + '.0';
  packageSolutionJson.solution.includeClientSideAssets = false;
  fs.writeFileSync('./config/package-solution.json', JSON.stringify(packageSolutionJson, null, 2));

  var webPartManfiestJson = require('./src/WebPart/WebPart.manifest.json');
  webPartManfiestJson.version = version;
  fs.writeFileSync('./src/WebPart/WebPart.manifest.json', JSON.stringify(webPartManfiestJson, null, 2));

  var writeManfiestJson = require('./config/write-manifests.json');
  writeManfiestJson.cdnBasePath = `https://cdn.jsdelivr.net/gh/nbelyh/diagram-frame-releases/${version}`;
  fs.writeFileSync('./config/write-manifests.json', JSON.stringify(writeManfiestJson, null, 2));

  cb();
});

build.tslintCmd.enabled = false;

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

const eslint = require('gulp-eslint');
const eslintSubTask = build.subTask('eslint-subTask', function (gulp, buildOptions, done) {
  return gulp.src(['src/**/*.{ts,tsx}'])
    .pipe(eslint('./eslint.json'))
    .pipe(eslint.format())
    .pipe(eslint.failAfterError());
});
build.rig.addPreBuildTask(build.task('eslint', eslintSubTask));

build.initialize(gulp);

build.mergeConfig({
  showToast: false
});
