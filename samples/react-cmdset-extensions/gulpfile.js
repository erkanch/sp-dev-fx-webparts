'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

gulp.task('version-sync', function (done) {
  const gutil = require('gulp-util');
  const fs = require('fs');
  var pkgConfig = require('./package.json');
  var pkgSolution = require('./config/package-solution.json');

  console.log('Old Version:\t' + pkgSolution.solution.version);
  //var newVersionNumber = pkgConfig.version.split('-')[0] + '.0';
  var newVersionNumber = pkgConfig.version + '.0';
  pkgSolution.solution.version = newVersionNumber;
  pkgSolution.solution.features.forEach((f) => f.version = newVersionNumber);
  console.log('New Version:\t' + pkgSolution.solution.version);
  // write changed package-solution file
  fs.writeFile('./config/package-solution.json', JSON.stringify(pkgSolution, null, 4), err => console.log(err));
  done();
});

/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(require('gulp'));

