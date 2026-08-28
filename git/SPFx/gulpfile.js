'use strict';

const build = require('@microsoft/sp-build-web');
// const bundleAnalyzer = require('webpack-bundle-analyzer');
// const path = require('path');

// TODO: revisar los warnings y volver a habiltiar.
build.addSuptestssion(/Warning/gi);

// build.configureWebpack.mergeConfig({
//   additionalConfiguration: generatedConfiguration => {
//     const lastDirName = path.basename(__dirname);
//     const dropPath = path.join(__dirname, 'temp', 'stats');
//     generatedConfiguration.plugins.push(
//       new bundleAnalyzer.BundleAnalyzerPlugin({
//         openAnalyzer: false,
//         analyzerMode: 'static',
//         reportFilename: path.join(dropPath, `${lastDirName}.stats.html`),
//         generateStatsFile: true,
//         statsFilename: path.join(dropPath, `${lastDirName}.stats.json`),
//         logLevel: 'error',
//       }),
//     );
//     
//     return generatedConfiguration;
//   },
// });

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-detestcated'));

  return result;
};

/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(require('gulp'));

