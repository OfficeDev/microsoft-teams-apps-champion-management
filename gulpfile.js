'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Font loader configuration for webfonts
const fontLoaderConfig = {
    test: /\.(woff|woff2|eot|ttf)$/,
    use: [{
        loader: 'url-loader',
    }]
};
build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
        generatedConfiguration.module.rules.push(fontLoaderConfig);
        return generatedConfiguration;
    }

});
var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};
build.initialize(gulp);

build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
      generatedConfiguration.module.rules.push(
        {
            test: /\.(woff2)$/,
            loader: 'url-loader?limit=100000'
        }
      );
      return generatedConfiguration;
    }
  });

  build.sass.setConfig({ warnOnNonCSSModules: false, useCssModules:true});
