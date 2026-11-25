'use strict';

const path = require('path');

module.exports = {
  entry: path.resolve(
    __dirname,
    'lib/classic/ClassicBootstrap.js'
  ),

  output: {
    path: path.resolve(__dirname, 'dist/classic'),
    filename: 'ESignClassic.js',
    library: 'ESignClassic',
    libraryTarget: 'var'
  },

  resolve: {
    extensions: ['.js']
  },

  devtool: 'source-map',

  module: {
    rules: []
  }
};
