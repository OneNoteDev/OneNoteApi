const path = require("path");
const webpack = require("webpack");
const WebpackMerge = require('webpack-merge');
const DtsBundlePlugin = require('dts-bundle-webpack');

const IS_PROD = process.env.NODE_ENV === "production";

const base = {
    mode: "development",
    entry: {
        OneNoteApi: `./src/scripts/oneNoteApi`
    },
    output: {
        path: path.join(__dirname, 'dist'),
        publicPath: "/dist/",
        filename: '[name].js',
        library: 'OneNoteApi',
        libraryTarget: 'umd',
        umdNamedDefine: true
    },
    target: 'web',
    devtool: 'cheap-module-eval-source-map',
    resolve: {
        extensions: ['.js', '.ts', '.tsx'],
    },
    plugins: [
        new webpack.ProvidePlugin({
            Promise: 'es6-promise'
        })
    ],
    module: {
        rules: [
            {
                test: /\.ts$/,
                use: 'ts-loader'
            },
        ]
    }
};

const prod = {
    devtool: 'source-map',
    plugins: [
        new webpack.DefinePlugin({
            'process.env': {
                'NODE_ENV': JSON.stringify('production')
            }
        }),
        new DtsBundlePlugin({
            name: 'OneNoteApi',
            main: path.resolve(__dirname, 'dist', 'types', 'scripts', 'oneNoteApi.d.ts'),
            out: path.resolve(__dirname, 'dist', 'OneNoteApi.d.ts'),
            verbose: true,
            removeSource: false,
            outputAsModuleFolder: true,
            emitOnIncludedFileNotFound: true,
            headerText: `TypeScript Definition for OneNoteApi`
        })
    ]
};

let webpackConfiguration = base;

if (IS_PROD) {
    webpackConfiguration = WebpackMerge(webpackConfiguration, prod);
}

module.exports = webpackConfiguration;