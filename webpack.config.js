 module.exports = {
    mode: "development",
    entry: {
        OneNoteApi: "./src/scripts/oneNoteApi"
    },
     resolve: {
         extensions: ['.js', '.ts', '.tsx'],
     },
     module: {
         rules: [
             {
                 test: /\.ts$/,
                 use: 'ts-loader'
             },
         ]
     }
 };
