const path = require('path');

module.exports = {
    entry: './src/architect-excel.ts',
    output: {
        filename: 'taskpane.js',
        path: path.resolve(__dirname, 'public')
    },
    resolve: {
        extensions: ['.ts', '.js']
    },
    module: {
        rules: [
            {
                test: /\.ts$/,
                use: 'ts-loader',
                exclude: /node_modules/
            }
        ]
    },
    mode: 'development'
};
