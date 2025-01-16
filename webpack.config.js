const path = require('path');

module.exports = {
    entry: './src/architect-excel.ts',
    output: {
        filename: 'architect-excel.js',
        path: path.resolve(__dirname, 'docs')
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
