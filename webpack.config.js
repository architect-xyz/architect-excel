const path = require('path');

module.exports = {
    entry: {
        taskpane: './src/taskpane.ts',
        functions: './src/functions.ts'
    },
    output: {
        filename: '[name].js',
        path: path.resolve(__dirname, 'docs'),
        library: {
            name: '[name]',
            type: 'umd',
        },
        globalObject: 'this',
    },
    resolve: {
        extensions: ['.ts', '.js'],
    },
    module: {
        rules: [
             {
                test: /\.ts$/,
                use: 'ts-loader',
                exclude: /node_modules/
            },
        ]
    },
    mode: 'production',
    optimization: {
        minimize: true
    }
};
