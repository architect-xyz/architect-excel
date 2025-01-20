const path = require('path');

module.exports = {
    entry: {
        functions: './src/functions.mts',
        taskpane: './src/taskpane.mts'
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
        extensions: ['.ts', '.js', '.mts'],
    },
    module: {
        rules: [
            {
                test: /\.mts$/,
                use: 'ts-loader',
                exclude: /node_modules/
            }
        ]
    },
    mode: 'production',
    optimization: {
        minimize: true
    }
};
