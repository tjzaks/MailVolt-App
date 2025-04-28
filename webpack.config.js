const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: {
        taskpane: './taskpane.js'
    },
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: '[name].bundle.js'
    },
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: {
                    loader: 'babel-loader',
                    options: {
                        presets: ['@babel/preset-env']
                    }
                }
            },
            {
                test: /\.css$/,
                use: ['style-loader', 'css-loader']
            }
        ]
    },
    plugins: [
        new CleanWebpackPlugin(),
        new HtmlWebpackPlugin({
            template: './taskpane.html',
            filename: 'taskpane.html',
            chunks: ['taskpane']
        }),
        new CopyWebpackPlugin({
            patterns: [
                {
                    from: 'assets',
                    to: 'assets',
                    noErrorOnMissing: true
                },
                {
                    from: 'manifest.xml',
                    to: '[name][ext]'
                }
            ]
        })
    ],
    devServer: {
        static: {
            directory: path.join(__dirname, 'dist'),
        },
        headers: {
            "Access-Control-Allow-Origin": "*"
        },
        https: true,
        port: 3000
    }
}; 