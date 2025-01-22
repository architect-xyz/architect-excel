/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

/* global require, module, process, __dirname */

module.exports = {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
      functions: "./src/functions/functions.ts",
    },
    output: {
      clean: true,
    },  
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },  
    module: {
      rules: [
        {   
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },  
        },  
        {   
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },  
        {   
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },  
        },  
      ],  
    },  
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",

  return config;
        input: "./src/functions/functions.ts",
      }), 
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "functions", "commands"],
      }), 
      new CopyWebpackPlugin({
        patterns: [
          {   
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(urlDev, urlProd);
              }
            },
          },
        ],
      }),
    ]
  };
};
