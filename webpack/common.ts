import { CleanWebpackPlugin } from "clean-webpack-plugin";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import * as path from "path";
import * as webpack from "webpack";

export const BASE_PATH = path.resolve(__dirname, "..");
export const PATH_SRC = path.resolve(BASE_PATH, "src");
export const PATH_OFFICE = path.resolve(BASE_PATH, "office");

const config: webpack.Configuration = {
  entry: {
    dialog: path.resolve(PATH_SRC, "dialog", "index.tsx"),
    taskpane: path.resolve(PATH_SRC, "taskpane", "index.tsx"),
  },
  output: {
    filename: "[name].[contenthash].bundle.js",
    path: path.resolve(BASE_PATH, "dist"),
  },
  optimization: {
    splitChunks: {
      chunks: "all",
      cacheGroups: {
        defaultVendors: {
          test: /[\\/]node_modules[\\/]/,
          priority: -10,
        },
        default: {
          minChunks: 2,
          priority: -20,
          reuseExistingChunk: true,
        },
      },
    },
  },
  module: {
    rules: [
      {
        test: /\.(tsx?|jsx?)$/,
        exclude: /node_modules/,
        loader: "ts-loader",
      },
      {
        test: /\.(png|svg|jpg|gif)$/,
        use: ["file-loader"],
      },
    ],
  },
  plugins: [
    new CleanWebpackPlugin(),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: path.resolve(PATH_OFFICE, "icons"),
          to: "icons",
        },
        {
          from: path.resolve(PATH_OFFICE, "manifest.xml"),
          to: "[name].[ext]",
        },
      ],
    }),
    new HtmlWebpackPlugin({
      filename: "dialog.html",
      template: path.resolve(PATH_OFFICE, "pages", "dialog.html"),
      chunks: ["dialog"],
    }),
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: path.resolve(PATH_OFFICE, "pages", "taskpane.html"),
      chunks: ["taskpane"],
    }),
  ],
  resolve: {
    extensions: [".ts", ".tsx", ".js", ".jsx", ".json"],
  },
};

export default config;
