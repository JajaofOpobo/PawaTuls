/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const TerserPlugin = require("terser-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://localhost:3000/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    // Full source maps in dev only: prod maps double peak memory and the
    // production minify workers crash under constrained heaps (see TEST_REPORT).
    devtool: dev ? "source-map" : false,
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts",
      "footnote-dialog": "./src/dialogs/footnote-dialog.tsx",
      "tools-dialog": "./src/dialogs/tools-dialog.tsx",
      "confirm-dialog": "./src/dialogs/confirm-dialog.tsx",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          // Type-checking happens via `npx tsc --noEmit` / `npm test`;
          // the bundler only transpiles (memory + speed).
          use: [{ loader: "ts-loader", options: { transpileOnly: true } }],
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "footnote-dialog.html",
        template: "./src/dialogs/footnote-dialog.html",
        chunks: ["polyfill", "footnote-dialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "tools-dialog.html",
        template: "./src/dialogs/tools-dialog.html",
        chunks: ["polyfill", "tools-dialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "confirm-dialog.html",
        template: "./src/dialogs/confirm-dialog.html",
        chunks: ["polyfill", "confirm-dialog"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/",
            to: "assets/",
            noErrorOnMissing: true,
          },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              if (dev) {
                return content;
              }
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    // Cap minifier workers: jest-worker cloning OOMs on big bundles in this
    // environment — run terser in-process instead (slower, reliable).
    optimization: {
      minimizer: [new TerserPlugin({ parallel: false })],
    },
  };

  // Dev-only concerns stay out of production builds: HTTPS cert access and
  // the devServer block are attached only when serving.
  const isServe = process.argv.some((arg) => arg.includes("serve"));
  if (isServe) {
    config.devServer = {
      server: {
        type: "https",
        options: await getHttpsOptions(),
      },
      port: 3000,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      static: {
        directory: __dirname,
      },
    };
  }

  return config;
};
