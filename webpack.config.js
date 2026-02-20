/* eslint-disable no-undef */
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

module.exports = async (env, options) => {
  const dev = options.mode !== "production";

  const config = {
    devtool: "source-map",

    entry: {
      functions: "./src/functions/functions.ts",
      taskpane: "./src/taskpane/taskpane.ts",
    },

    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
    },

    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },

    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
            options: { transpileOnly: true },
          },
        },
      ],
    },

    plugins: [
      // Processes taskpane.html template and injects taskpane.js bundle
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        // In shared-runtime mode the taskpane page IS the functions runtime.
        // functions.js must be bundled here so CustomFunctions.associate() runs
        // inside the shared runtime where CustomFunctions is available.
        chunks: ["functions", "taskpane"],
      }),

      // Copies static assets and the hand-authored functions.json to dist/
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext]",
          },
          {
            // Static metadata file — describes FACTCHECK to Excel.
            // Re-run `npm run gen-metadata` if you change JSDoc parameters.
            from: "src/functions/functions.json",
            to: "functions.json",
          },
          {
            // Minimal HTML page for the dedicated functions runtime iframe.
            from: "src/functions/functions.html",
            to: "functions.html",
          },
        ],
      }),
    ],

    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: dev ? await devCerts.getHttpsServerOptions() : {},
      },
      port: 3000,
      static: {
        directory: path.join(__dirname, "dist"),
      },
    },
  };

  return config;
};
