/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000";
const urlProd = "https://www.contoso.com"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      assignpane: ["./src/taskpane/Js/taskpane_main.js", "./src/taskpane/HTML/assignsignature.html"],
      editpane: ["./src/taskpane/HTML/editsignature.html"],
      autorun: ["./src/runtime/Js/autorunshared.js", "./src/runtime/HTML/autorunweb.html"],
      autorunshared: ["./src/runtime/Js/autorunshared.js"],
    },
    output: {

      publicPath: "",
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      parser: {
        javascript: {
          dynamicImportMode: 'eager'
        }
      },
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
          },
        },
        // {
        //   test: /\.mjs$/,
        //   use: {
        //     loader: "babel-loader",
        //     options: {
        //       plugins: ["@babel/plugin-transform-async-to-generator"],
        //     },
        //   },
        // },
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
      new HtmlWebpackPlugin({
        filename: "editsignature.html",
        template: "./src/taskpane/HTML/editsignature.html",
        chunks: ["polyfill", "editpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "assignsignature.html",
        template: "./src/taskpane/HTML/assignsignature.html",
        chunks: ["polyfill", "assignpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "autorunweb.html",
        template: "./src/runtime/HTML/autorunweb.html",
        chunks: ["polyfill", "autorun"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          {
            from: "well-known/*",
            to: ".well-known/[name][ext][query]",
          },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/public",
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
