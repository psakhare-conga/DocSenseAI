const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

module.exports = async (env, options) => {
  const isDev = options.mode !== "production";
  let httpsOptions = {};

  if (isDev) {
    try {
      const devCerts = require("office-addin-dev-certs");
      httpsOptions = await devCerts.getHttpsServerOptions();
    } catch (e) {
      console.warn("office-addin-dev-certs not ready — run 'npm run install-certs' first.");
    }
  }

  return {
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts"
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx"]
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"]
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["taskpane"]
      }),
      new HtmlWebpackPlugin({
        template: "./src/commands/commands.html",
        filename: "commands.html",
        chunks: ["commands"]
      }),
      new webpack.DefinePlugin({
        "process.env.API_URL": JSON.stringify(process.env.API_URL || "http://localhost:5000")
      })
    ],
    devServer: isDev
      ? {
          hot: true,
          headers: { "Access-Control-Allow-Origin": "*" },
          server: {
            type: "https",
            options: httpsOptions
          },
          port: 3000
        }
      : {}
  };
};
