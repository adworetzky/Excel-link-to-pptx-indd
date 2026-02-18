const path = require("path");

module.exports = {
  entry: "./src/taskpane.js",
  output: {
    filename: "taskpane.bundle.js",
    path: path.resolve(__dirname, "dist"),
    clean: true,
  },
  module: {
    rules: [
      {
        test: /\.(js|jsx)$/,
        exclude: /node_modules/,
        use: {
          loader: "babel-loader",
          options: {
            presets: [
              ["@babel/preset-env", { targets: { esmodules: true } }],
              ["@babel/preset-react", { runtime: "automatic" }],
            ],
          },
        },
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  resolve: {
    extensions: [".js", ".jsx"],
    fallback: {
      // Node builtins — not available in Office JS browser environment
      path: false,
      fs: false,
      os: false,
      crypto: false,
      buffer: false,
      stream: false,
    },
  },
  externals: {
    // Office JS is loaded via CDN script tag in taskpane.html
    // Do not bundle it.
  },
  target: "web",
  performance: {
    hints: false,
  },
};
