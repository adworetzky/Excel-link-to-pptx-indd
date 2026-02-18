const path = require("path");

module.exports = {
  entry: "./src/index.js",
  output: {
    filename: "bundle.js",
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
    // UXP provides these as host globals — don't bundle them
    fallback: {
      path: false,
      fs: false,
      os: false,
      crypto: false,
      buffer: false,
    },
  },
  externals: {
    // UXP built-ins are injected by the host at runtime
    uxp: "commonjs uxp",
    photoshop: "commonjs photoshop",
    indesign: "commonjs indesign",
    "fs": "commonjs fs",
  },
  // UXP uses a browser-like global environment
  target: "web",
  performance: {
    hints: false,
  },
};
