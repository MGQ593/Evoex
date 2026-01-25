const path = require("path");
const fs = require("fs");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const isProduction = process.env.NODE_ENV === "production";

// Cargar variables de entorno desde .env manualmente
function loadEnvVariables() {
  const envPath = path.resolve(__dirname, ".env");
  const envVars = {};

  if (fs.existsSync(envPath)) {
    const content = fs.readFileSync(envPath, "utf-8");
    const lines = content.split("\n");

    for (const line of lines) {
      const trimmed = line.trim();
      // Ignorar comentarios y líneas vacías
      if (!trimmed || trimmed.startsWith("#")) continue;

      const eqIndex = trimmed.indexOf("=");
      if (eqIndex > 0) {
        const key = trimmed.substring(0, eqIndex).trim();
        const value = trimmed.substring(eqIndex + 1).trim();
        envVars[key] = value;
      }
    }
  }

  return envVars;
}

const envVars = loadEnvVariables();

// Generar archivo env.generated.ts con los valores del .env
// Este archivo se importa en config.ts y contiene los valores "horneados"
const envGeneratedPath = path.resolve(__dirname, "src/config/env.generated.ts");
const envGeneratedContent = `// AUTO-GENERADO por webpack - NO EDITAR MANUALMENTE
// Los valores vienen del archivo .env

export const ENV = {
  AZURE_OPENAI_ENDPOINT: ${JSON.stringify(envVars.AZURE_OPENAI_ENDPOINT || "")},
  AZURE_OPENAI_API_KEY: ${JSON.stringify(envVars.AZURE_OPENAI_API_KEY || "")},
  AVAILABLE_MODELS: ${JSON.stringify(envVars.AVAILABLE_MODELS || "")},
  DEFAULT_MODEL_ID: ${JSON.stringify(envVars.DEFAULT_MODEL_ID || "")},
  SEARXNG_PROXY_URL: ${JSON.stringify(envVars.SEARXNG_PROXY_URL || "")},
};
`;
fs.writeFileSync(envGeneratedPath, envGeneratedContent, "utf-8");
console.log("[Webpack] env.generated.ts actualizado");

// Certificados SSL de Office Add-in
const certPath = path.join(require("os").homedir(), ".office-addin-dev-certs");
const httpsOptions = fs.existsSync(path.join(certPath, "localhost.key")) ? {
  key: fs.readFileSync(path.join(certPath, "localhost.key")),
  cert: fs.readFileSync(path.join(certPath, "localhost.crt")),
  ca: fs.readFileSync(path.join(certPath, "ca.crt")),
} : undefined;

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.ts",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js", ".jsx", ".json"],
    alias: {
      "@": path.resolve(__dirname, "src"),
    },
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
      {
        test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
        type: "asset/resource",
        generator: {
          filename: "assets/[name][ext]",
        },
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      filename: "taskpane.html",
      chunks: ["taskpane"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "assets",
          to: "assets",
          noErrorOnMissing: true,
        },
        {
          from: "manifest.xml",
          to: "manifest.xml",
        },
      ],
    }),
    // Inyectar variables de entorno con nombres personalizados
    // Esto evita conflictos con dotenv-webpack y process.env
    new webpack.DefinePlugin({
      __ENV_AZURE_OPENAI_ENDPOINT__: JSON.stringify(envVars.AZURE_OPENAI_ENDPOINT || ""),
      __ENV_AZURE_OPENAI_API_KEY__: JSON.stringify(envVars.AZURE_OPENAI_API_KEY || ""),
      __ENV_AVAILABLE_MODELS__: JSON.stringify(envVars.AVAILABLE_MODELS || ""),
      __ENV_DEFAULT_MODEL_ID__: JSON.stringify(envVars.DEFAULT_MODEL_ID || ""),
      __ENV_SEARXNG_PROXY_URL__: JSON.stringify(envVars.SEARXNG_PROXY_URL || ""),
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    // Seguridad: Solo escuchar en localhost, no en todas las interfaces
    host: "localhost",
    allowedHosts: ["localhost"],
    headers: {
      // Solo permitir origen de Office Add-in y localhost
      "Access-Control-Allow-Origin": "https://localhost:3000",
    },
    server: httpsOptions ? {
      type: "https",
      options: httpsOptions,
    } : "https",
    port: 3000,
    hot: true,
    open: false,
    // Deshabilitar acceso desde la red
    webSocketServer: {
      options: {
        host: "localhost",
      },
    },
  },
  // Usar source-map en ambos modos para que DefinePlugin funcione correctamente
  devtool: "source-map",
  mode: isProduction ? "production" : "development",
};
