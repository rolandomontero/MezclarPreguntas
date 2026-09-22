const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const command = process.argv[2] || "start";
const platform = process.argv[3] || "desktop";
const additionalArguments = process.argv.slice(4);
const manifestPath = path.resolve("manifest.xml");
const temporaryDirectory = path.join(os.tmpdir(), "mezclar-preguntas-office");
const temporaryManifestPath = path.join(temporaryDirectory, "manifest.xml");
const debuggingCli = require.resolve("office-addin-debugging/cli.js");

fs.mkdirSync(temporaryDirectory, { recursive: true });
fs.copyFileSync(manifestPath, temporaryManifestPath);

const cliArguments = [debuggingCli, command, temporaryManifestPath, platform, ...additionalArguments];

if (command === "start") {
  cliArguments.push(
    "--dev-server",
    "node node_modules/webpack/bin/webpack.js serve --mode development"
  );
}

const result = spawnSync(process.execPath, cliArguments, { stdio: "inherit" });

if (result.error) {
  throw result.error;
}

process.exit(result.status ?? 1);