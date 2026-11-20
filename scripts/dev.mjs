import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const wrangler = fileURLToPath(new URL("../node_modules/wrangler/bin/wrangler.js", import.meta.url));
const options = process.argv.slice(2)
  .filter((value) => value !== "--strictPort")
  .map((value) => value.replace(/^--host(?==|$)/, "--ip"));
const server = spawn(process.execPath, [wrangler, "dev", "--local", ...options], { stdio: "inherit" });
server.on("error", (error) => {
  console.error(error.message);
  process.exitCode = 1;
});
server.on("exit", (code) => { process.exitCode = code ?? 1; });
