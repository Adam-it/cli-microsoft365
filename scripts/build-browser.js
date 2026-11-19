// Builds the browser bundle of the CLI engine: the SharePoint (spo) commands
// and the browser host from dist/browser, bundled with esbuild into
// dist-browser/. Run after 'npm run build' (see the build:browser script).
//
// Options:
//   --sourcemap   also write source maps
//   --no-minify   keep the output readable

import { execSync } from 'child_process';
import esbuild from 'esbuild';
import fs from 'fs';
import { createRequire } from 'module';
import path from 'path';
import url from 'url';

const require = createRequire(import.meta.url);
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
const rootDir = path.join(__dirname, '..');
const distDir = path.join(rootDir, 'dist');
const srcDir = path.join(rootDir, 'src');
const tempDir = path.join(rootDir, '.browser-build');
const outDir = path.join(rootDir, 'dist-browser');
const shimsDir = path.join(distDir, 'browser', 'shims');
const stubsDir = path.join(distDir, 'browser', 'stubs');
const entryName = 'm365-cli';

const sourcemap = process.argv.includes('--sourcemap');
const minify = !process.argv.includes('--no-minify');

// Commands that can't work in the browser for a reason that isn't visible in
// their imports. Key: command name, value: the reason shown to the user.
const denyList = {
};

const localFilesReason = 'it reads or writes local files';

const toPosix = p => p.replace(/\\/g, '/');
const toImportPath = (from, to) => {
  const relativePath = toPosix(path.relative(from, to));
  return relativePath.startsWith('.') ? relativePath : `./${relativePath}`;
};

function selectCommands() {
  const allCommandsFile = path.join(rootDir, 'allCommands.json');
  if (!fs.existsSync(allCommandsFile) || !fs.existsSync(path.join(distDir, 'browser', 'index.js'))) {
    throw new Error(`Run 'npm run build' first: allCommands.json or dist/browser is missing`);
  }

  const allCommands = JSON.parse(fs.readFileSync(allCommandsFile, 'utf8'));
  const kept = [];
  const skipped = [];

  allCommands
    .filter(c => c.name.startsWith('spo '))
    .sort((a, b) => a.name.localeCompare(b.name))
    .forEach(c => {
      // on Windows, paths in allCommands.json contain backslashes
      const command = {
        name: c.name,
        description: c.description,
        file: toPosix(c.file),
        help: c.help ? toPosix(c.help) : undefined,
        aliases: c.aliases
      };

      if (denyList[command.name]) {
        skipped.push({ name: command.name, reason: denyList[command.name] });
        return;
      }

      const source = fs.readFileSync(path.join(srcDir, 'm365', command.file.replace(/\.js$/, '.ts')), 'utf8');
      if (/from\s+['"](node:)?(fs|child_process)['"]/.test(source) ||
        /from\s+['"][^'"]*\/fsUtil\.js['"]/.test(source)) {
        skipped.push({ name: command.name, reason: localFilesReason });
        return;
      }

      kept.push(command);
    });

  return { kept, skipped, total: allCommands.length };
}

function writeBuildInputs(kept, skipped) {
  fs.rmSync(tempDir, { recursive: true, force: true });
  fs.mkdirSync(tempDir, { recursive: true });

  const commandsDir = path.join(distDir, 'm365');
  const loaders = kept
    .map(c => `  ${JSON.stringify(c.file)}: () => import(${JSON.stringify(toImportPath(tempDir, path.join(commandsDir, c.file)))})`)
    .join(',\n');
  fs.writeFileSync(path.join(tempDir, 'commands.generated.js'), `export const infos = ${JSON.stringify(kept, null, 2)};

export const skipped = ${JSON.stringify(skipped, null, 2)};

// keys equal the 'file' values in infos: cli.loadCommandFromArgs passes them
// to cli.loadCommandFromFile
export const loaders = {
${loaders}
};
`);

  // Stands in for the Node.js 'module' module. The CLI uses createRequire to
  // load JSON files, which are bundled here instead.
  const packageJson = JSON.parse(fs.readFileSync(path.join(rootDir, 'package.json'), 'utf8'));
  const { name, version, description, homepage } = packageJson;
  fs.writeFileSync(path.join(tempDir, 'module-shim.js'), `import csom from ${JSON.stringify(toImportPath(tempDir, path.join(rootDir, 'csom.json')))};

const files = {
  'package.json': ${JSON.stringify({ name, version, description, homepage })},
  'csom.json': csom
};

export function createRequire() {
  return id => {
    const fileName = String(id).split(/[\\\\/]/).pop();
    if (Object.prototype.hasOwnProperty.call(files, fileName)) {
      return files[fileName];
    }

    throw new Error(\`require('\${id}') is not available in the browser build\`);
  };
}

export default { createRequire };
`);

  // With code splitting, esbuild gives a dynamically imported CommonJS module
  // only a default export. cli.ts uses jmespath's named export.
  fs.writeFileSync(path.join(tempDir, 'jmespath-wrapper.js'), `import jmespath from 'jmespath';

export const { compile, search, strictDeepEqual, tokenize } = jmespath;
export default jmespath;
`);

  fs.writeFileSync(path.join(tempDir, 'entry.js'), `import { registerCommands } from ${JSON.stringify(toImportPath(tempDir, path.join(distDir, 'browser', 'index.js')))};
import { infos, loaders, skipped } from './commands.generated.js';

registerCommands(infos, loaders, skipped);

export { init, executeCommandLine, listCommands, listSkippedCommands } from ${JSON.stringify(toImportPath(tempDir, path.join(distDir, 'browser', 'index.js')))};
`);
}

function getAliases() {
  const shim = name => path.join(shimsDir, `${name}.js`);
  const stub = name => path.join(stubsDir, `${name}.js`);

  const nodeBuiltIns = {
    fs: shim('fs'),
    os: shim('os'),
    path: shim('path'),
    url: shim('url'),
    module: path.join(tempDir, 'module-shim.js'),
    child_process: shim('child_process'),
    crypto: shim('crypto'),
    stream: shim('stream'),
    'timers/promises': shim('timers-promises'),
    assert: shim('assert'),
    util: shim('util'),
    net: shim('empty'),
    http: shim('empty'),
    https: shim('empty'),
    querystring: shim('empty'),
    // the npm package, not the Node.js built-in
    buffer: require.resolve('buffer/')
  };

  const aliases = {
    configstore: stub('configstore'),
    chalk: stub('chalk'),
    '@azure/msal-common': stub('msal-common'),
    '@azure/msal-node': stub('empty'),
    '@azure/msal-node-extensions': stub('empty'),
    // only used to sign in with a certificate
    'node-forge': stub('empty'),
    'proper-lockfile': stub('proper-lockfile'),
    open: stub('empty'),
    clipboardy: stub('empty'),
    'update-notifier': stub('empty'),
    omelette: stub('empty'),
    '@inquirer/confirm': stub('inquirer'),
    '@inquirer/input': stub('inquirer'),
    '@inquirer/select': stub('inquirer')
  };

  Object.keys(nodeBuiltIns).forEach(name => {
    aliases[name] = nodeBuiltIns[name];
    aliases[`node:${name}`] = nodeBuiltIns[name];
  });

  return aliases;
}

// src/telemetry.ts is imported with a relative path, which aliases don't
// cover. Replacing it keeps Google Analytics and the utils it needs (pid,
// cache, session) out of the bundle.
const replaceTelemetry = {
  name: 'replace-telemetry',
  setup(build) {
    const telemetryFile = path.join(distDir, 'telemetry.js');
    build.onResolve({ filter: /telemetry\.js$/ }, args => {
      if (args.kind !== 'entry-point' &&
        path.resolve(args.resolveDir, args.path) === telemetryFile) {
        return { path: path.join(stubsDir, 'telemetry.js') };
      }

      return undefined;
    });
  }
};

const wrapJmespath = {
  name: 'wrap-jmespath',
  setup(build) {
    const wrapperFile = path.join(tempDir, 'jmespath-wrapper.js');
    build.onResolve({ filter: /^jmespath$/ }, args => {
      return args.importer === wrapperFile ? undefined : { path: wrapperFile };
    });
  }
};

const adjustZod = {
  name: 'adjust-zod',
  setup(build) {
    // Zod's default export references all its locales, which takes about
    // 300 kB. The CLI only uses English.
    build.onLoad({ filter: /[\\/]zod[\\/]v4[\\/]locales[\\/]index\.js$/ }, () => ({
      contents: 'export { default as en } from "./en.js";',
      loader: 'js'
    }));

    // When a schema is created, Zod checks if it can compile schemas with
    // new Function. The content security policy of browser extensions doesn't
    // allow that and reports each attempt. Turning the compilation off has to
    // happen before the first schema is created, so it's done in Zod itself:
    // with code splitting the order in which other modules run isn't certain.
    build.onLoad({ filter: /[\\/]zod[\\/]v4[\\/]core[\\/]core\.js$/ }, async args => {
      const contents = await fs.promises.readFile(args.path, 'utf8');
      if (!contents.includes('export const globalConfig')) {
        throw new Error(`Zod's globalConfig was not found in ${args.path}`);
      }

      return {
        contents: `${contents}\nglobalConfig.jitless = true;\n`,
        loader: 'js'
      };
    });
  }
};

function getGitCommit() {
  try {
    return execSync('git rev-parse --short HEAD', { cwd: rootDir, stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
  }
  catch {
    return undefined;
  }
}

const formatSize = bytes => bytes > 1024 * 1024 ? `${(bytes / 1024 / 1024).toFixed(2)} MB` : `${(bytes / 1024).toFixed(1)} kB`;

async function build() {
  const { kept, skipped, total } = selectCommands();
  console.log(`Commands: ${kept.length} in the bundle, ${skipped.length} spo commands skipped, ${total} in the CLI`);
  skipped.forEach(c => console.log(`  skipped '${c.name}': ${c.reason}`));

  writeBuildInputs(kept, skipped);
  fs.rmSync(outDir, { recursive: true, force: true });

  const result = await esbuild.build({
    entryPoints: { [entryName]: path.join(tempDir, 'entry.js') },
    outdir: outDir,
    absWorkingDir: rootDir,
    bundle: true,
    platform: 'browser',
    format: 'esm',
    // each command becomes a chunk that's loaded when the command is run
    splitting: true,
    chunkNames: 'chunks/[name]-[hash]',
    target: 'chrome120',
    minify,
    sourcemap,
    metafile: true,
    logLevel: 'warning',
    define: { global: 'globalThis' },
    inject: [
      path.join(shimsDir, 'process-global.js'),
      path.join(shimsDir, 'buffer-global.js')
    ],
    alias: getAliases(),
    plugins: [replaceTelemetry, wrapJmespath, adjustZod]
  });
  fs.writeFileSync(path.join(tempDir, 'meta.json'), JSON.stringify(result.metafile));

  const packageJson = JSON.parse(fs.readFileSync(path.join(rootDir, 'package.json'), 'utf8'));
  const outputs = Object.entries(result.metafile.outputs).filter(([file]) => file.endsWith('.js'));
  const totalSize = outputs.reduce((sum, [, output]) => sum + output.bytes, 0);
  const entryOutput = outputs.find(([, output]) => output.entryPoint);
  // chunks the entry point imports statically: loaded before the first command
  const initialFiles = new Set();
  const addInitial = file => {
    if (initialFiles.has(file)) {
      return;
    }
    initialFiles.add(file);
    result.metafile.outputs[file].imports
      .filter(i => i.kind === 'import-statement')
      .forEach(i => addInitial(i.path));
  };
  addInitial(entryOutput[0]);
  const initialSize = [...initialFiles].reduce((sum, file) => sum + result.metafile.outputs[file].bytes, 0);

  fs.writeFileSync(path.join(outDir, 'skipped-commands.json'), JSON.stringify(skipped, null, 2));
  fs.writeFileSync(path.join(outDir, 'build-info.json'), JSON.stringify({
    name: packageJson.name,
    version: packageJson.version,
    commit: getGitCommit(),
    built: new Date().toISOString(),
    entry: `${entryName}.js`,
    commands: kept.length,
    skippedCommands: skipped.length,
    files: outputs.length,
    totalBytes: totalSize,
    initialBytes: initialSize
  }, null, 2));

  console.log(`Output: ${toPosix(path.relative(rootDir, outDir))}/${entryName}.js, ${outputs.length} files, ${formatSize(totalSize)} in total, ${formatSize(initialSize)} loaded before the first command`);
  console.log('Largest inputs:');
  Object.entries(result.metafile.inputs)
    .sort(([, a], [, b]) => b.bytes - a.bytes)
    .slice(0, 10)
    .forEach(([file, input]) => console.log(`  ${formatSize(input.bytes).padStart(10)}  ${file}`));

  // extension pages don't allow eval, see the content security policy of
  // Manifest V3
  const evalFiles = outputs
    .map(([file]) => file)
    .filter(file => /\beval\(|new Function\(/.test(fs.readFileSync(path.join(rootDir, file), 'utf8')));
  if (evalFiles.length > 0) {
    console.warn(`Warning: eval or new Function found in: ${evalFiles.join(', ')}`);
  }
}

build().catch(err => {
  console.error(err.message ?? err);
  process.exit(1);
});
