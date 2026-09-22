#!/usr/bin/env node
/**
 * run.js — runs tests/Tests.js under Node against the Apps Script mocks.
 *
 * Every file in src/server is evaluated into ONE shared vm context, which is exactly
 * how Apps Script loads a project: no module system, one global scope. The code under
 * test is therefore byte-for-byte the code that gets deployed.
 *
 * Usage:
 *   node tests/node/run.js              run every test
 *   node tests/node/run.js --case 18    run acceptance test 18 only
 *   node tests/node/run.js --grep quote run tests whose name contains "quote"
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createGasEnvironment } = require('./mocks/GasMocks');

const ROOT = path.resolve(__dirname, '..', '..');
const SERVER_DIR = path.join(ROOT, 'src', 'server');

/**
 * Files load in plain alphabetical order, on purpose.
 *
 * Apps Script decides for itself in what order it evaluates the files of a
 * project, and alphabetical is the worst case this codebase is likely to meet
 * (it puts Rules.js before StatusEngine.js, and Api.js before everything it
 * calls). Loading that way here means a module that reaches for another module
 * at load time fails in this suite instead of failing in production, where it
 * would take the entire script down rather than one feature.
 */
function serverFiles() {
  return fs.readdirSync(SERVER_DIR)
    .filter((f) => f.endsWith('.js'))
    .sort()
    .map((f) => path.join(SERVER_DIR, f));
}

/**
 * The client HTML files, keyed by the name clasp gives them in the Apps Script
 * project: the path below `src` with the trailing `.html` removed. So
 * `src/client/Index.html` becomes `client/Index`, and `src/client/App.js.html`
 * becomes `client/App.js` — only the final extension is stripped.
 */
function clientHtmlFiles() {
  const root = path.join(ROOT, 'src');
  const files = {};
  (function walk(dir) {
    fs.readdirSync(dir, { withFileTypes: true }).forEach((entry) => {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) return walk(full);
      if (!entry.name.endsWith('.html')) return;
      const name = path.relative(root, full).split(path.sep).join('/').replace(/\.html$/, '');
      files[name] = fs.readFileSync(full, 'utf8');
    });
  })(path.join(root, 'client'));
  return files;
}

function buildContext() {
  const env = createGasEnvironment();
  const sandbox = Object.assign({}, env);
  sandbox.globalThis = sandbox;
  const context = vm.createContext(sandbox);

  serverFiles().forEach((file) => {
    const code = fs.readFileSync(file, 'utf8');
    try {
      vm.runInContext(code, context, { filename: path.relative(ROOT, file) });
    } catch (e) {
      console.error(`Failed loading ${path.relative(ROOT, file)}: ${e.message}`);
      throw e;
    }
  });

  // The mock calls back into the loaded script's include(), and serves the real
  // client files under their real names, so a template test exercises the whole chain.
  const htmlFiles = clientHtmlFiles();
  sandbox.__test.setGlobalScope(sandbox);
  sandbox.__test.setHtmlFiles(htmlFiles);
  sandbox.__test.defaultHtmlFiles = htmlFiles;

  const testFile = path.join(ROOT, 'tests', 'Tests.js');
  vm.runInContext(fs.readFileSync(testFile, 'utf8'), context, { filename: 'tests/Tests.js' });

  return { context, sandbox, env };
}

function parseArgs(argv) {
  const opts = { case: null, grep: null };
  for (let i = 2; i < argv.length; i++) {
    if (argv[i] === '--case') opts.case = argv[++i];
    else if (argv[i] === '--grep') opts.grep = argv[++i];
  }
  return opts;
}

function main() {
  const opts = parseArgs(process.argv);
  const { sandbox } = buildContext();

  if (typeof sandbox.runAllTests !== 'function') {
    console.error('tests/Tests.js did not define runAllTests()');
    process.exit(1);
  }

  const summary = sandbox.runAllTests({ caseNumber: opts.case, grep: opts.grep });

  const width = summary.results.reduce((m, r) => Math.max(m, r.name.length), 0);
  summary.results.forEach((r) => {
    const mark = r.ok ? '\x1b[32mPASS\x1b[0m' : '\x1b[31mFAIL\x1b[0m';
    console.log(`${mark}  ${r.name.padEnd(width)}  ${r.ms}ms`);
    if (!r.ok) {
      console.log(`      ${r.error}`);
      if (r.stack) console.log(r.stack.split('\n').slice(1, 6).map((l) => '      ' + l.trim()).join('\n'));
    }
  });

  console.log('');
  console.log(`${summary.passed} passed, ${summary.failed} failed, ${summary.results.length} total`);
  process.exit(summary.failed === 0 ? 0 : 1);
}

main();
