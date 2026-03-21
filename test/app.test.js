import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(process.cwd(), '..');

function read(relativePath) {
  return fs.readFileSync(path.join(root, relativePath), 'utf8');
}

test('required project files exist', () => {
  [
    'Code.gs',
    'index.html',
    'styles.html',
    'scripts.html',
    'print.html',
    'appsscript.json',
    'docs/README.md',
    'docs/AGENTS.md',
  ].forEach((file) => {
    assert.equal(fs.existsSync(path.join(root, file)), true, `${file} should exist`);
  });
});

test('manifest configures the app as a user-accessed web app with required scopes', () => {
  const manifest = JSON.parse(read('appsscript.json'));
  assert.equal(manifest.runtimeVersion, 'V8');
  assert.equal(manifest.webapp.executeAs, 'USER_ACCESSING');
  assert.equal(manifest.webapp.access, 'ANYONE');

  const scopes = manifest.oauthScopes || [];
  [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/userinfo.email',
  ].forEach((scope) => assert.ok(scopes.includes(scope), `missing scope ${scope}`));
});

test('Code.gs includes required authorized users, sheet names, and task labels', () => {
  const code = read('Code.gs');

  [
    'skhun@dublincleaners.com',
    'skhun1@dublincleaners.com',
    'ss.sku@gmail.com',
    'Sales',
    'CSR_Performance',
    'Tasks',
    'Schedule',
    'Competitions',
    'Employee_Of_Week',
    'Employees',
    'Front Counter',
    'Phones',
    'Counter Bags',
    'Wedding Gowns',
    'Shirt Hangers',
    'Shake Shirts',
    'Hand Pressed',
    'Sort Hangers',
    'Paper Hangers',
    'Lobby',
    'Breakroom',
    'Vacuum',
    'Trash',
    '# of Carts',
    'Regina',
    'Shelly',
    'Nellie',
    'Demetria',
    'Dipali',
    'Heather',
    'Lynn',
    'Lisa',
    'Angela',
    'Karmen',
    'Omar',
    'Kaylee',
    'Ingrid',
    'Kelly',
    'Brandy',
    'Cheyenne',
  ].forEach((value) => assert.match(code, new RegExp(value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))));

  assert.match(code, /Session\.getActiveUser\(\)\.getEmail\(\)/);
  assert.match(code, /updateTaskStatus/);
  assert.match(code, /ensureAllSheets_/);
});

test('client HTML includes auth-aware shell, refresh handling, and print mode', () => {
  const indexHtml = read('index.html');
  const scriptsHtml = read('scripts.html');
  const printHtml = read('print.html');

  assert.match(indexHtml, /CSR_DASHBOARD_INITIAL_STATE/);
  assert.match(printHtml, /CSR_DASHBOARD_PRINT_MODE = true/);
  assert.match(scriptsHtml, /setInterval/);
  assert.match(scriptsHtml, /renderUnauthorized/);
  assert.match(scriptsHtml, /updateTask\(/);
  assert.match(scriptsHtml, /Competition Module/);
  assert.match(scriptsHtml, /data-loading-timer/);
  assert.match(scriptsHtml, /formatLoadingDuration/);
  assert.doesNotMatch(scriptsHtml, /Reseed demo data/);
  assert.doesNotMatch(read('Code.gs'), /seedDashboardData/);
});
