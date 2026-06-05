// Test Setup - Debug Script
// Run this to check if everything is configured correctly

const fs = require('fs');
const path = require('path');

console.log('\n=== Migration Toolkit Setup Test ===\n');

// 1. Check if files exist
const files = [
  'main.js',
  'preload.js',
  'renderer.js',
  'index.html',
  'package.json',
  'src/styles/main.css',
  'src/styles/design-system.css',
  'public/icon.ico'
];

console.log('1. Checking required files...');
let allFilesExist = true;
files.forEach(file => {
  const exists = fs.existsSync(path.join(__dirname, file));
  console.log(`   ${exists ? '✓' : '✗'} ${file}`);
  if (!exists) allFilesExist = false;
});

if (!allFilesExist) {
  console.log('\n❌ Missing files detected!\n');
  process.exit(1);
}

// 2. Check PowerShell scripts path
console.log('\n2. Checking PowerShell scripts...');
const psPath = path.join(__dirname, '..', 'VGMigrations');
const psScripts = ['menu.ps1', 'discovery-menu.ps1', 'Domain-Removal-Workflow.ps1'];

console.log(`   Scripts path: ${psPath}`);
psScripts.forEach(script => {
  const scriptPath = path.join(psPath, script);
  const exists = fs.existsSync(scriptPath);
  console.log(`   ${exists ? '✓' : '✗'} ${script}`);
});

// 3. Check node_modules
console.log('\n3. Checking dependencies...');
const hasNodeModules = fs.existsSync(path.join(__dirname, 'node_modules'));
console.log(`   ${hasNodeModules ? '✓' : '✗'} node_modules`);

if (!hasNodeModules) {
  console.log('\n⚠️  Dependencies not installed. Run: npm install\n');
}

const hasElectron = fs.existsSync(path.join(__dirname, 'node_modules', 'electron'));
console.log(`   ${hasElectron ? '✓' : '✗'} electron`);

// 4. Check package.json
console.log('\n4. Checking package.json...');
const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
console.log(`   Name: ${pkg.name}`);
console.log(`   Version: ${pkg.version}`);
console.log(`   Main: ${pkg.main}`);

console.log('\n=== Test Complete ===\n');

if (allFilesExist && hasNodeModules && hasElectron) {
  console.log('✓ Everything looks good!\n');
  console.log('To start the app, run: npm start\n');
} else {
  console.log('❌ Issues detected. Fix the problems above.\n');
}
