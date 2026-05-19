#!/usr/bin/env node
/**
 * sync-template.js
 *
 * Copies the latest source files from src/ and public/ into brand-template/,
 * stripping all Ruff Liners-specific branding and replacing it with generic
 * placeholders so the template stays shippable to other brand owners.
 *
 * Run from the repo root:
 *   node scripts/sync-template.js
 */

const fs   = require('fs');
const path = require('path');

const ROOT     = path.resolve(__dirname, '..');
const TEMPLATE = path.join(ROOT, 'brand-template');

// ─── substitution rules ──────────────────────────────────────────────────────
// Each entry is [search: string|RegExp, replacement: string].
// Applied in order to every synced file.

const GLOBAL_SUBS = [
  // Session-storage cache key
  [/rl_data_v1/g,  'ch_data_v1'],

  // Excel export filename
  [/Ruff-Liners-TikTok-/g, 'Creator-Hub-TikTok-'],

  // Header logo: swap branded <img> for a generic emoji placeholder
  [
    /<img src="\/RuffLinersB\.png" alt="Ruff Liners" style=\{\{height:48,width:48,objectFit:"contain",borderRadius:"50%",background:"#fff",flexShrink:0\}\}\/>/,
    '<div style={{height:48,width:48,borderRadius:"50%",background:"rgba(255,255,255,0.1)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>📦</div>',
  ],

  // Brand name in header (JSX)
  [/>RUFF LINERS</g, '>YOUR BRAND<'],

  // Paw-print icon in the admin sign-in modal → generic box
  [/🐾/g, '📦'],

  // Email placeholder in inline login form
  [/you@ruffliners\.com/g, 'you@yourbrand.com'],
];

// Extra substitutions applied only to specific files
const FILE_SUBS = {
  'Login.tsx': [
    // Stand-alone Login screen brand name
    [/Ruff Liners/g, 'Your Brand'],
    [/you@ruffliners\.com/g, 'you@yourbrand.com'],
  ],
  'favicon.svg': [
    // Monogram initials: RL → YB
    [/>RL</g, '>YB<'],
  ],
  'index.html': [
    [/Ruff Liners Creator Hub/g, 'Creator Hub'],
  ],
};

// ─── files to sync ───────────────────────────────────────────────────────────
// Paths are relative to the repo root. Each is read from ROOT/<path> and
// written to TEMPLATE/<path> after substitution.

const FILES = [
  'src/TikTokShopReporter.tsx',
  'src/AuthWrapper.tsx',
  'src/Login.tsx',
  'src/index.tsx',
  'src/index.css',
  'src/supabaseClient.ts',
  'src/react-app-env.d.ts',
  'src/reportWebVitals.ts',
  'public/index.html',
  'public/favicon.svg',
];

// ─── runner ──────────────────────────────────────────────────────────────────

let synced = 0;
let skipped = 0;

FILES.forEach(relPath => {
  const srcFile  = path.join(ROOT, relPath);
  const destFile = path.join(TEMPLATE, relPath);
  const filename = path.basename(relPath);

  if (!fs.existsSync(srcFile)) {
    console.warn(`  SKIP  ${relPath}  (source not found)`);
    skipped++;
    return;
  }

  let content = fs.readFileSync(srcFile, 'utf8');

  // Apply global subs
  GLOBAL_SUBS.forEach(([search, replacement]) => {
    content = content.replace(search, replacement);
  });

  // Apply per-file subs
  (FILE_SUBS[filename] || []).forEach(([search, replacement]) => {
    content = content.replace(search, replacement);
  });

  fs.mkdirSync(path.dirname(destFile), { recursive: true });
  fs.writeFileSync(destFile, content, 'utf8');
  console.log(`  synced  ${relPath}`);
  synced++;
});

console.log(`\n✓  ${synced} file(s) synced, ${skipped} skipped.`);
console.log('  brand-template/ is up to date.\n');
