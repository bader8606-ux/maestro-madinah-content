import { build } from 'esbuild';
import { cp } from 'node:fs/promises';
await build({entryPoints:['src/cloud-adapter.js'],bundle:true,format:'iife',platform:'browser',target:['es2022'],outfile:'dist/cloud.js',minify:true,legalComments:'eof'});
await cp('dist','docs',{recursive:true});
console.log('Built dist/cloud.js and synchronized docs/ for GitHub Pages; no workbook data is included.');
