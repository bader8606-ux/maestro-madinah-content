import { build } from 'esbuild';
await build({entryPoints:['src/cloud-adapter.js'],bundle:true,format:'iife',platform:'browser',target:['es2022'],outfile:'dist/cloud.js',minify:true,legalComments:'eof'});
console.log('Built dist/cloud.js; no workbook data is included.');
