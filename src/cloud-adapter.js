import { parseXlsx, view, mutate, exportXlsx } from './workbook-engine.js';

let current = null;
let pendingImport = null;
let authGeneration = 0;
const subscribers = new Set();
const config = () => window.MAESTRO_CONFIG || {};
const fail = (message, status = 400) => Object.assign(new Error(message), {status});
const clone = value => structuredClone(value);
function bytesToBase64(bytes) {
  let result = '';
  for (let at=0; at<bytes.length; at+=32768) result += String.fromCharCode(...bytes.subarray(at,at+32768));
  return btoa(result);
}
function base64ToBytes(text) { return Uint8Array.from(atob(text), character => character.charCodeAt(0)); }
async function digest(text) { return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(text))),x=>x.toString(16).padStart(2,'0')).join(''); }
async function rpc(name, args) {
  const generation = authGeneration;
  const c = config();
  if (!c.supabaseUrl || !c.publishableKey || !c.workspaceId) throw fail('لم يكتمل ربط مساحة العمل بعد.',503);
  const token = await window.MaestroAuth.requestToken();
  requireGeneration(generation);
  const response = await fetch(`${c.supabaseUrl.replace(/\/$/,'')}/rest/v1/rpc/${name}`, {
    method:'POST',cache:'no-store',headers:{apikey:c.publishableKey,Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify(args)
  });
  const value = await response.json().catch(()=>({}));
  requireGeneration(generation);
  if (!response.ok) {
    const message = value.message || value.error || 'تعذر الاتصال بالحفظ المشترك.';
    if (/revision_conflict|conflict|نسخة أحدث/i.test(message)) throw fail('حُفظ تعديل أحدث من جهاز آخر. ألغِ النموذج وحدّث البيانات، ثم أعد التعديل.',409);
    if (response.status === 401) throw fail('انتهت جلسة الدخول. سجّل الدخول مرة أخرى.',401);
    if (/not_authorized|permission|access_denied|membership/i.test(message) || response.status===403) throw fail('حسابك غير مخوّل للوصول إلى مساحة العمل أو تعديلها.',403);
    throw fail(message,response.status);
  }
  return value;
}
function requireGeneration(generation) {
  if (generation!==authGeneration) throw fail('تغيرت جلسة الدخول. سجّل الدخول وحدّث البيانات.',401);
}
function acceptEnvelope(envelope) {
  if (!envelope || typeof envelope!=='object' || Array.isArray(envelope) || !Number.isSafeInteger(envelope.revision) || envelope.revision<0 || !['owner','editor','viewer'].includes(envelope.role) || !Object.hasOwn(envelope,'payload') || (envelope.payload!==null && (!envelope.payload || typeof envelope.payload!=='object' || !envelope.payload.model || !envelope.payload.sourceBase64))) {
    throw fail('لم تصل حالة صالحة من مساحة العمل. حدّث البيانات وأعد المحاولة.',502);
  }
  // An older GET may finish after a successful commit or a newer poll. Keep
  // revisions monotonic so late responses cannot roll back the visible model.
  if (!current || envelope.revision>=current.revision) current=envelope;
  return stateOf();
}
function stateOf(envelope = current) {
  if (!envelope) throw fail('حدّث البيانات أولًا.',409);
  const payload = envelope.payload;
  const result = payload?.model ? view(payload.model) : {sheets:[],warnings:[],cellCount:0,formulaCount:0};
  return {...result,empty:!payload?.model,role:envelope.role,canEdit:['owner','editor'].includes(envelope.role),
    document:{sourceName:payload?.sourceName || '',sourceHash:payload?.sourceHash || '',revision:envelope.revision,importedAt:payload?.importedAt},
    savedAt:envelope.savedAt,history:envelope.history || [],canUndo:!!envelope.canUndo};
}
async function load() {
  return acceptEnvelope(await rpc('maestro_get_state',{p_workspace_id:config().workspaceId}));
}
function expected(revision) {
  if (!current || revision !== current.revision) throw fail('بيانات العرض قديمة. حدّث الصفحة قبل الحفظ.',409);
  if (!['owner','editor'].includes(current.role)) throw fail('صلاحية حسابك للعرض فقط.',403);
}
async function commit(payload,action,detail,revision) {
  expected(revision);
  return acceptEnvelope(await rpc('maestro_commit',{p_workspace_id:config().workspaceId,p_expected_revision:revision,p_payload:payload,p_action:action,p_detail:detail}));
}
function compareModels(before,after) {
  // Stored cell.value is the import-time Excel cache for formula cells. Compare
  // recalculated views so reimporting our own export does not invent changes.
  const entries = model => new Map((model ? view(model).sheets : []).flatMap(sheet=>sheet.rows.flatMap(row=>row.cells).filter(cell=>cell.formula || cell.value!==null && cell.value!==undefined && cell.value!=='').map(cell=>[`${sheet.name}!${cell.ref}`,{value:cell.value??null,formula:cell.formula??null}])));
  const a=entries(before),b=entries(after),changes=[];
  const summary={added:0,removed:0,changed:0,formulasChanged:0,unit:'cells'};
  for (const key of new Set([...a.keys(),...b.keys()])) {
    const first=a.get(key),second=b.get(key);
    if (JSON.stringify(first)===JSON.stringify(second)) continue;
    const kind=!first?'added':!second?'removed':'changed'; summary[kind]++;
    if (first?.formula!==second?.formula && (first?.formula || second?.formula)) summary.formulasChanged++;
    const split=key.lastIndexOf('!'); changes.push({sheet:key.slice(0,split),ref:key.slice(split+1),before:first??null,after:second??null,kind});
  }
  return {summary,changes};
}
async function importPreview(options) {
  const generation=authGeneration;
  expected(current?.revision);
  const data=new Uint8Array(options.body);
  if (data.byteLength>10*1024*1024) throw fail('حجم الملف أكبر من الحد المدعوم (10 ميغابايت).');
  const name=decodeURIComponent(options.headers?.['X-Filename'] || 'updated.xlsx');
  let payload,action='import';
  if (/\.json$/i.test(name)) {
    const backup=JSON.parse(new TextDecoder().decode(data));
    if (backup.format!=='maestro-madinah-backup-v1' || !backup.payload?.sourceBase64 || !backup.payload?.model) throw fail('النسخة الاحتياطية غير مدعومة.');
    if (await digest(JSON.stringify(backup.payload))!==backup.checksum) throw fail('فشل التحقق من سلامة النسخة الاحتياطية.');
    payload=backup.payload;
    const parsed=await parseXlsx(base64ToBytes(payload.sourceBase64),payload.sourceName);
    if (parsed.sourceHash!==payload.model.sourceHash || payload.sourceHash!==parsed.sourceHash || payload.model.version!==1) throw fail('النسخة الاحتياطية لا تطابق مصدرها.');
    view(payload.model);
    await exportXlsx(payload.model,base64ToBytes(payload.sourceBase64));
    action='restore';
  } else {
    const model=await parseXlsx(data,name);
    payload={sourceName:model.sourceName,sourceHash:model.sourceHash,sourceBase64:bytesToBase64(data),model,importedAt:new Date().toISOString()};
  }
  requireGeneration(generation);
  const compared=compareModels(current.payload?.model,payload.model);
  const token=crypto.randomUUID();
  requireGeneration(generation);
  pendingImport={token,payload,action,revision:current.revision,expires:Date.now()+900000};
  return {token,sourceName:payload.sourceName,...compared,warnings:view(payload.model).warnings};
}
async function request(path,options={}) {
  let payload={};
  if (typeof options.body==='string') payload=JSON.parse(options.body);
  switch(path) {
    case '/api/state': return load();
    case '/api/save': expected(payload.revision); return load();
    case '/api/record-preview': {
      expected(payload.revision);
      if (!current.payload) throw fail('استورد ملف المرجع أولًا.');
      return {...view(mutate(current.payload.model,payload.action,payload.groupId,payload.recordId,payload.fields || {})),preview:true};
    }
    case '/api/records': {
      expected(payload.revision);
      if (!current.payload) throw fail('استورد ملف المرجع أولًا.');
      const next=clone(current.payload);
      next.model=mutate(next.model,payload.action,payload.groupId,payload.recordId,payload.fields || {});
      return commit(next,payload.action,{groupId:payload.groupId,recordId:payload.recordId || null},payload.revision);
    }
    case '/api/undo': expected(payload.revision); return acceptEnvelope(await rpc('maestro_undo',{p_workspace_id:config().workspaceId,p_expected_revision:payload.revision}));
    case '/api/import-preview': return importPreview(options);
    case '/api/import-confirm': {
      const p=pendingImport;
      if (!p || p.token!==payload.token || p.expires<Date.now() || p.revision!==payload.revision) throw fail('انتهت المعاينة أو تغيرت البيانات. أعد استيراد الملف لمراجعته.');
      const result=await commit(p.payload,p.action,{sourceName:p.payload.sourceName},payload.revision);
      pendingImport=null; return result;
    }
    default: throw fail('عملية غير معروفة.',404);
  }
}
async function download(kind) {
  const generation=authGeneration;
  // Read the current shared revision so exports include the latest saved data.
  const previous=current?.revision;
  const next=await load();
  requireGeneration(generation);
  // A download can discover a remote save before polling does. Notify the UI
  // now, because subsequent polls will already see this revision in the cache.
  if (previous!==next.document.revision) for (const callback of subscribers) {
    requireGeneration(generation);
    callback(next);
  }
  requireGeneration(generation);
  if (!current.payload) throw fail('استورد ملف المرجع أولًا.');
  const p=current.payload;
  let bytes,type,name;
  if (kind==='original') { bytes=base64ToBytes(p.sourceBase64); type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; name=p.sourceName; }
  else if (kind==='export') { bytes=await exportXlsx(p.model,base64ToBytes(p.sourceBase64)); type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; name=`Maestro_Madinah_${new Date().toISOString().slice(0,10)}.xlsx`; }
  else if (kind==='backup') { bytes=JSON.stringify({format:'maestro-madinah-backup-v1',createdAt:new Date().toISOString(),revision:current.revision,payload:p,history:current.history,checksum:await digest(JSON.stringify(p))},null,2);type='application/json';name=`Maestro_Madinah_${new Date().toISOString().slice(0,10)}.backup.json`; }
  else throw fail('نوع تنزيل غير معروف.');
  requireGeneration(generation);
  const url=URL.createObjectURL(new Blob([bytes],{type}));
  const link=document.createElement('a');link.href=url;link.download=name;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function clear() { authGeneration++;current=null;pendingImport=null; }
function subscribe(callback) { subscribers.add(callback);return ()=>subscribers.delete(callback); }
async function poll() {
  if (!window.MaestroAuth.getSession() || document.hidden || document.querySelector('#modal[open]')) return;
  const previous=current?.revision;
  try { const next=await load(); if (previous!==next.document.revision) for (const callback of subscribers) callback(next); }
  catch(error) { for (const callback of subscribers) callback(null,error); }
}
window.MaestroCloud=Object.freeze({request,download,clear,subscribe,load});
setInterval(poll,15000);
window.addEventListener('online',poll);
