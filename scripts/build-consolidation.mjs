// One-time, reviewable migration. Writes only an audit directory; never the repo data or network.
import fs from 'node:fs';
import path from 'node:path';
import {execFileSync} from 'node:child_process';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';
import {mergeSources,validateData} from '../netlify/functions/data.mjs';
const merge=createRequire(import.meta.url)('../static/project-merge.js');
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const git=(...args)=>execFileSync('git',args,{cwd:root,encoding:'utf8',maxBuffer:30*1024*1024});
const source=git('rev-parse',process.argv[2]||'origin/main').trim();
const audit=JSON.parse(fs.readFileSync(process.argv[3]||path.join(root,'.netlify/qa/consolidation-candidates.json'),'utf8'));
if(!source.startsWith(audit.snapshot))throw Error('Audit source revision mismatch');
const out=path.resolve(root,process.argv[4]||'.netlify/qa/consolidation-20260910');
if(!out.startsWith(path.join(root,'.netlify',path.sep)))throw Error('Output must stay inside .netlify');
fs.mkdirSync(path.join(out,'before'),{recursive:true});fs.mkdirSync(path.join(out,'after'),{recursive:true});
const files=git('ls-tree','--name-only',source).trim().split('\n').filter(f=>/^_\d{4}_.+\.json$/.test(f));
const users=new Map();
for(const file of files){
  const [,month,user]=/^_(\d{4})_(.+)\.json$/.exec(file);
  const data=JSON.parse(git('show',`${source}:${file}`));validateData(data);
  if(!users.has(user))users.set(user,[]);
  users.get(user).push({file,month,data,sha:git('rev-parse',`${source}:${file}`).trim()});
}
const manifest={source_commit:source,created_at:new Date().toISOString(),users:[],deferred:audit.potentialDifferentFieldsPairs||[]};
for(const [user,sources] of users){
  const mergedSources=mergeSources(sources);
  if(mergedSources._warnings.length)throw Error('Unexpected ID collision for '+user);
  const before={projects:mergedSources.projects,daily:mergedSources.daily};
  let after=structuredClone(before);
  const groups=[];
  for(const candidate of audit.groups.filter(g=>g.user===user&&g.status==='recommended')){
    const group={ids:candidate.ids,primary_id:candidate.ids[0],reason:'同じ客先・工事件名・担当者等の連続／重複する月別登録を統合。元月別ファイルは保持。',resolutions:{},daily_overrides:{},documented_loss:candidate.documented_loss||[]};
    // Bee Valley and Bee　Valley differ only in space width. Include June in the July/August continuation.
    if(user==='谷口'&&group.ids.includes('blg910m1')&&group.ids.includes('bos9omwb')){
      group.ids=['65iydlpb',...group.ids];group.primary_id='65iydlpb';
      group.normalizations=group.ids.filter(id=>id!=='65iydlpb').map(id=>({id,field:'partner',from:'Bee　Valley',to:'Bee Valley'}));
      for(const n of group.normalizations){const p=after.projects.find(p=>p.id===n.id);if(p.partner!==n.from)throw Error('Normalization source differs');p.partner=n.to;}
      group.reason+=' 協力会社名の全角空白を半角空白に統一し、6月の同一工事も引き継ぐ。';
    }
    for(const c of candidate.conflicts||[]){
      if(c.resolution?.status!=='recommended')throw Error('Unresolved conflict '+c.date);
      group.resolutions[c.date]=c.resolution.projectId;
    }
    after=merge.merge(after,group.ids,group.resolutions);
    groups.push(group);
  }
  validateData(after);
  const relative=encodeURIComponent(user)+'.json';
  const entry={user,sources:sources.map(({file,sha})=>({file,sha})),before:'before/'+relative,after:'after/'+relative,groups};
  fs.writeFileSync(path.join(out,entry.before),JSON.stringify(before,null,2)+'\n');
  fs.writeFileSync(path.join(out,entry.after),JSON.stringify(after,null,2)+'\n');
  manifest.users.push(entry);
}
fs.writeFileSync(path.join(out,'manifest.json'),JSON.stringify(manifest,null,2)+'\n');
console.log('Audit files generated at '+out);
console.log(JSON.stringify({source,users:manifest.users.length,groups:manifest.users.reduce((n,u)=>n+u.groups.length,0)},null,2));
