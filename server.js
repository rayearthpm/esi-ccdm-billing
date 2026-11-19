// ESI CCDM Billing server. No npm packages needed: Node 22.5+ only.
'use strict';
const http=require('node:http'),fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const {DatabaseSync}=require('node:sqlite');
const PORT=+process.env.PORT||3000,HOST=process.env.HOST||'::';
const DATA=process.env.DATA_DIR||path.join(__dirname,'data');
fs.mkdirSync(DATA,{recursive:true});
const db=new DatabaseSync(path.join(DATA,'billing.db'));
db.exec(`PRAGMA journal_mode=WAL;
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,username TEXT UNIQUE NOT NULL COLLATE NOCASE,salt TEXT NOT NULL,hash TEXT NOT NULL,role TEXT NOT NULL CHECK(role IN('admin','viewer')),created TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,user_id INTEGER NOT NULL,expires INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS docs(path TEXT PRIMARY KEY,body TEXT NOT NULL,updated INTEGER NOT NULL);`);
const q={
  userBy:db.prepare('SELECT * FROM users WHERE username=?'),
  userId:db.prepare('SELECT * FROM users WHERE id=?'),
  users:db.prepare('SELECT id,username,role,created FROM users ORDER BY username'),
  addUser:db.prepare('INSERT INTO users(username,salt,hash,role,created) VALUES(?,?,?,?,?)'),
  setPw:db.prepare('UPDATE users SET salt=?,hash=? WHERE id=?'),
  delUser:db.prepare('DELETE FROM users WHERE id=?'),
  admins:db.prepare("SELECT COUNT(*) n FROM users WHERE role='admin'"),
  addSess:db.prepare('INSERT INTO sessions VALUES(?,?,?)'),
  sess:db.prepare('SELECT u.id,u.username,u.role FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=? AND s.expires>?'),
  delSess:db.prepare('DELETE FROM sessions WHERE token=?'),
  delUserSess:db.prepare('DELETE FROM sessions WHERE user_id=?'),
  purge:db.prepare('DELETE FROM sessions WHERE expires<=?'),
  getDoc:db.prepare('SELECT body,updated FROM docs WHERE path=?'),
  setDoc:db.prepare('INSERT INTO docs(path,body,updated) VALUES(?,?,?) ON CONFLICT(path) DO UPDATE SET body=excluded.body,updated=excluded.updated'),
  delDoc:db.prepare('DELETE FROM docs WHERE path=?'),
};
const hashPw=(pw,salt)=>crypto.scryptSync(pw,salt,64).toString('hex');
function addUser(name,pw,role){const salt=crypto.randomBytes(16).toString('hex');q.addUser.run(name,salt,hashPw(pw,salt),role,new Date().toISOString())}
function setPassword(id,pw){const salt=crypto.randomBytes(16).toString('hex');q.setPw.run(salt,hashPw(pw,salt),id);q.delUserSess.run(id)}
function checkPw(u,pw){const a=Buffer.from(hashPw(pw,u.salt),'hex'),b=Buffer.from(u.hash,'hex');return a.length===b.length&&crypto.timingSafeEqual(a,b)}

// CLI: node server.js adduser <name> <password> [admin|viewer]  |  passwd <name> <password>
if(process.argv[2]==='adduser'){const[,, , n,p,r]=process.argv;if(!n||!p){console.log('usage: node server.js adduser <name> <password> [admin|viewer]');process.exit(1)}
  try{addUser(n,p,r==='admin'?'admin':'viewer');console.log('created',n,r==='admin'?'(admin)':'(viewer)')}catch(e){console.log('could not create:',e.message)}process.exit(0)}
if(process.argv[2]==='passwd'){const u=q.userBy.get(process.argv[3]);if(!u||!process.argv[4]){console.log('usage: node server.js passwd <name> <newpassword>');process.exit(1)}setPassword(u.id,process.argv[4]);console.log('password changed');process.exit(0)}

if(!q.admins.get().n){
  const name=process.env.ADMIN_USER||'admin',pw=process.env.ADMIN_PASSWORD||crypto.randomBytes(6).toString('base64url');
  addUser(name,pw,'admin');
  console.log('\n  First run: created the publisher account.\n  Username: '+name+'\n  Password: '+pw+'\n  Sign in and change it from the Users page.\n');
}


// ---- OneDrive / SharePoint source for the monthly workbook ----
let CFG={};try{CFG=JSON.parse(fs.readFileSync(path.join(DATA,'config.json'),'utf8'))}catch(e){}
const cfg=k=>process.env[k]||CFG[k]||'';
const b64u=s=>Buffer.from(s).toString('base64').replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');
async function graphToken(){
  const r=await fetch(`https://login.microsoftonline.com/${encodeURIComponent(cfg('MS_TENANT_ID'))}/oauth2/v2.0/token`,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},signal:AbortSignal.timeout(20000),
    body:new URLSearchParams({client_id:cfg('MS_CLIENT_ID'),client_secret:cfg('MS_CLIENT_SECRET'),scope:'https://graph.microsoft.com/.default',grant_type:'client_credentials'})});
  const j=await r.json();if(!r.ok)throw new Error('Microsoft sign-in failed: '+(j.error_description||j.error||r.status));return j.access_token;
}
async function fetchWorkbook(){
  const link=cfg('ONEDRIVE_LINK');if(!link)throw new Error('No OneDrive link is set on the server.');
  const tries=[],host=(()=>{try{return new URL(link).hostname}catch(e){return ''}})();
  if(cfg('MS_CLIENT_ID')&&cfg('MS_CLIENT_SECRET')&&cfg('MS_TENANT_ID')){
    tries.push(async()=>fetch('https://graph.microsoft.com/v1.0/shares/u!'+b64u(link)+'/driveItem/content',{headers:{Authorization:'Bearer '+await graphToken()},signal:AbortSignal.timeout(120000)}));
  }
  if(/(^|\.)(1drv\.ms|onedrive\.live\.com)$/.test(host))tries.push(()=>fetch('https://api.onedrive.com/v1.0/shares/u!'+b64u(link)+'/root/content',{signal:AbortSignal.timeout(120000)}));
  tries.push(()=>fetch(link+(link.includes('?')?'&':'?')+'download=1',{signal:AbortSignal.timeout(120000)}));
  let last='';
  for(const t of tries){
    try{
      const r=await t();if(!r.ok){last='HTTP '+r.status;continue}
      const buf=Buffer.from(await r.arrayBuffer());
      if(buf.length<4||buf[0]!==0x50||buf[1]!==0x4b){last='the link returned a web page, not the workbook (is it set to "anyone with the link can view"?)';continue}
      const cd=r.headers.get('content-disposition')||'',m=cd.match(/filename\*?=(?:UTF-8'')?"?([^";]+)/i);
      return{buf,name:cfg('ONEDRIVE_NAME')||(m&&decodeURIComponent(m[1]))||'CCDM_DB.xlsm',modified:r.headers.get('last-modified')||''};
    }catch(e){last=e.message}
  }
  throw new Error('Could not download the workbook: '+last);
}



// ---- Sign in to a personal OneDrive once (delegated, PKCE). No public link needed. ----
const TOKFILE=path.join(DATA,'ms_token.json');
const AUTHORITY=()=>`https://login.microsoftonline.com/${cfg('MS_AUTHORITY')||'consumers'}/oauth2/v2.0`;
const REDIRECT=()=>cfg('MS_REDIRECT')||`http://localhost:${PORT}/auth/callback`;
const PENDING=new Map(); // state -> code_verifier
let ACCESS={token:'',exp:0};
const delegated=()=>!!cfg('MS_CLIENT_ID')&&!cfg('MS_CLIENT_SECRET');
const loadTok=()=>{try{return JSON.parse(fs.readFileSync(TOKFILE,'utf8'))}catch(e){return null}};
async function tokenCall(params){
  const r=await fetch(AUTHORITY()+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},signal:AbortSignal.timeout(30000),body:new URLSearchParams(Object.assign({client_id:cfg('MS_CLIENT_ID'),scope:'offline_access Files.Read'},params))});
  const j=await r.json().catch(()=>({}));if(!r.ok)throw new Error('Microsoft sign-in failed: '+(j.error_description||j.error||r.status));
  if(j.refresh_token)fs.writeFileSync(TOKFILE,JSON.stringify({refresh_token:j.refresh_token,saved:new Date().toISOString()}),{mode:0o600});
  ACCESS={token:j.access_token,exp:Date.now()+(j.expires_in-120)*1000};return ACCESS.token;
}
async function delegatedToken(){
  if(ACCESS.token&&Date.now()<ACCESS.exp)return ACCESS.token;
  const t=loadTok();if(!t)throw new Error('NOT_CONNECTED');
  return tokenCall({grant_type:'refresh_token',refresh_token:t.refresh_token});
}

// ---- OneDrive folder: one workbook per month ----
const MON=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
function monthOf(name){let m=name.match(/(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*[ _\-.]*(20\d\d)/i);if(m)return m[2]+'-'+String(MON.indexOf(m[1].toUpperCase())+1).padStart(2,'0');
  m=name.match(/(20\d\d)[ _\-.]?(0[1-9]|1[0-2])(?!\d)/);return m?m[1]+'-'+m[2]:''}
let FOLDER_CACHE={at:0,items:[]};
function listLocal(dir){
  let names;try{names=fs.readdirSync(dir)}catch(e){throw new Error('Cannot open the folder '+dir+' ('+(e.code==='ENOENT'?'it does not exist':e.message)+'). Check LOCAL_FOLDER in data/config.json.')}
  const items=[];
  names.forEach(n=>{if(!/\.(xlsm|xlsx)$/i.test(n)||n.startsWith('~$'))return;let st;try{st=fs.statSync(path.join(dir,n))}catch(e){return}if(!st.isFile())return;
    items.push({id:n,name:n,month:monthOf(n),modified:st.mtime.toISOString(),file:path.join(dir,n)})});
  items.sort((a,b)=>(b.month||'').localeCompare(a.month||'')||(b.modified||'').localeCompare(a.modified||''));
  return items;
}
async function listFolder(force){
  if(cfg('LOCAL_FOLDER'))return listLocal(cfg('LOCAL_FOLDER')); // a folder on this PC (for example a synced OneDrive folder)
  const link=cfg('ONEDRIVE_FOLDER'),dpath=cfg('ONEDRIVE_PATH');
  if(!link&&!(delegated()&&dpath))throw new Error('No OneDrive folder is set on the server.');
  if(!force&&Date.now()-FOLDER_CACHE.at<60000)return FOLDER_CACHE.items;
  const enc=b64u(link),host=(()=>{try{return new URL(link).hostname}catch(e){return ''}})(),graph=cfg('MS_CLIENT_ID')&&cfg('MS_CLIENT_SECRET')&&cfg('MS_TENANT_ID');
  let url,headers={};
  if(delegated()&&dpath){url='https://graph.microsoft.com/v1.0/me/drive/root:/'+dpath.replace(/^\/+|\/+$/g,'').split('/').map(encodeURIComponent).join('/')+':/children?$top=200';headers={Authorization:'Bearer '+await delegatedToken()}}
  else if(graph){url='https://graph.microsoft.com/v1.0/shares/u!'+enc+'/driveItem/children?$top=200';headers={Authorization:'Bearer '+await graphToken()}}
  else if(/(^|\.)(1drv\.ms|onedrive\.live\.com)$/.test(host))url='https://api.onedrive.com/v1.0/shares/u!'+enc+'/root/children?$top=200';
  else throw new Error('A SharePoint / work OneDrive folder cannot be listed with a public link. Ask your IT admin for the app details (see README, "Private").');
  const items=[];
  for(let pages=0;url&&pages<10;pages++){
    const r=await fetch(url,{headers,signal:AbortSignal.timeout(30000)}),j=await r.json().catch(()=>({}));
    if(!r.ok&&delegated()&&r.status===404)throw new Error('The folder path was not found in your OneDrive: '+dpath);
    if(!r.ok)throw new Error('Could not open the folder: '+((j.error&&(j.error.message||j.error.code))||r.status)+'. Is the folder shared as "Anyone with the link can view"?');
    (j.value||[]).forEach(v=>{if(v.file&&/\.(xlsm|xlsx)$/i.test(v.name)&&!v.name.startsWith('~$'))items.push({id:v.id,name:v.name,month:monthOf(v.name),modified:v.lastModifiedDateTime||'',url:v['@microsoft.graph.downloadUrl']||v['@content.downloadUrl']||''})});
    url=j['@odata.nextLink']||null;
  }
  items.sort((a,b)=>(b.month||'').localeCompare(a.month||'')||(b.modified||'').localeCompare(a.modified||''));
  FOLDER_CACHE={at:Date.now(),items};return items;
}

const SESSION_MS=1000*60*60*12;
const cookie=(req)=>Object.fromEntries((req.headers.cookie||'').split(';').map(s=>s.trim().split('=')).filter(a=>a[0]).map(a=>[a[0],a.slice(1).join('=')]));
function authed(req){q.purge.run(Date.now());const t=cookie(req).sid;return t?q.sess.get(t,Date.now())||null:null}
const secure=(req)=>req.headers['x-forwarded-proto']==='https'||req.socket.encrypted;
const send=(res,code,body,type,extra)=>{res.writeHead(code,Object.assign({'Content-Type':type||'application/json; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff','X-Frame-Options':'DENY','Referrer-Policy':'same-origin'},extra||{}));res.end(typeof body==='string'||Buffer.isBuffer(body)?body:JSON.stringify(body))};
const json=(res,code,o)=>send(res,code,o);
function body(req,limit=1500000){return new Promise((ok,no)=>{let n=0,c=[];req.on('data',d=>{n+=d.length;if(n>limit){no(new Error('too large'));req.destroy()}else c.push(d)});req.on('end',()=>ok(Buffer.concat(c).toString('utf8')));req.on('error',no)})}
const fails=new Map(); // login throttling per ip+user
function throttled(k){const f=fails.get(k);return f&&f.n>=5&&Date.now()<f.until}
function noteFail(k){const f=fails.get(k)||{n:0,until:0};f.n++;if(f.n>=5)f.until=Date.now()+5*60*1000;fails.set(k,f)}
const PATH_OK=/^[A-Za-z0-9_\-.~:@+]{1,200}(\/[A-Za-z0-9_\-.~:@+]{1,200}){0,6}$/;
const page=f=>fs.readFileSync(path.join(__dirname,'public',f),'utf8');
const sameOrigin=(req)=>{const o=req.headers.origin;return !o||new URL(o).host===req.headers.host};

const server=http.createServer(async(req,res)=>{
  try{
    const url=new URL(req.url,'http://x'),p=url.pathname,me=authed(req);
    if(req.method!=='GET'&&!sameOrigin(req))return json(res,403,{error:'bad origin'});

    if(p==='/login'&&req.method==='GET')return send(res,200,page('login.html'),'text/html; charset=utf-8');
    if(p==='/api/login'&&req.method==='POST'){
      const d=JSON.parse(await body(req,5000)||'{}'),k=req.socket.remoteAddress+'|'+String(d.username||'').toLowerCase();
      if(throttled(k))return json(res,429,{error:'Too many attempts. Try again in a few minutes.'});
      const u=q.userBy.get(String(d.username||''));
      if(!u||!checkPw(u,String(d.password||''))){noteFail(k);return json(res,401,{error:'Wrong username or password.'})}
      fails.delete(k);const t=crypto.randomBytes(32).toString('base64url');q.addSess.run(t,u.id,Date.now()+SESSION_MS);
      return send(res,200,{ok:true},null,{'Set-Cookie':`sid=${t}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${SESSION_MS/1000}${secure(req)?'; Secure':''}`});
    }
    if(p==='/api/logout'&&req.method==='POST'){const t=cookie(req).sid;if(t)q.delSess.run(t);return send(res,200,{ok:true},null,{'Set-Cookie':'sid=; HttpOnly; Path=/; Max-Age=0'})}

    if(p==='/auth/callback'&&req.method==='GET'){ // no session cookie needed: Microsoft returns here, possibly on a different host name
      const st=url.searchParams.get('state'),ver=PENDING.get(st);PENDING.delete(st);
      const page=(t,ok)=>send(res,200,`<!doctype html><meta charset=utf-8><body style="font:16px system-ui;padding:40px;max-width:520px;margin:auto"><h2>${ok?'OneDrive connected':'Could not connect'}</h2><p>${t}</p>`,'text/html; charset=utf-8');
      if(!ver||!url.searchParams.get('code'))return page('The sign-in was cancelled or expired. '+(url.searchParams.get('error_description')||'Go back to the billing page and click Connect OneDrive again.').replace(/[<>&]/g,''),false);
      try{await tokenCall({grant_type:'authorization_code',code:url.searchParams.get('code'),redirect_uri:REDIRECT(),code_verifier:ver});FOLDER_CACHE.at=0;return page('You can close this tab and go back to the billing page. It will pick up your files.',true)}
      catch(e){return page(String(e.message).replace(/[<>&]/g,''),false)}
    }
    if(!me){ if(p.startsWith('/api/'))return json(res,401,{error:'signed out'}); res.writeHead(302,{Location:'/login'});return res.end()}

    if(p==='/'||p==='/index.html'){ // the app, with the server connector injected before its own script
      let h=page('app.html');const i=h.indexOf('<script>');h=h.slice(0,i)+'<script src="/shim.js"></script>\n'+h.slice(i);
      return send(res,200,h,'text/html; charset=utf-8')}
    if(p==='/shim.js')return send(res,200,fs.readFileSync(path.join(__dirname,'public','shim.js')),'application/javascript; charset=utf-8');
    if(p==='/users')return me.role==='admin'?send(res,200,page('users.html'),'text/html; charset=utf-8'):(res.writeHead(302,{Location:'/'}),res.end());

    if(p==='/auth/login'&&req.method==='GET'){
      if(me.role!=='admin'||!delegated())return json(res,403,{error:'not allowed'});
      const ver=crypto.randomBytes(32).toString('base64url'),st=crypto.randomBytes(16).toString('hex');PENDING.set(st,ver);setTimeout(()=>PENDING.delete(st),600000);
      const ch=crypto.createHash('sha256').update(ver).digest('base64url');
      const u=AUTHORITY()+'/authorize?'+new URLSearchParams({client_id:cfg('MS_CLIENT_ID'),response_type:'code',redirect_uri:REDIRECT(),response_mode:'query',scope:'offline_access Files.Read',state:st,code_challenge:ch,code_challenge_method:'S256'});
      res.writeHead(302,{Location:u});return res.end();
    }
    if(p==='/api/source'&&req.method==='GET')return json(res,200,{configured:me.role==='admin'&&!!(cfg('LOCAL_FOLDER')||cfg('ONEDRIVE_FOLDER')||cfg('ONEDRIVE_LINK')||(delegated()&&cfg('ONEDRIVE_PATH'))),folder:!!(cfg('LOCAL_FOLDER')||cfg('ONEDRIVE_FOLDER')||(delegated()&&cfg('ONEDRIVE_PATH'))),connect:delegated()&&!!cfg('ONEDRIVE_PATH')&&!loadTok()});
    if(p==='/api/source/list'&&req.method==='GET'){
      if(me.role!=='admin')return json(res,403,{error:'Only the publisher can load the database.'});
      try{return json(res,200,(await listFolder(url.searchParams.get('force')==='1')).map(x=>({id:x.id,name:x.name,month:x.month,modified:x.modified})))}catch(e){return json(res,e.message==='NOT_CONNECTED'?409:502,{error:e.message==='NOT_CONNECTED'?'Not connected to OneDrive yet.':e.message,connect:e.message==='NOT_CONNECTED'})}
    }
    if(p==='/api/source/file'&&url.searchParams.get('id')){
      if(me.role!=='admin')return json(res,403,{error:'Only the publisher can load the database.'});
      try{
        let it=(await listFolder(false)).find(x=>x.id===url.searchParams.get('id'));if(!it)it=(await listFolder(true)).find(x=>x.id===url.searchParams.get('id'));
        if(!it||(!it.url&&!it.file))throw new Error('That file was not found in the folder.');
        let data;
        if(it.file)data=fs.readFileSync(it.file); // OneDrive downloads a cloud-only file when it is opened
        else{const r=await fetch(it.url,{signal:AbortSignal.timeout(180000)});if(!r.ok)throw new Error('Download failed (HTTP '+r.status+')');data=Buffer.from(await r.arrayBuffer())}
        return send(res,200,data,'application/octet-stream',{'X-File-Name':encodeURIComponent(it.name),'X-Last-Modified':it.modified})
      }catch(e){return json(res,502,{error:e.message})}
    }
    if(p==='/api/source/file'&&req.method==='GET'){
      if(me.role!=='admin')return json(res,403,{error:'Only the publisher can load the database.'});
      try{const f=await fetchWorkbook();return send(res,200,f.buf,'application/octet-stream',{'X-File-Name':encodeURIComponent(f.name),'X-Last-Modified':f.modified})}
      catch(e){return json(res,502,{error:e.message})}
    }
    if(p==='/api/me')return json(res,200,{id:me.id,username:me.username,role:me.role});
    if(p==='/api/doc'){
      const dp=url.searchParams.get('path')||'';if(!PATH_OK.test(dp))return json(res,400,{error:'bad path'});
      if(req.method==='GET'){const r=q.getDoc.get(dp);return json(res,200,r?{exists:true,data:JSON.parse(r.body),updated:r.updated}:{exists:false})}
      if(me.role!=='admin')return json(res,403,{error:'Only the publisher can change shared results.'});
      if(req.method==='PUT'){const b=await body(req);JSON.parse(b);q.setDoc.run(dp,b,Date.now());return json(res,200,{ok:true})}
      if(req.method==='DELETE'){q.delDoc.run(dp);return json(res,200,{ok:true})}
    }
    if(p.startsWith('/api/users')){
      if(me.role!=='admin')return json(res,403,{error:'admins only'});
      if(req.method==='GET'&&p==='/api/users')return json(res,200,q.users.all());
      if(req.method==='POST'&&p==='/api/users'){const d=JSON.parse(await body(req,5000));const n=String(d.username||'').trim(),pw=String(d.password||'');
        if(!/^[A-Za-z0-9_.@-]{2,40}$/.test(n))return json(res,400,{error:'Username: 2-40 letters, numbers or . _ @ -'});if(pw.length<8)return json(res,400,{error:'Password must be at least 8 characters.'});
        try{addUser(n,pw,d.role==='admin'?'admin':'viewer')}catch(e){return json(res,400,{error:'That username already exists.'})}return json(res,200,{ok:true})}
      const m=p.match(/^\/api\/users\/(\d+)$/);
      if(m){const u=q.userId.get(+m[1]);if(!u)return json(res,404,{error:'no such user'});
        if(req.method==='PUT'){const d=JSON.parse(await body(req,5000));if(String(d.password||'').length<8)return json(res,400,{error:'Password must be at least 8 characters.'});setPassword(u.id,String(d.password));return json(res,200,{ok:true})}
        if(req.method==='DELETE'){if(u.id===me.id)return json(res,400,{error:'You cannot delete your own account.'});if(u.role==='admin'&&q.admins.get().n<2)return json(res,400,{error:'Keep at least one admin.'});q.delUserSess.run(u.id);q.delUser.run(u.id);return json(res,200,{ok:true})}}
    }
    json(res,404,{error:'not found'});
  }catch(e){json(res,400,{error:e.message})}
});
const up=()=>console.log(`ESI CCDM Billing running on http://localhost:${PORT}`);
server.on('error',e=>{
  if(HOST==='::'&&(e.code==='EAFNOSUPPORT'||e.code==='EADDRNOTAVAIL')){server.listen(PORT,'0.0.0.0',up);return} // no IPv6 on this PC
  if(e.code==='EADDRINUSE'){console.log(`\n  Port ${PORT} is already in use. Close the other copy, or start with a different port:  PORT=3001 node server.js\n`);process.exit(1)}
  throw e});
server.listen(PORT,HOST,up);
