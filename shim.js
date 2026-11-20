// Gives the billing page the same sharing connector it has on claude.ai, backed by this server.
(function(){
  window.ESI_SERVER=true;
  window.ESI_ME=null;var me=null,mep=fetch('/api/me').then(function(r){return r.ok?r.json():null}).then(function(j){me=j;return j}).catch(function(){return null});window.ESI_ME=mep;
  function call(m,p,b){return fetch('/api/doc?path='+encodeURIComponent(p),{method:m,headers:{'Content-Type':'application/json'},body:b===undefined?undefined:JSON.stringify(b)}).then(function(r){
    if(r.status===401){location.href='/login';throw {code:'signed_out'}}
    return r.json().then(function(j){if(!r.ok)throw {code:r.status===403?'invalid_argument':'unavailable',message:j.error};return j})})}
  function snap(p,j){return {id:p.split('/').pop(),exists:!!j.exists,data:function(){return j.exists?j.data:undefined},metadata:{fromCache:false,hasPendingWrites:false}}}
  var db={doc:function(p){return{
    path:p,
    get:function(){return call('GET',p).then(function(j){return snap(p,j)})},
    set:function(d){return call('PUT',p,d).then(function(){})},
    delete:function(){return call('DELETE',p).then(function(){})},
    onSnapshot:function(next,err){
      var last=null,dead=false,t;
      function tick(){if(dead)return;call('GET',p).then(function(j){var k=j.exists?j.updated:-1;if(k!==last){last=k;next(snap(p,j))}}).catch(function(e){if(err&&e&&e.code!=='unavailable'){dead=true;err(e)}}).then(function(){if(!dead)t=setTimeout(tick,10000)})}
      tick();return function(){dead=true;clearTimeout(t)};
    }}}};
  window.claude={use:function(n){
    if(n==='db')return Promise.resolve(db);
    if(n==='user')return mep.then(function(m){if(!m)return null;return{canEdit:function(){return m.role==='admin'},isOwner:function(){return m.role==='admin'},id:function(){return Promise.resolve(String(m.id))}}});
    return Promise.resolve(null);
  }};
  document.addEventListener('DOMContentLoaded',function(){
    mep.then(function(m){if(!m)return;var h=document.querySelector('.mast-in')||document.querySelector('.mast');if(!h)return;
      var d=document.createElement('div');d.style.cssText='margin-left:auto;display:flex;gap:14px;align-items:center;font-size:13px;color:#c5cde0';
      d.innerHTML='<span></span>'+(window.ESI_LIVE?'<a href="/" style="color:#fff">Billing</a>':'')+(m.role==='admin'?'<a href="/users" style="color:#fff">Users</a>':'')+'<a href="#" id="so" style="color:#fff">Sign out</a>';
      if(!window.ESI_LIVE&&!location.hash&&typeof setTab==='function'){try{setTab('live')}catch(e){}}
      d.firstChild.textContent=m.username+(m.role==='admin'?' · publisher':'');
      if(m.role==='admin'&&!window.ESI_LIVE)fetch('/api/source').then(function(r){return r.json()}).then(function(c){
        var bar=document.getElementById('s1');if(!c.configured||!bar)return;
        var pub=document.getElementById('pubBtn'),msg=document.createElement('span'),loaded=null,files=[],busy=false;msg.className='note-s';
        function status(t){msg.textContent=t}
        function fetchFile(url,after){busy=true;status('Fetching from OneDrive…');
          return fetch(url).then(function(r){if(!r.ok)return r.json().then(function(j){throw new Error(j.error||'failed')});
            var nm=decodeURIComponent(r.headers.get('X-File-Name')||'CCDM_DB.xlsm'),lm=r.headers.get('X-Last-Modified');
            return r.blob().then(function(bl){loaded=nm;status(nm+(lm?' · changed '+new Date(lm).toLocaleString():''));load(new File([bl],nm))})})
          .catch(function(e){status(e.message)}).then(function(){busy=false})}
        if(!c.folder){
          var b=document.createElement('button');b.className='p';b.textContent='Refresh from OneDrive';b.style.marginLeft='auto';
          bar.insertBefore(b,pub);bar.insertBefore(msg,pub);
          b.onclick=function(){b.disabled=true;fetchFile('/api/source/file').then(function(){b.disabled=false})};b.onclick();return}
        var sel=document.createElement('select');sel.id='odMonth';sel.title='Month';sel.style.cssText='min-width:160px';
        var rb=document.createElement('button');rb.textContent='Refresh list';
        bar.insertBefore(sel,pub);bar.insertBefore(rb,pub);bar.insertBefore(msg,pub);
        function label(f){return f.month?new Date(f.month+'-01T00:00:00Z').toLocaleDateString('en-US',{month:'long',year:'numeric',timeZone:'UTC'}):f.name}
        function openFile(f){if(!f||busy||f.name===loaded)return;sel.value=f.id;var c0=document.getElementById('cutoff');return Promise.resolve(fetchFile('/api/source/file?id='+encodeURIComponent(f.id))).then(function(){setTimeout(function(){if(c0&&f.month&&!c0.value){c0.value=f.month;c0.dispatchEvent(new Event('change'))}},800)})}
        function conn(){var a=document.createElement('a');a.href='/auth/login';a.target='_blank';a.className='p';a.textContent='Connect OneDrive';a.style.cssText='background:var(--accent);color:#fff;padding:8px 14px;border-radius:8px;text-decoration:none;font-weight:600';return a}
        function list(force){status('Looking in the OneDrive folder…');
          return fetch('/api/source/list'+(force?'?force=1':'')).then(function(r){return r.json().then(function(j){if(!r.ok){var e=new Error(j.error||'failed');e.connect=j.connect;throw e}return j})}).then(function(fs){
            files=fs;sel.innerHTML=fs.map(function(f){return'<option value="'+f.id+'">'+label(f).replace(/</g,'&lt;')+(f.month?'':' (name has no month)')+'</option>'}).join('');
            if(!fs.length){status('No .xlsm or .xlsx files found in that folder.');return}
            status(fs.length+' monthly file'+(fs.length>1?'s':'')+' found.');return fs})
          .catch(function(e){status(e.message);if(e.connect&&!document.getElementById('odConn')){var a=conn();a.id='odConn';bar.insertBefore(a,msg);status('Connect your OneDrive once, then click Refresh list.')}})}
        sel.onchange=function(){openFile(files.filter(function(f){return f.id===sel.value})[0])};
        rb.onclick=function(){list(true)};
        // choosing another billing month in the generator switches to that month's file
        var cut=document.getElementById('cutoff');
        var cl=cut&&cut.closest('label');if(cl)cl.style.display='none'; // the month is chosen on this Database bar
        if(cut)cut.addEventListener('change',function(){var f=files.filter(function(x){return x.month===cut.value})[0];
          if(f&&f.name!==loaded)openFile(f);else if(!f&&cut.value&&files.length)status('No file for that month in the OneDrive folder.')});
        list(false).then(function(fs){if(fs&&fs.length)openFile(fs[0])});
      });
      h.style.display='flex';h.style.alignItems='center';h.appendChild(d);
      d.querySelector('#so').onclick=function(e){e.preventDefault();fetch('/api/logout',{method:'POST'}).then(function(){location.href='/login'})};
    });
  });
})();
