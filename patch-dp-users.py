#!/usr/bin/env python3
from pathlib import Path
import argparse
import re

MARKER = "DISPOPILOT_DP_USERS_AUTH_V1"

AUTH_JS = r'''
/* DISPOPILOT_DP_USERS_AUTH_V1 */
let dpAuthInfo=null;
let dpAuthInstalled=false;

function dpAuthSiteUrl(){
  try{
    if(window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl){
      return String(window._spPageContextInfo.webAbsoluteUrl).replace(/\/$/,'');
    }
  }catch(_){}
  const m=window.location.pathname.match(/^(.*?\/sites\/[^/]+)/i);
  if(m)return window.location.origin+m[1];
  return window.location.origin;
}

async function dpAuthJson(url){
  const r=await fetch(url,{method:'GET',credentials:'same-origin',headers:{'Accept':'application/json;odata=nometadata'}});
  if(!r.ok)throw new Error('SharePoint '+r.status+' '+r.statusText);
  return await r.json();
}

function dpAuthNorm(v){return String(v||'').trim().toLocaleLowerCase('de')}

function dpAuthBlock(title,message){
  document.body.classList.remove('dp-auth-pending');
  const app=document.querySelector('.app');
  if(app)app.style.display='none';
  let box=document.getElementById('dpAuthBlock');
  if(!box){box=document.createElement('div');box.id='dpAuthBlock';document.body.appendChild(box)}
  box.innerHTML='<div class="dp-auth-card"><div class="dp-auth-logo">DP</div><h2>'+String(title||'Zugriff nicht möglich')+'</h2><p>'+String(message||'')+'</p><p class="dp-auth-small">DispoPilot · kontrollierter Zugriff über Microsoft 365</p></div>';
}

async function dpAuthWaitForState(){
  const until=Date.now()+15000;
  while(Date.now()<until){
    if(typeof state!=='undefined' && state && typeof renderAll==='function')return true;
    await new Promise(r=>setTimeout(r,60));
  }
  return false;
}

function dpAuthDecorateSettings(){
  const sel=document.getElementById('currentUser');
  if(!sel || !dpAuthInfo)return;
  sel.disabled=true;
  sel.innerHTML='<option>'+esc(dpAuthInfo.displayName+' · '+dpAuthInfo.role)+'</option>';
  const card=sel.closest('.card');
  if(card){
    const label=card.querySelector('label');
    if(label)label.textContent='Angemeldeter Microsoft-365-Benutzer';
    let note=card.querySelector('.dp-auth-role-note');
    if(!note){note=document.createElement('div');note.className='small dp-auth-role-note';note.style.marginTop='7px';card.appendChild(note)}
    note.textContent=dpAuthInfo.email+' · Rolle: '+dpAuthInfo.role+(dpAuthInfo.driverName?' · Fahrer: '+dpAuthInfo.driverName:'');
  }
}

function dpAuthDecorateTop(){
  if(!dpAuthInfo)return;
  const status=document.getElementById('storageStatus');
  if(status)status.textContent='Online · '+dpAuthInfo.role+' · '+dpAuthInfo.displayName;
}

function dpAuthDecorateOrders(){
  if(!dpAuthInfo || dpAuthInfo.role!=='Admin')return;
  const chips=document.querySelectorAll('#orders .toolbar .chip');
  if(chips.length){
    chips.forEach(c=>c.classList.remove('active'));
    if(chips[0]){chips[0].textContent='Alle Aufträge · Admin';chips[0].classList.add('active');chips[0].onclick=()=>{orderFilter='all';renderOrders()}}
    if(chips[1])chips[1].style.display='none';
  }
}

function dpAuthInstallRuntimeHooks(){
  if(dpAuthInstalled)return;
  dpAuthInstalled=true;

  const baseRenderSettings=renderSettings;
  renderSettings=function(){baseRenderSettings();dpAuthDecorateSettings()};

  const baseRenderOrders=renderOrders;
  renderOrders=function(){if(dpAuthInfo && dpAuthInfo.role==='Admin')orderFilter='all';baseRenderOrders();dpAuthDecorateOrders()};

  changeUser=async function(){dpAuthDecorateSettings()};

  if(typeof dp13IsDisposition==='function'){
    dp13IsDisposition=function(){return !!dpAuthInfo && (dpAuthInfo.role==='Disposition'||dpAuthInfo.role==='Admin')};
  }
}

async function dpAuthBoot(){
  const ready=await dpAuthWaitForState();
  if(!ready){dpAuthBlock('DispoPilot konnte nicht gestartet werden','Die Anwendung wurde nicht vollständig geladen. Bitte die Seite neu laden.');return}

  try{
    const web=dpAuthSiteUrl();
    const current=await dpAuthJson(web+"/_api/web/currentuser?$select=Id,Title,Email,LoginName");
    const users=await dpAuthJson(web+"/_api/web/lists/getbytitle('DP_Users')/items?$select=Id,Title,Rolle,Aktiv,Fahrername,Benutzer/Id,Benutzer/Title,Benutzer/EMail,Benutzer/Name&$expand=Benutzer&$top=5000");

    const currentEmail=dpAuthNorm(current.Email),currentLogin=dpAuthNorm(current.LoginName);
    const row=(users.value||[]).find(x=>{
      const p=x.Benutzer||{};
      if(Number(p.Id||0)===Number(current.Id||-1))return true;
      const mail=dpAuthNorm(p.EMail),login=dpAuthNorm(p.Name);
      return !!((mail&&currentEmail&&mail===currentEmail)||(login&&currentLogin&&login===currentLogin));
    });

    if(!row){dpAuthBlock('Nicht für DispoPilot freigeschaltet','Dein Microsoft-Konto ist nicht in der DispoPilot-Benutzerliste eingetragen.');return}
    if(row.Aktiv!==true){dpAuthBlock('DispoPilot-Zugang deaktiviert','Dieses Benutzerkonto ist in DispoPilot derzeit nicht aktiv.');return}

    const role=String(row.Rolle||'').trim();
    if(!['Admin','Disposition','Fahrer'].includes(role)){dpAuthBlock('Keine gültige Rolle','Für dieses Benutzerkonto ist keine gültige DispoPilot-Rolle hinterlegt.');return}

    const driverName=String(row.Fahrername||'').trim();
    if(role==='Fahrer'&&!driverName){dpAuthBlock('Fahrer nicht zugeordnet','Für dieses Benutzerkonto fehlt der Fahrername in DP_Users.');return}

    dpAuthInfo={role,driverName,displayName:String(current.Title||row.Title||current.Email||'Benutzer'),email:String(current.Email||''),loginName:String(current.LoginName||''),userId:Number(current.Id||0)};
    window.DP_AUTH=dpAuthInfo;

    if(!state.user||typeof state.user!=='object')state.user={name:''};
    if(role==='Fahrer'){
      state.user.name=driverName;
      if(Array.isArray(state.drivers)&&!state.drivers.includes(driverName))state.drivers.push(driverName);
      selectedDriver=driverName;
      orderFilter='mine';
    }else{
      state.user.name='Disposition';
      selectedDriver='all';
      if(role==='Admin')orderFilter='all';
    }

    dpAuthInstallRuntimeHooks();
    initUI();renderAll();dpAuthDecorateTop();dpAuthDecorateSettings();dpAuthDecorateOrders();
    document.body.classList.remove('dp-auth-pending');
  }catch(e){
    console.error('DispoPilot Benutzerprüfung fehlgeschlagen',e);
    dpAuthBlock('Benutzerprüfung fehlgeschlagen','DP_Users oder die Microsoft-365-Anmeldung konnte nicht gelesen werden. Bitte den DispoPilot-Administrator informieren.');
  }
}

document.addEventListener('DOMContentLoaded',()=>{dpAuthBoot()});
/* DISPOPILOT_DP_USERS_AUTH_V1_END */
'''

AUTH_CSS = r'''
/* DISPOPILOT_DP_USERS_AUTH_V1 */
body.dp-auth-pending .app{visibility:hidden}
#dpAuthBlock{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;background:#f3f5f7;font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1f2937}
.dp-auth-card{width:min(520px,100%);background:#fff;border:1px solid #d6dbe1;border-radius:18px;padding:26px;box-shadow:0 2px 10px #0001;text-align:center}
.dp-auth-card h2{margin:12px 0 8px}.dp-auth-card p{line-height:1.45}
.dp-auth-logo{width:58px;height:58px;margin:auto;border-radius:14px;background:#14532d;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:22px}
.dp-auth-small{font-size:12px;color:#6b7280}
'''

def patch_html(text: str) -> str:
    if MARKER in text:
        raise SystemExit('Die Quelldatei enthält den DP_Users-Patch bereits.')
    if '<body' not in text.lower() or '</script>' not in text.lower():
        raise SystemExit('HTML-Grundstruktur nicht gefunden.')

    def body_repl(m):
        attrs=m.group(1)
        if re.search(r'class\s*=',attrs,re.I):
            tag='<body'+attrs+'>'
            return re.sub(r'class\s*=\s*(["\'])(.*?)\1',lambda c:'class='+c.group(1)+c.group(2)+' dp-auth-pending'+c.group(1),tag,count=1,flags=re.I)
        return '<body'+attrs+' class="dp-auth-pending">'

    text=re.sub(r'<body(\s*[^>]*)>',body_repl,text,count=1,flags=re.I)
    idx_style=text.lower().rfind('</style>')
    if idx_style>=0:
        text=text[:idx_style]+'\n'+AUTH_CSS+'\n'+text[idx_style:]
    else:
        idx_head=text.lower().rfind('</head>')
        if idx_head<0: raise SystemExit('Kein </head>-Tag gefunden.')
        text=text[:idx_head]+'<style>'+AUTH_CSS+'</style>\n'+text[idx_head:]

    idx_script=text.lower().rfind('</script>')
    text=text[:idx_script]+'\n'+AUTH_JS+'\n'+text[idx_script:]
    return text

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--input',required=True)
    ap.add_argument('--output',required=True)
    args=ap.parse_args()
    src=Path(args.input);dst=Path(args.output)
    out=patch_html(src.read_text(encoding='utf-8'))
    dst.write_text(out,encoding='utf-8')
    print('DP_Users Benutzerprüfung eingefügt:',dst)
    print('Marker:',MARKER)

if __name__=='__main__':
    main()
