#!/usr/bin/env python3
from __future__ import annotations
import argparse, json, re
from pathlib import Path

SOLUTION_ID = "9b7ca846-1b54-4d77-b31d-e8f8fe0fd7d0"
COMPONENT_ID = "02f85dd7-6c1b-4bc2-97a2-84b0c3c92617"

FINAL_MARKER = "DispoPilot SPFx FINAL JSONC v1"

def load_jsonc(path: Path):
    """Liest JSON/JSONC: erlaubt //- und /* */-Kommentare sowie trailing commas."""
    src = path.read_text(encoding="utf-8")
    out = []
    i = 0
    quote = None
    escaped = False
    while i < len(src):
        ch = src[i]
        nxt = src[i + 1] if i + 1 < len(src) else ""
        if quote:
            out.append(ch)
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == quote:
                quote = None
            i += 1
            continue
        if ch in ('"', "'"):
            quote = ch
            out.append(ch)
            i += 1
            continue
        if ch == "/" and nxt == "/":
            i += 2
            while i < len(src) and src[i] not in "\r\n":
                i += 1
            continue
        if ch == "/" and nxt == "*":
            i += 2
            while i + 1 < len(src) and not (src[i] == "*" and src[i + 1] == "/"):
                i += 1
            i += 2
            continue
        out.append(ch)
        i += 1

    cleaned = "".join(out)
    # Nachlaufende Kommas vor } oder ] entfernen, ohne Strings anzufassen.
    prev = None
    while prev != cleaned:
        prev = cleaned
        cleaned = re.sub(r",\s*([}\]])", r"\1", cleaned)

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        preview = "\n".join(cleaned.splitlines()[:30])
        raise SystemExit(
            f"JSON/JSONC konnte nicht gelesen werden: {path}\n"
            f"{e}\n--- Anfang der bereinigten Datei ---\n{preview}"
        )



def extract_app(app_path: Path):
    text = app_path.read_text(encoding="utf-8")
    styles = re.findall(r"<style[^>]*>(.*?)</style>", text, flags=re.I | re.S)
    scripts = re.findall(r"<script[^>]*>(.*?)</script>", text, flags=re.I | re.S)
    if not scripts:
        raise SystemExit("Kein App-Script gefunden")
    main_script = max(scripts, key=len)
    main_script = re.sub(
        r"\s*document\.addEventListener\('DOMContentLoaded',\s*async\(\)=>\{await loadState\(\);await dp13Init\(\);\}\);\s*$",
        "\n",
        main_script,
        flags=re.S,
    )
    body_m = re.search(r"<body[^>]*>(.*?)</body>", text, flags=re.I | re.S)
    if not body_m:
        raise SystemExit("Kein BODY gefunden")
    body = re.sub(r"<script[^>]*>.*?</script>", "", body_m.group(1), flags=re.I | re.S)

    def repl(m):
        ev = m.group(1).lower()[2:]
        return f" data-dp-{ev}={m.group(2)}{m.group(3)}{m.group(2)}"

    body = re.sub(
        r"\s+(on(?:click|change|input|keydown))\s*=\s*([\"'])(.*?)\2",
        repl,
        body,
        flags=re.I | re.S,
    )
    static_html = (
        '<!doctype html><html lang="de"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
        f"<style>{' '.join(styles)}</style></head><body>{body}</body></html>"
    )

    handlers = []
    for m in re.finditer(
        r"\bon(?:click|change|input|keydown)\s*=\s*(?:\"([^\"]*)\"|'([^']*)')",
        text,
        flags=re.I | re.S,
    ):
        handlers.append(m.group(1) if m.group(1) is not None else m.group(2))
    names = set()
    for h in handlers:
        for n in re.findall(r"(?<![.\w$])([A-Za-z_$][\w$]*)\s*\(", h):
            if n != "if":
                names.add(n)
    names.add("selectPlanningDate")
    declared = set(re.findall(r"\b(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(", main_script))
    declared |= set(re.findall(r"\b([A-Za-z_$][\w$]*)\s*=\s*(?:async\s+)?function\s*\(", main_script))
    return static_html, main_script, sorted(names & declared)


def js_string(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


def runtime_js(static_html: str, app_script: str, fn_names: list[str]) -> str:
    fn_map = ",\n    ".join(f"{n}: (...args) => {n}(...args)" for n in fn_names)
    template = r"""export const staticHtml = __STATIC__;

export async function bootDispoPilot(win, doc, config) {
  const window = win;
  const document = doc;
  window.DP_SHAREPOINT = config;
  const navigator = window.navigator;
  const location = window.location;
  const localStorage = window.localStorage;
  const fetch = window.fetch.bind(window);
  const alert = window.alert.bind(window);
  const confirm = window.confirm.bind(window);
  const FileReader = window.FileReader;
  const Blob = window.Blob;
  const URL = window.URL;
  const TextDecoder = window.TextDecoder;
  const Image = window.Image;
  const setTimeout = window.setTimeout.bind(window);
  const clearTimeout = window.clearTimeout.bind(window);

__APP__

  await loadState();
  await dp13Init();

  const DP_FUN = {
    __FNMAP__
  };
  const DP_GET = { currentOrderId: () => currentOrderId };
  const DP_SET = { selectedDriver: (v) => { selectedDriver = v; } };

  function splitTop(text, delimiter) {
    const out=[]; let cur=''; let quote=null; let escaped=false; let depth=0;
    for(let i=0;i<text.length;i++) {
      const ch=text[i];
      if(escaped){cur+=ch; escaped=false; continue;}
      if(quote){cur+=ch; if(ch==='\\')escaped=true; else if(ch===quote)quote=null; continue;}
      if(ch==='"' || ch==="'"){quote=ch; cur+=ch; continue;}
      if(ch==='(' || ch==='[' || ch==='{'){depth++; cur+=ch; continue;}
      if(ch===')' || ch===']' || ch==='}'){depth--; cur+=ch; continue;}
      if(ch===delimiter && depth===0){if(cur.trim())out.push(cur.trim()); cur=''; continue;}
      cur+=ch;
    }
    if(cur.trim())out.push(cur.trim());
    return out;
  }

  function parseQuoted(token) {
    const inner=token.slice(1,-1);
    return inner.replace(/\\([\\'"nrt])/g, (_m,c) => c==='n'?'\n':c==='r'?'\r':c==='t'?'\t':c);
  }

  function parseArg(token, el, event) {
    const t=token.trim();
    if(!t)return undefined;
    if((t[0]==='"' && t.at(-1)==='"') || (t[0]==="'" && t.at(-1)==="'"))return parseQuoted(t);
    if(/^-?\d+(?:\.\d+)?$/.test(t))return Number(t);
    if(t==='true')return true; if(t==='false')return false; if(t==='null')return null;
    if(t==='this')return el;
    if(t==='this.value')return el && 'value' in el ? el.value : undefined;
    if(t==='event')return event;
    if(DP_GET[t])return DP_GET[t]();
    throw new Error('Nicht unterstütztes Handler-Argument: '+t);
  }

  async function runStatement(stmt, el, event) {
    const s=stmt.trim(); if(!s)return;
    let m=s.match(/^if\(event\.key===(['"])(.*?)\1\)(.+)$/);
    if(m){if(event && event.key===m[2])await runHandler(m[3],el,event); return;}
    if(s==='event.stopPropagation()'){event && event.stopPropagation(); return;}
    if(s==='event.preventDefault()'){event && event.preventDefault(); return;}
    if(s==='window.print()'){window.print(); return;}
    m=s.match(/^([A-Za-z_$][\w$]*)\s*=\s*(.+)$/);
    if(m && DP_SET[m[1]]){DP_SET[m[1]](parseArg(m[2],el,event)); return;}
    m=s.match(/^([A-Za-z_$][\w$]*)\((.*)\)$/s);
    if(m){
      const fn=DP_FUN[m[1]]; if(!fn)throw new Error('Nicht freigegebener Handler: '+m[1]);
      const args=m[2].trim()?splitTop(m[2],',').map(x=>parseArg(x,el,event)):[];
      await fn(...args); return;
    }
    throw new Error('Nicht unterstützter Handler: '+s);
  }

  async function runHandler(code, el, event){
    for(const stmt of splitTop(String(code||''),';'))await runStatement(stmt,el,event);
  }

  const DP_EVENTS=['click','change','input','keydown'];
  function scan(root){
    const els=[];
    if(root && root.nodeType===1)els.push(root);
    if(root && root.querySelectorAll)root.querySelectorAll('*').forEach(el=>els.push(el));
    els.forEach(el=>DP_EVENTS.forEach(ev=>{
      const attr='on'+ev;
      if(el.hasAttribute && el.hasAttribute(attr)){
        const code=el.getAttribute(attr)||'';
        el.setAttribute('data-dp-'+ev,code);
        el.removeAttribute(attr);
      }
    }));
  }
  scan(document.documentElement);
  const observer=new window.MutationObserver(ms=>ms.forEach(m=>m.addedNodes.forEach(n=>{if(n.nodeType===1)scan(n)})));
  observer.observe(document.documentElement,{childList:true,subtree:true});
  DP_EVENTS.forEach(ev=>document.addEventListener(ev,event=>{
    let el=event.target;
    while(el && el!==document.documentElement){
      const code=el.getAttribute && el.getAttribute('data-dp-'+ev);
      if(code){
        runHandler(code,el,event).catch(err=>{
          console.error('DispoPilot handler',code,err);
          alert('Bedienfehler in DispoPilot: '+(err&&err.message?err.message:err));
        });
        return;
      }
      el=el.parentElement;
    }
  },false));
}
"""
    return template.replace("__STATIC__", js_string(static_html)).replace("__APP__", app_script).replace("__FNMAP__", fn_map)


def webpart_ts() -> str:
    return """import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { staticHtml, bootDispoPilot } from './app-runtime';

export interface IDispoPilotOnlineWebPartProps {}

export default class DispoPilotOnlineWebPart extends BaseClientSideWebPart<IDispoPilotOnlineWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = '<div class=\"dp-spfx-host\" style=\"width:100%;min-height:760px\"></div>';
    const host = this.domElement.querySelector('.dp-spfx-host') as HTMLDivElement;
    const frame = document.createElement('iframe');
    frame.title = 'DispoPilot Online V1.3';
    frame.style.cssText = 'width:100%;height:calc(100vh - 150px);min-height:760px;border:0;border-radius:10px;background:#f3f5f7;display:block';
    host.appendChild(frame);
    frame.addEventListener('load', () => {
      const win = frame.contentWindow;
      const doc = frame.contentDocument;
      if (!win || !doc) { this.showError('Iframe-Kontext konnte nicht geöffnet werden.'); return; }
      void bootDispoPilot(win, doc, {
        webUrl: this.context.pageContext.web.absoluteUrl,
        userDisplayName: this.context.pageContext.user.displayName || '',
        userEmail: this.context.pageContext.user.email || ''
      }).catch((e: unknown) => {
        const msg = e instanceof Error ? e.message : String(e);
        console.error('DispoPilot boot error', e);
        this.showError(msg);
      });
    }, { once: true });
    frame.srcdoc = staticHtml;
  }

  private showError(message: string): void {
    this.domElement.innerHTML = '<div style=\"padding:16px;border:1px solid #fecaca;background:#fef2f2;color:#991b1b;border-radius:10px\"><strong>DispoPilot konnte nicht gestartet werden.</strong><br>' + this.escapeHtml(message) + '</div>';
  }

  private escapeHtml(value: string): string {
    return value.replace(/[&<>\"']/g, (c: string) => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',\"'\":'&#39;'}[c] || c));
  }
}
"""


def main():
    print(FINAL_MARKER)
    ap = argparse.ArgumentParser()
    ap.add_argument('--app', required=True)
    ap.add_argument('--project', required=True)
    args = ap.parse_args()
    app = Path(args.app).resolve(); project = Path(args.project).resolve()
    static_html, app_script, fn_names = extract_app(app)
    webparts = list(project.glob('src/webparts/**/*WebPart.ts'))
    if not webparts: raise SystemExit('WebPart.ts nicht gefunden')
    webpart = webparts[0]; folder = webpart.parent
    webpart.write_text(webpart_ts(), encoding='utf-8')
    (folder/'app-runtime.js').write_text(runtime_js(static_html, app_script, fn_names), encoding='utf-8')
    (folder/'app-runtime.d.ts').write_text(
        "export const staticHtml: string;\nexport function bootDispoPilot(win: Window, doc: Document, config: { webUrl: string; userDisplayName: string; userEmail: string }): Promise<void>;\n",
        encoding='utf-8')

    manifest = list(project.glob('src/webparts/**/*.manifest.json'))[0]
    data = load_jsonc(manifest)
    data['id'] = COMPONENT_ID; data['alias'] = 'DispoPilotOnlineWebPart'; data['requiresCustomScript'] = False
    entries = data.get('preconfiguredEntries') or [{}]
    entries[0]['title'] = {'default':'DispoPilot Online'}
    entries[0]['description'] = {'default':'DispoPilot Online Pilot V1.3'}
    entries[0]['officeFabricIconFontName'] = 'Truck'
    data['preconfiguredEntries'] = entries
    manifest.write_text(json.dumps(data, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')

    ps = project/'config'/'package-solution.json'
    pdata = load_jsonc(ps)
    sol = pdata.setdefault('solution', {})
    sol['name'] = 'dispopilot-online-v13-client-side-solution'
    sol['id'] = SOLUTION_ID
    sol['version'] = '1.3.2.0'
    sol['includeClientSideAssets'] = True
    sol['skipFeatureDeployment'] = True
    sol['isDomainIsolated'] = False
    pdata.setdefault('paths', {})['zippedPackage'] = 'solution/DispoPilot_Online_V1.3_SPFX.sppkg'
    ps.write_text(json.dumps(pdata, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print('Prepared', webpart, 'handler functions', len(fn_names))

if __name__ == '__main__':
    main()
