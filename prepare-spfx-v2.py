#!/usr/bin/env python3
from __future__ import annotations
import argparse, json, re, uuid
from pathlib import Path

SOLUTION_ID = "9b7ca846-1b54-4d77-b31d-e8f8fe0fd7d0"
COMPONENT_ID = "02f85dd7-6c1b-4bc2-97a2-84b0c3c92617"


def extract_app(app_path: Path):
    text = app_path.read_text(encoding='utf-8')
    style_parts = re.findall(r'<style[^>]*>(.*?)</style>', text, flags=re.I|re.S)
    scripts = re.findall(r'<script[^>]*>(.*?)</script>', text, flags=re.I|re.S)
    if not scripts:
        raise SystemExit('Kein App-Script gefunden')
    main_script = max(scripts, key=len)
    # DOMContentLoaded fires before we evaluate the app inside the iframe; start explicitly later.
    main_script = re.sub(
        r"\s*document\.addEventListener\('DOMContentLoaded',\s*async\(\)=>\{await loadState\(\);await dp13Init\(\);\}\);\s*$",
        "\n",
        main_script,
        flags=re.S,
    )
    # The interactive state variables must be accessible to the delegated handler bridge.
    main_script = re.sub(r'^let\s+state=', 'var state=', main_script, count=1, flags=re.M)

    body_m = re.search(r'<body[^>]*>(.*?)</body>', text, flags=re.I|re.S)
    if not body_m:
        raise SystemExit('Kein BODY gefunden')
    body = body_m.group(1)
    body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.I|re.S)

    # Inline handlers are blocked by modern SharePoint CSP. Convert them to inert data attributes.
    def repl(m):
        ev = m.group(1).lower()[2:]
        quote = m.group(2)
        code = m.group(3)
        return f' data-dp-{ev}={quote}{code}{quote}'
    body = re.sub(r'\s+(on[a-zA-Z]+)\s*=\s*(["\'])(.*?)\2', repl, body, flags=re.S)

    css = '\n'.join(style_parts)
    static_html = f'''<!doctype html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><style>{css}</style></head><body>{body}</body></html>'''
    return static_html, main_script


def js_string(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


def make_ts(static_html: str, app_script: str) -> str:
    return f'''import {{ BaseClientSideWebPart }} from '@microsoft/sp-webpart-base';

export interface IDispoPilotWebPartProps {{
  readonly placeholder?: string;
}}

type DpWindow = Window & typeof globalThis & {{
  DP_SHAREPOINT?: {{ webUrl: string; userDisplayName: string; userEmail: string }};
  __dpThis?: Element;
  __dpEvent?: Event;
}};

const STATIC_HTML: string = {js_string(static_html)};
const APP_SCRIPT: string = {js_string(app_script)};
const DP_EVENTS: readonly string[] = ['click', 'change', 'input', 'keydown'];

export default class DispoPilotWebPart extends BaseClientSideWebPart<IDispoPilotWebPartProps> {{
  public render(): void {{
    this.domElement.innerHTML = '<div class="dp-spfx-host" style="width:100%;min-height:760px"></div>';
    const host = this.domElement.querySelector('.dp-spfx-host') as HTMLDivElement;
    const frame = document.createElement('iframe');
    frame.title = 'DispoPilot Online V1.3';
    frame.style.cssText = 'width:100%;height:calc(100vh - 150px);min-height:760px;border:0;border-radius:10px;background:#f3f5f7;display:block';
    host.appendChild(frame);
    frame.addEventListener('load', () => this.bootFrame(frame), {{ once: true }});
    frame.srcdoc = STATIC_HTML;
  }}

  private bootFrame(frame: HTMLIFrameElement): void {{
    try {{
      const win = frame.contentWindow as DpWindow | null;
      const doc = frame.contentDocument;
      if (!win || !doc) {{ throw new Error('Iframe-Kontext konnte nicht geöffnet werden.'); }}

      win.DP_SHAREPOINT = {{
        webUrl: this.context.pageContext.web.absoluteUrl,
        userDisplayName: this.context.pageContext.user.displayName || '',
        userEmail: this.context.pageContext.user.email || ''
      }};

      this.installHandlerBridge(win, doc);
      win.eval(APP_SCRIPT + '\\n//# sourceURL=dispopilot-v13-runtime.js');
      win.eval('(async function(){{ await loadState(); if (typeof dp13Init === "function") {{ await dp13Init(); }} }})().catch(function(e){{ console.error(e); alert("DispoPilot Startfehler: " + (e && e.message ? e.message : e)); }});');
    }} catch (e) {{
      const msg = e instanceof Error ? e.message : String(e);
      this.domElement.innerHTML = '<div style="padding:16px;border:1px solid #fecaca;background:#fef2f2;color:#991b1b;border-radius:10px"><strong>DispoPilot konnte nicht gestartet werden.</strong><br>' + this.escapeHtml(msg) + '</div>';
      console.error('DispoPilot SPFx boot error', e);
    }}
  }}

  private installHandlerBridge(win: DpWindow, doc: Document): void {{
    const scan = (root: ParentNode): void => {{
      const elements: Element[] = [];
      if (root instanceof win.Element) {{ elements.push(root as Element); }}
      root.querySelectorAll?.('*').forEach((el: Element) => elements.push(el));
      elements.forEach((el: Element) => {{
        DP_EVENTS.forEach((ev: string) => {{
          const attr = 'on' + ev;
          if (el.hasAttribute(attr)) {{
            const code = el.getAttribute(attr) || '';
            el.setAttribute('data-dp-' + ev, code);
            el.removeAttribute(attr);
          }}
        }});
      }});
    }};

    scan(doc);
    const Observer = win.MutationObserver;
    if (Observer) {{
      const observer = new Observer((mutations: MutationRecord[]) => {{
        mutations.forEach((m: MutationRecord) => {{
          m.addedNodes.forEach((n: Node) => {{
            if (n.nodeType === 1) {{ scan(n as Element); }}
          }});
        }});
      }});
      observer.observe(doc.documentElement, {{ childList: true, subtree: true }});
    }}

    DP_EVENTS.forEach((ev: string) => {{
      doc.addEventListener(ev, (event: Event) => {{
        let el = event.target as Element | null;
        while (el && el !== doc.documentElement) {{
          const code = el.getAttribute?.('data-dp-' + ev);
          if (code) {{
            win.__dpThis = el;
            win.__dpEvent = event;
            try {{
              win.eval('(function(event){{' + code + '\\n}}).call(window.__dpThis, window.__dpEvent);');
            }} catch (err) {{
              console.error('DispoPilot event handler error', code, err);
            }}
            return;
          }}
          el = el.parentElement;
        }}
      }}, false);
    }});
  }}

  private escapeHtml(value: string): string {{
    return value.replace(/[&<>"']/g, (c: string) => {{
      switch (c) {{
        case '&': return '&amp;';
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '"': return '&quot;';
        case "'": return '&#39;';
        default: return c;
      }}
    }});
  }}
}}
'''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--app', required=True)
    ap.add_argument('--project', required=True)
    args = ap.parse_args()
    app = Path(args.app).resolve()
    project = Path(args.project).resolve()
    static_html, app_script = extract_app(app)

    webparts = list(project.glob('src/webparts/**/*WebPart.ts'))
    if not webparts:
        raise SystemExit('WebPart.ts im SPFx-Projekt nicht gefunden')
    webpart = webparts[0]
    webpart.write_text(make_ts(static_html, app_script), encoding='utf-8')

    manifests = list(project.glob('src/webparts/**/*.manifest.json'))
    if not manifests:
        raise SystemExit('WebPart manifest nicht gefunden')
    manifest = manifests[0]
    data = json.loads(manifest.read_text(encoding='utf-8'))
    data['id'] = COMPONENT_ID
    data['alias'] = 'DispoPilotOnlineWebPart'
    data['requiresCustomScript'] = False
    entries = data.get('preconfiguredEntries') or [{}]
    entries[0]['title'] = {'default': 'DispoPilot Online'}
    entries[0]['description'] = {'default': 'DispoPilot Online Pilot V1.3'}
    entries[0]['officeFabricIconFontName'] = 'Truck'
    data['preconfiguredEntries'] = entries
    manifest.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')

    ps = project / 'config' / 'package-solution.json'
    pdata = json.loads(ps.read_text(encoding='utf-8'))
    sol = pdata.setdefault('solution', {})
    sol['name'] = 'dispopilot-online-v13-csp-client-side-solution'
    sol['id'] = SOLUTION_ID
    sol['version'] = '1.3.1.0'
    sol['includeClientSideAssets'] = True
    sol['skipFeatureDeployment'] = True
    sol['isDomainIsolated'] = False
    paths = pdata.setdefault('paths', {})
    paths['zippedPackage'] = 'solution/DispoPilot_Online_V1.3_CSP.sppkg'
    ps.write_text(json.dumps(pdata, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')

    print('Prepared:', webpart)
    print('Manifest:', manifest)
    print('Package:', ps)

if __name__ == '__main__':
    main()
