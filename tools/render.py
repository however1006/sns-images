"""SNS画像の生成。使い方: python3 render.py spec.json 出力.png
spec の type: line(トーク画面風) / memo(メモ帳風) / check(あてはまりチェック)
共通: account(kango/hoiku/daini/main)。必要なフォント: Noto Sans CJK JP / Noto Color Emoji
line : title(相手の名前), fiction(true なら「※この話は創作です」), messages=[{from:"them"|"me"|"sys", text, time}], footer(問いかけ)
memo : time(例 "7:42"), title, items=[...](**で囲むとマーカー), footer(省略可)
check: lead, title, items=[...], result(**で囲むとマーカー)
"""
import json, sys, html, re
from playwright.sync_api import sync_playwright

NAMES = {'kango': '夜勤やめたい看護師の行き先メモ', 'hoiku': '子どもは好き。でも、この園はムリ',
         'daini': '新卒1年目、もう無理かも', 'main': 'フリーランス準備室'}
NOTE = {'kango': '運営は看護師ではありません', 'hoiku': '運営は保育園に子どもを預ける親です'}
COLOR = {'kango': '#4682b4', 'hoiku': '#ef7f8c', 'daini': '#5b5b82', 'main': '#3a8f6b'}
FONT = '"Noto Sans CJK JP","Noto Color Emoji",sans-serif'

def esc(s):
    s = html.escape(s).replace('\n', '<br>')
    return re.sub(r'\*\*(.+?)\*\*', r'<span class="hl">\1</span>', s)

def line(sp):
    rows = []
    for m in sp['messages']:
        if m['from'] == 'sys':
            rows.append(f'<div class="sys"><span>{esc(m["text"])}</span></div>')
        elif m['from'] == 'me':
            rows.append(f'<div class="row me"><div class="t">既読<br>{esc(m.get("time",""))}</div><div class="b">{esc(m["text"])}</div></div>')
        else:
            rows.append(f'<div class="row"><div class="av">{sp.get("avatar","👤")}</div><div class="b">{esc(m["text"])}</div><div class="t">{esc(m.get("time",""))}</div></div>')
    badge = '<div class="badge">※この話は創作です</div>' if sp.get('fiction') else ''
    return f'''<style>body{{background:#8ca4c4}}.hd{{height:150px;background:#2b3a55;color:#fff;display:flex;align-items:center;padding:0 50px;font-size:46px;font-weight:700}}
.badge{{position:absolute;top:40px;right:40px;background:#fff;color:#2b3a55;font-size:28px;font-weight:700;padding:8px 20px;border-radius:30px}}
.chat{{padding:40px}}.row{{display:flex;align-items:flex-end;margin-bottom:30px}}.row.me{{justify-content:flex-end}}
.av{{width:90px;height:90px;border-radius:50%;background:#dfe6ef;display:flex;align-items:center;justify-content:center;font-size:50px;margin-right:18px;flex:none}}
.b{{max-width:720px;padding:24px 32px;border-radius:40px;font-size:42px;line-height:1.45;background:#fff}}.me .b{{background:#8de055}}
.t{{font-size:26px;color:#f3f6fa;margin:0 14px}}.sys{{text-align:center;margin:6px 0 30px}}.sys span{{background:rgba(0,0,0,.25);color:#fff;font-size:30px;padding:8px 26px;border-radius:30px}}
.foot{{position:absolute;bottom:0;left:0;right:0;background:#fff;padding:28px 50px;font-size:32px;color:#555;line-height:1.5}}</style>
<div class="hd">{esc(sp.get("title","上司"))}</div>{badge}<div class="chat">{"".join(rows)}</div>
<div class="foot">{esc(sp.get("footer",""))}<br><b>{NAMES[sp["account"]]}</b></div>'''

def memo(sp):
    items = ''.join(f'<div><span class="d">・</span><span>{esc(i)}</span></div>' for i in sp['items'])
    foot = NAMES[sp['account']] + ('|' + NOTE[sp['account']] if sp['account'] in NOTE else '')
    return f'''<style>body{{background:#fbf8f1;color:#1c1c1e}}.sb{{height:90px;display:flex;justify-content:space-between;align-items:center;padding:0 60px;font-size:34px;font-weight:700}}
.nav{{height:80px;display:flex;align-items:center;padding:0 40px;color:#e0a800;font-size:40px}}.t{{font-size:60px;font-weight:900;padding:30px 70px 0;line-height:1.35}}
.list{{padding:50px 70px 0;font-size:48px;line-height:1.5}}.list div{{margin-bottom:34px;display:flex}}.d{{color:#e0a800;margin-right:18px}}
.hl{{background:linear-gradient(transparent 55%,#ffe27a 55%)}}.ex{{padding:10px 70px 0;font-size:40px;font-weight:700;line-height:1.5}}
.foot{{position:absolute;bottom:60px;left:70px;right:70px;font-size:30px;color:#8e8e93;border-top:2px solid #e5e0d5;padding-top:24px}}</style>
<div class="sb"><span>{esc(sp.get("time","7:42"))}</span><span>📶 🔋</span></div><div class="nav">‹ メモ</div>
<div class="t">{esc(sp["title"])}</div><div class="list">{items}</div><div class="ex">{esc(sp.get("footer",""))}</div><div class="foot">{foot}</div>'''

def check(sp):
    c = COLOR[sp['account']]
    items = ''.join(f'<div class="it"><div class="bx"></div>{esc(i)}</div>' for i in sp['items'])
    return f'''<style>body{{background:#fff6f3;color:#3d2a2f}}.top{{background:{c};color:#fff;padding:56px 70px 48px;border-radius:0 0 60px 60px}}
.s{{font-size:38px;font-weight:700;opacity:.9}}.tt{{font-size:80px;font-weight:900;line-height:1.25;margin-top:14px}}.list{{padding:40px 70px 0}}
.it{{display:flex;align-items:center;background:#fff;border-radius:28px;padding:22px 30px;margin-bottom:18px;font-size:42px;font-weight:700;box-shadow:0 6px 0 #eadad6}}
.bx{{width:56px;height:56px;border:6px solid {c};border-radius:14px;margin-right:26px;flex:none}}.res{{margin:30px 70px 0;font-size:44px;font-weight:900;line-height:1.5}}
.hl{{background:linear-gradient(transparent 55%,#ffd36b 55%)}}.foot{{position:absolute;bottom:36px;left:70px;font-size:30px;color:#8a7378}}</style>
<div class="top"><div class="s">{esc(sp.get("lead",""))}</div><div class="tt">{esc(sp["title"])}</div></div><div class="list">{items}</div>
<div class="res">{esc(sp.get("result",""))}</div><div class="foot">{NAMES[sp["account"]]}</div>'''

def render(sp, out):
    body = {'line': line, 'memo': memo, 'check': check}[sp['type']](sp)
    doc = f'<html><head><meta charset="utf-8"><style>*{{margin:0;padding:0;box-sizing:border-box}}body{{width:1080px;height:1350px;font-family:{FONT};position:relative;overflow:hidden}}</style></head><body>{body}</body></html>'
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_page(viewport={'width': 1080, 'height': 1350})
        pg.set_content(doc); pg.wait_for_timeout(300)
        # はみ出し検査
        over = pg.evaluate('document.body.scrollHeight') > 1350
        pg.screenshot(path=out); b.close()
    if over: print('WARN: はみ出し', out)
    return not over

if __name__ == '__main__':
    render(json.load(open(sys.argv[1])), sys.argv[2])
