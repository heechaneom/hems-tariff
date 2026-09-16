# -*- coding: utf-8 -*-
"""template/tariff_board.template.html + data/data.json → docs/index.html (GitHub Pages용)
실행: python scripts/build.py   (extract.py를 먼저 실행하거나 --all 옵션)
"""
import json, os, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL = os.path.join(ROOT, 'template', 'tariff_board.template.html')
DATA = os.path.join(ROOT, 'data', 'data.json')
OUT = os.path.join(ROOT, 'docs', 'index.html')


def main():
    if '--all' in sys.argv:
        subprocess.check_call([sys.executable, os.path.join(ROOT, 'scripts', 'extract.py')])
    t = open(TPL, encoding='utf-8').read()
    D = json.load(open(DATA, encoding='utf-8'))
    assert '/*DATA*/' in t, 'template에 /*DATA*/ 자리표시자가 없음'
    out = t.replace('/*DATA*/', json.dumps(D, ensure_ascii=False).replace('</script', '<\\/script'))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    open(OUT, 'w', encoding='utf-8').write('<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head><body>\n' + out + '\n</body></html>')
    print(f"build ok: {D['version']['ver']} {D['version']['date']} -> {OUT} ({os.path.getsize(OUT)//1024}KB)")


if __name__ == '__main__':
    main()
