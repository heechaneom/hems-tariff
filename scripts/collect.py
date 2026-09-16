# -*- coding: utf-8 -*-
"""1차 원문 사이트를 받아 snapshots/YYYY-MM-DD/ 에 텍스트로 저장한다.
GitHub Actions(인터넷 제한 없음)가 매일 08:30 KST에 실행 → Claude 루틴은 이 스냅샷을 읽는다.
실행: python scripts/collect.py   (requests 필요)
"""
import datetime as dt, html, json, os, re, sys, time
import requests

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KST = dt.timezone(dt.timedelta(hours=9))
TODAY = dt.datetime.now(KST).strftime('%Y-%m-%d')
OUT = os.path.join(ROOT, 'snapshots', TODAY)
os.makedirs(OUT, exist_ok=True)
UA = {'User-Agent': 'Mozilla/5.0 (hems-tariff collector; +https://github.com/heechaneom/hems-tariff)'}

# (파일명, 설명, URL, 종류)  종류: html → 태그 제거 텍스트, rss → 그대로
SOURCES = [
    ('kepco_notice', '한전 공지사항 목록 (연료비조정단가·약관개정·캐시백 공지)', 'https://www.kepco.co.kr/home/media/newsroom/notice/boardList.do?boardMngNo=14', 'html'),
    ('kepco_rule_revision', '한전 약관 개정 사전예고 게시판', 'https://www.kepco.co.kr/home/disclosure/regulations/revision/revision/boardList.do?boardMngNo=69', 'html'),
    ('kepco_internalrule', '한전 기본공급약관·시행세칙 게시판', 'https://www.kepco.co.kr/home/disclosure/regulations/internalrule/boardList.do', 'html'),
    ('koreakr_search_electric', '정책브리핑 검색: 전기요금', 'https://www.korea.kr/news/policyNewsList.do?srchWord=%EC%A0%84%EA%B8%B0%EC%9A%94%EA%B8%88', 'html'),
    ('koreakr_search_dr', '정책브리핑 검색: 수요반응', 'https://www.korea.kr/news/policyNewsList.do?srchWord=%EC%88%98%EC%9A%94%EB%B0%98%EC%9D%91', 'html'),
    ('koreakr_search_cashback', '정책브리핑 검색: 에너지캐시백', 'https://www.korea.kr/news/policyNewsList.do?srchWord=%EC%97%90%EB%84%88%EC%A7%80%EC%BA%90%EC%8B%9C%EB%B0%B1', 'html'),
    ('kpx_predisclosure', '전력거래소 사전정보공표 (수요자원거래시장 현황)', 'https://kpx.or.kr/board.es?mid=a10102000000&bid=0088', 'html'),
    ('kpx_notice', '전력거래소 공지사항', 'https://www.kpx.or.kr/board.es?mid=a10101000000&bid=0026', 'html'),
    ('kpx_rule', '전력거래소 전력시장운영규칙 게시판', 'https://www.kpx.or.kr/board.es?mid=a10403010000&bid=0032', 'html'),
]


def to_text(raw: bytes, resp) -> str:
    enc = resp.encoding or 'utf-8'
    for e in (enc, 'utf-8', 'euc-kr', 'cp949'):
        try:
            t = raw.decode(e)
            if '�' not in t[:5000]:
                break
        except Exception:
            continue
    t = re.sub(r'<script.*?</script>|<style.*?</style>|<!--.*?-->', ' ', t, flags=re.S | re.I)
    t = re.sub(r'<br\s*/?>|</p>|</li>|</tr>|</h\d>|</div>', '\n', t, flags=re.I)
    t = re.sub(r'<[^>]+>', ' ', t)
    t = html.unescape(t).replace('\xa0', ' ')
    t = re.sub(r'[ \t\r]+', ' ', t)
    t = re.sub(r'\n\s*\n+', '\n', t)
    return t.strip()


def main():
    index = {'date': TODAY, 'collected_at': dt.datetime.now(KST).strftime('%Y-%m-%d %H:%M KST'), 'items': []}
    for name, desc, url, kind in SOURCES:
        rec = {'name': name, 'desc': desc, 'url': url, 'kind': kind}
        try:
            r = requests.get(url, headers=UA, timeout=30)
            rec['status'] = r.status_code
            body = r.content
            text = body.decode(r.encoding or 'utf-8', errors='replace') if kind == 'rss' else to_text(body, r)
            text = text[:300_000]
            path = os.path.join(OUT, f'{name}.txt')
            open(path, 'w', encoding='utf-8').write(f'# {desc}\n# {url}\n# {index["collected_at"]} status={r.status_code}\n\n{text}')
            rec['bytes'] = len(text)
            m = re.search(r'<title>(.*?)</title>', body.decode(r.encoding or 'utf-8', errors='replace'), re.S | re.I)
            rec['title'] = html.unescape(m.group(1)).strip()[:120] if m else ''
            print(f'ok   {name} {r.status_code} {len(text)}B')
        except Exception as e:  # 한 소스 실패가 전체를 죽이지 않게
            rec['status'] = 'error'; rec['error'] = str(e)[:200]
            print(f'FAIL {name} {e}')
        index['items'].append(rec)
        time.sleep(1)
    json.dump(index, open(os.path.join(OUT, 'index.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    ok = sum(1 for i in index['items'] if i.get('status') == 200)
    print(f'collected {ok}/{len(SOURCES)} -> {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
