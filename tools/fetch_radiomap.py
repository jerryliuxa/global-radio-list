# -*- coding: utf-8 -*-
"""下载 radiomap.eu 的 54 个城市 FM 页面，供 build_data.py --radiomap 使用。

用法:
  python tools/fetch_radiomap.py [输出目录]        # 默认 ../_fetch/radiomap
  需要代理时: set HTTPS_PROXY=http://127.0.0.1:7897  (或 export)
"""
import os
import sys
import time
import urllib.request

CITIES = [
    'al/tirana', 'at/wien', 'ba/sarajevo', 'be/bruxelles', 'ch/geneve', 'ch/zuerich',
    'ch/bern', 'cy/lefkosia', 'cz/praha', 'de/berlin', 'de/hamburg', 'de/koeln',
    'de/muenchen', 'dk/kobenhavn', 'ee/tallinn', 'es/barcelona', 'es/madrid',
    'fi/helsinki', 'fr/marseille', 'fr/paris', 'gr/athina', 'hr/zagreb', 'hu/budapest',
    'ie/dublin', 'is/reykjavik', 'it/milano', 'it/roma', 'lt/vilnius', 'lu/letzebuerg',
    'lv/riga', 'mc/monaco', 'md/chisinau', 'mk/skopje', 'mt/valletta', 'nl/amsterdam',
    'no/oslo', 'pl/warszawa', 'pt/lisboa', 'ro/bucuresti', 'rs/beograd', 'se/stockholm',
    'si/ljubljana', 'sk/bratislava', 'tr/ankara', 'tr/istanbul', 'ua/kyiv', 'uk/aberdeen',
    'uk/belfast', 'uk/birmingham', 'uk/bristol', 'uk/cardiff', 'uk/edinburgh',
    'uk/london', 'uk/manchester',
]
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}


def main():
    outdir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), '..', '_fetch', 'radiomap')
    os.makedirs(outdir, exist_ok=True)
    ok = fail = 0
    for slug in CITIES:
        name = slug.replace('/', '_') + '.htm'
        dest = os.path.join(outdir, name)
        if os.path.isfile(dest) and os.path.getsize(dest) > 5000:
            ok += 1
            continue
        url = 'https://radiomap.eu/%s.htm' % slug
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=20) as resp:
                body = resp.read()
            with open(dest, 'wb') as fp:
                fp.write(body)
            ok += 1
            print('ok  %s (%d B)' % (slug, len(body)))
        except Exception as e:  # noqa: BLE001 — 抓取脚本，逐页容错
            fail += 1
            print('FAIL %s: %s' % (slug, e))
        time.sleep(0.4)  # 礼貌抓取
    print('done: %d ok, %d fail -> %s' % (ok, fail, os.path.abspath(outdir)))


if __name__ == '__main__':
    main()
