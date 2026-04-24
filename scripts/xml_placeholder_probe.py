"""Placeholder XML/nvSpPr 태그가 category로 사용되는지 확인.

pptx 파일의 slide XML에서 <p:ph type="..." idx="..."/> 분포 조사.
"""
import zipfile
import re
from collections import Counter

PATH = 'c:/Users/y2kbo/Apps/PPT/docs/references/_master_templates/PPT 템플릿.pptx'

ph_type_ctr = Counter()
ph_idx_ctr = Counter()
descr_ctr = Counter()
custData = 0
ext_uri_ctr = Counter()

# Check first 50 slides XML
with zipfile.ZipFile(PATH) as z:
    slide_names = sorted([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])
    print(f"Total slide xml files: {len(slide_names)}")
    for i, nm in enumerate(slide_names):
        with z.open(nm) as f:
            data = f.read().decode('utf-8', errors='replace')
        # <p:ph type="..." idx="..."/>
        for m in re.finditer(r'<p:ph\s+([^/>]*)/?>', data):
            attrs = m.group(1)
            t_m = re.search(r'type="([^"]+)"', attrs)
            i_m = re.search(r'idx="([^"]+)"', attrs)
            ph_type_ctr[t_m.group(1) if t_m else 'DEFAULT'] += 1
            ph_idx_ctr[i_m.group(1) if i_m else 'NONE'] += 1
        # descr attribute (accessibility description)
        for m in re.finditer(r'descr="([^"]+)"', data):
            descr_ctr[m.group(1)[:60]] += 1
        if '<p:custDataLst>' in data or '<p:custData ' in data:
            custData += 1
        # extensions (Section codes?)
        for m in re.finditer(r'<p:ext uri="([^"]+)"', data):
            ext_uri_ctr[m.group(1)] += 1

print(f"\nPlaceholder <p:ph type=> 분포:")
for t, c in ph_type_ctr.most_common():
    print(f"  type={t!r:15s} {c}")

print(f"\nPlaceholder <p:ph idx=> 분포 (상위 20):")
for t, c in ph_idx_ctr.most_common(20):
    print(f"  idx={t!r:10s} {c}")

print(f"\ndescr(접근성) 속성 샘플 (상위 20):")
for t, c in descr_ctr.most_common(20):
    print(f"  {t!r} {c}")

print(f"\n<p:custData> 포함 슬라이드: {custData}")

print(f"\n<p:ext uri> 종류:")
for t, c in ext_uri_ctr.most_common():
    print(f"  {t}: {c}")

# ph가 존재하는 슬라이드 개수
with zipfile.ZipFile(PATH) as z:
    slide_names = sorted([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])
    has_title_ph = 0
    has_any_ph = 0
    for nm in slide_names:
        with z.open(nm) as f:
            data = f.read().decode('utf-8', errors='replace')
        if '<p:ph' in data:
            has_any_ph += 1
        if re.search(r'<p:ph[^/>]*type="title"', data):
            has_title_ph += 1
    print(f"\n<p:ph> 있는 슬라이드: {has_any_ph}/{len(slide_names)}")
    print(f"title placeholder 있는 슬라이드: {has_title_ph}/{len(slide_names)}")
