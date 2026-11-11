#!/usr/bin/env python3
"""Read-only PPTX inventory in presentation order; review flags are not edit commands."""
import argparse
from collections import Counter
import json
from pathlib import Path
import posixpath
import re
from zipfile import ZipFile
import xml.etree.ElementTree as E

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def relationships(package, part):
    folder, filename = posixpath.split(part)
    relpart = posixpath.join(folder, '_rels', filename + '.rels')
    if relpart not in package.namelist():
        return {}
    result = {}
    for rel in E.fromstring(package.read(relpart)):
        if rel.get('TargetMode') == 'External':
            continue
        target = rel.get('Target', '')
        resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(folder, target))
        result[rel.get('Id')] = (rel.get('Type', ''), resolved)
    return result


def shape_text(shape):
    paragraphs = []
    for paragraph in shape.findall('./p:txBody/a:p', NS):
        parts = []
        for node in paragraph.iter():
            if node.tag == '{' + NS['a'] + '}t':
                parts.append(node.text or '')
            elif node.tag == '{' + NS['a'] + '}br':
                parts.append('\n')
        paragraphs.append(''.join(parts))
    return '\n'.join(paragraphs).strip()


def inspect(path):
    with ZipFile(path) as package:
        presentation = E.fromstring(package.read('ppt/presentation.xml'))
        links = relationships(package, 'ppt/presentation.xml')
        size = presentation.find('p:sldSz', NS)
        page_height = int(size.get('cy')) if size is not None else None
        slides = []
        for index, sid in enumerate(presentation.findall('p:sldIdLst/p:sldId', NS), 1):
            part = links[sid.get('{' + NS['r'] + '}id')][1]
            root = E.fromstring(package.read(part))
            records = []
            candidates = []
            for shape in root.findall('.//p:sp', NS):
                text = shape_text(shape)
                if not text:
                    continue
                properties = shape.find('p:nvSpPr/p:cNvPr', NS)
                if properties is not None and properties.get('hidden') in ('1', 'true'):
                    continue
                placeholder = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
                kind = placeholder.get('type', 'obj') if placeholder is not None else None
                off = shape.find('p:spPr/a:xfrm/a:off', NS)
                x, y = (int(off.get('x')), int(off.get('y'))) if off is not None else (None, None)
                sizes = [int(v.get('sz')) / 100 for v in shape.iter() if v.get('sz') and v.get('sz').isdigit()]
                font_size = max(sizes, default=0)
                record = {
                    'id': properties.get('id') if properties is not None else None,
                    'text': text, 'placeholder_type': kind, 'largest_font_pt': font_size,
                    'x_emu': x, 'y_emu': y,
                }
                records.append(record)
                score = font_size
                if kind in ('title', 'ctrTitle'):
                    score += 1000
                elif page_height and y is not None and 0 <= y < page_height * .14:
                    score += 50
                candidates.append((score, record))
            candidates.sort(key=lambda value: value[0], reverse=True)
            title = candidates[0][1]['text'] if candidates else ''
            flags = []
            if len(re.sub(r'\s', '', title)) > 20:
                flags.append('title_length_review')
            if any(re.search(r'(?<!\d)\d+(?:\.\d+){2,}(?!\d)', r['text']) for r in records):
                flags.append('multilevel_number_review')
            counters = []
            for record in records:
                match = re.fullmatch(r'\s*(\d{1,3})\s*/\s*(\d{1,3})\s*', record['text'])
                if match and (page_height is None or record['y_emu'] is None or record['y_emu'] > page_height * .75):
                    counters.append({'current': int(match.group(1)), 'total': int(match.group(2))})
            notes = []
            for reltype, note_part in relationships(package, part).values():
                if reltype.endswith('/notesSlide'):
                    note_root = E.fromstring(package.read(note_part))
                    for shape in note_root.findall('.//p:sp', NS):
                        ph = shape.find('p:nvSpPr/p:nvPr/p:ph', NS)
                        if ph is not None and ph.get('type') == 'body':
                            notes.append(shape_text(shape))
            slides.append({
                'slide': index, 'hidden': root.get('show') in ('0', 'false'), 'title_guess': title,
                'title_candidates': [record for _, record in candidates[:3]],
                'texts': records, 'notes': '\n'.join(filter(None, notes)),
                'visible_page_counters': counters, 'review_flags': flags,
            })
        duplicate_titles = Counter(s['title_guess'] for s in slides if s['title_guess'])
        for slide in slides:
            if duplicate_titles.get(slide['title_guess'], 0) > 1:
                slide['review_flags'].append('duplicate_title_review')
            if any(c['current'] != slide['slide'] or c['total'] != len(slides) for c in slide['visible_page_counters']):
                slide['review_flags'].append('page_counter_review')
        return {
            'schema': 'coconut-ppt.inventory.v1',
            'source': str(Path(path).resolve()), 'slide_count': len(slides),
            'hidden_slide_count': sum(s['hidden'] for s in slides),
            'title_detection': 'heuristic; verify against the actual slide',
            'coordinates': 'local shape coordinates; nested groups may transform them',
            'coverage': 'shape text and body notes only; inspect image and chart contents separately',
            'review_flag_counts': dict(Counter(flag for s in slides for flag in s['review_flags'])),
            'slides': slides,
        }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    result = inspect(args.pptx)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({key: result[key] for key in ('slide_count', 'hidden_slide_count', 'review_flag_counts')}, ensure_ascii=False))


if __name__ == '__main__':
    main()
