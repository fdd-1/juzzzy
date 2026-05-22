import json, sys
sys.stdout.reconfigure(encoding='utf-8')

with open('C:/Users/fengjianyi/Desktop/周报自动化/doc_blocks.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

items = data['data']['items']
block_map = {item['block_id']: item for item in items}

def get_text(block):
    for key in ['heading1','heading2','heading3','heading4','heading5','heading6','heading7','heading8','heading9','paragraph']:
        if key in block:
            elements = block[key].get('elements', [])
            text = ''
            for el in elements:
                if 'text_run' in el:
                    text += el['text_run'].get('content', '')
            return text
    return None

def get_cell_text(cell_id):
    cell = block_map.get(cell_id)
    if not cell:
        return ''
    children = cell.get('children', [])
    texts = []
    for child_id in children:
        child = block_map.get(child_id)
        if child:
            t = get_text(child)
            if t:
                texts.append(t)
    return '|'.join(texts) if texts else ''

# Find section 4 and parse tables (type=31)
in_section = False

for item in items:
    bt = item.get('block_type')
    text = get_text(item)
    bid = item.get('block_id')

    if not in_section:
        if text and '4' in text and '服务' in text and bt == 5:
            in_section = True
            print(f'=== {text} ===\n')
        continue

    # Stop at next major section
    if text and text.strip() and bt == 5:
        break

    # Sub-headings
    if bt == 8 and text:
        print(f'\n## {text}')

    # Tables (type=31)
    if bt == 31:
        table_info = item.get('table', {})
        prop = table_info.get('property', {})
        rows = prop.get('row_size', 0)
        cols = prop.get('column_size', 0)
        cells = table_info.get('cells', [])
        merge_info = table_info.get('merge_info', [])

        print(f'  [TABLE {rows}x{cols}]')
        for r in range(min(rows, 8)):
            row_cells = []
            for c in range(cols):
                idx = r * cols + c
                if idx < len(cells):
                    cell_text = get_cell_text(cells[idx])
                    row_cells.append(cell_text)
                else:
                    row_cells.append('')
            print(f'  R{r}: ' + ' | '.join(row_cells))
        if rows > 8:
            print(f'  ... ({rows - 8} more rows)')
        print()

    # Type 30 (embedded content like sheets/images)
    if bt == 30:
        embed_info = {}
        for k in item.keys():
            if k not in ['block_id', 'block_type', 'parent_id', 'children']:
                embed_info[k] = item[k]
        if embed_info:
            print(f'  [EMBED] {json.dumps(embed_info, ensure_ascii=False)[:300]}')

    # Paragraphs with text
    if bt == 2 and text and text.strip():
        print(f'  P: {text[:200]}')
