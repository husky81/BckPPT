"""PPTX → SVG 변환: 슬라이드 1의 도형을 SVG로 추출"""
from pptx import Presentation
from pptx.util import Emu
from pptx.oxml.ns import qn
from lxml import etree
import re

EMU_PER_PX = 914400 / 96  # 1px = 9525 EMU


def emu_to_px(emu):
    return round(emu / EMU_PER_PX, 2)


def parse_gradient_fill(grad_el):
    """Parse gradient fill element to SVG gradient."""
    stops = []
    for gs in grad_el.findall(qn('a:gsLst') + '/' + qn('a:gs')):
        pos = int(gs.get('pos', '0')) / 1000  # percentage
        srgb = gs.find(qn('a:srgbClr'))
        if srgb is not None:
            color = '#' + srgb.get('val')
            stops.append((pos, color))

    lin = grad_el.find(qn('a:lin'))
    angle = 90  # default top-to-bottom
    if lin is not None:
        angle = int(lin.get('ang', '5400000')) / 60000

    return stops, angle


def parse_solid_fill(fill_el):
    """Parse solid fill to hex color + alpha."""
    srgb = fill_el.find(qn('a:srgbClr'))
    if srgb is not None:
        color = '#' + srgb.get('val')
        alpha_el = srgb.find(qn('a:alpha'))
        alpha = 1.0
        if alpha_el is not None:
            alpha = int(alpha_el.get('val')) / 100000
        return color, alpha
    return None, None


def parse_custom_geometry(custGeom, off_x, off_y, ext_cx, ext_cy):
    """Convert OOXML custom geometry paths to SVG path d attribute."""
    path_lst = custGeom.find(qn('a:pathLst'))
    if path_lst is None:
        return ''

    svg_paths = []
    for path_el in path_lst.findall(qn('a:path')):
        pw = int(path_el.get('w', str(ext_cx)))
        ph = int(path_el.get('h', str(ext_cy)))
        sx = ext_cx / pw if pw else 1
        sy = ext_cy / ph if ph else 1

        d_parts = []
        for child in path_el:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'moveTo':
                pt = child.find(qn('a:pt'))
                x = emu_to_px(off_x + float(pt.get('x')) * sx)
                y = emu_to_px(off_y + float(pt.get('y')) * sy)
                d_parts.append(f'M {x},{y}')

            elif tag == 'lnTo':
                pt = child.find(qn('a:pt'))
                x = emu_to_px(off_x + float(pt.get('x')) * sx)
                y = emu_to_px(off_y + float(pt.get('y')) * sy)
                d_parts.append(f'L {x},{y}')

            elif tag == 'cubicBezTo':
                pts = child.findall(qn('a:pt'))
                coords = []
                for pt in pts:
                    x = emu_to_px(off_x + float(pt.get('x')) * sx)
                    y = emu_to_px(off_y + float(pt.get('y')) * sy)
                    coords.extend([str(x), str(y)])
                d_parts.append(f'C {",".join(coords)}')

            elif tag == 'close':
                d_parts.append('Z')

        svg_paths.append(' '.join(d_parts))

    return ' '.join(svg_paths)


def preset_rounded_rect(off_x, off_y, cx, cy, adj_val):
    """Generate SVG path for a rounded rectangle preset."""
    r = min(cx, cy) * adj_val / 100000
    x1, y1 = emu_to_px(off_x), emu_to_px(off_y)
    w, h = emu_to_px(cx), emu_to_px(cy)
    rr = emu_to_px(r)

    return (f'M {x1+rr},{y1} L {x1+w-rr},{y1} '
            f'Q {x1+w},{y1} {x1+w},{y1+rr} '
            f'L {x1+w},{y1+h-rr} '
            f'Q {x1+w},{y1+h} {x1+w-rr},{y1+h} '
            f'L {x1+rr},{y1+h} '
            f'Q {x1},{y1+h} {x1},{y1+h-rr} '
            f'L {x1},{y1+rr} '
            f'Q {x1},{y1} {x1+rr},{y1} Z')


def main():
    pptx_path = r'C:\Users\bckim\BckSoft\cellbox2\apps\tray\BckPPT\Assets\icon.pptx'
    svg_path = r'C:\Users\bckim\BckSoft\cellbox2\apps\tray\BckPPT\Assets\icon.svg'

    prs = Presentation(pptx_path)
    slide = prs.slides[0]

    # Find bounding box of all shapes
    all_shapes = list(slide.shapes._spTree.findall(qn('p:sp')))

    svg_elements = []
    defs = []
    grad_id = 0

    for sp in all_shapes:
        spPr = sp.find(qn('p:spPr'))
        if spPr is None:
            continue

        xfrm = spPr.find(qn('a:xfrm'))
        if xfrm is None:
            continue

        off = xfrm.find(qn('a:off'))
        ext = xfrm.find(qn('a:ext'))
        off_x = int(off.get('x', '0'))
        off_y = int(off.get('y', '0'))
        ext_cx = int(ext.get('cx', '0'))
        ext_cy = int(ext.get('cy', '0'))

        # Determine geometry
        d = ''
        prstGeom = spPr.find(qn('a:prstGeom'))
        custGeom = spPr.find(qn('a:custGeom'))

        if prstGeom is not None:
            prst = prstGeom.get('prst', '')
            if prst == 'roundRect':
                avLst = prstGeom.find(qn('a:avLst'))
                adj = 16667  # default
                if avLst is not None:
                    gd = avLst.find(qn('a:gd'))
                    if gd is not None:
                        fmla = gd.get('fmla', '')
                        m = re.search(r'val\s+(\d+)', fmla)
                        if m:
                            adj = int(m.group(1))
                d = preset_rounded_rect(off_x, off_y, ext_cx, ext_cy, adj)
            elif prst == 'rect':
                x1, y1 = emu_to_px(off_x), emu_to_px(off_y)
                w, h = emu_to_px(ext_cx), emu_to_px(ext_cy)
                d = f'M {x1},{y1} L {x1+w},{y1} L {x1+w},{y1+h} L {x1},{y1+h} Z'

        elif custGeom is not None:
            d = parse_custom_geometry(custGeom, off_x, off_y, ext_cx, ext_cy)

        if not d:
            continue

        # Determine fill
        fill_attr = ''
        gradFill = spPr.find(qn('a:gradFill'))
        solidFill = spPr.find(qn('a:solidFill'))

        if gradFill is not None:
            stops, angle = parse_gradient_fill(gradFill)
            grad_id += 1
            gid = f'grad{grad_id}'
            # SVG linear gradient
            if angle == 90:
                grad_svg = f'<linearGradient id="{gid}" x1="0" y1="0" x2="0" y2="1">'
            else:
                grad_svg = f'<linearGradient id="{gid}">'
            for pos, color in stops:
                grad_svg += f'<stop offset="{pos}%" stop-color="{color}"/>'
            grad_svg += '</linearGradient>'
            defs.append(grad_svg)
            fill_attr = f'fill="url(#{gid})"'

        elif solidFill is not None:
            color, alpha = parse_solid_fill(solidFill)
            if color:
                fill_attr = f'fill="{color}"'
                if alpha is not None and alpha < 1.0:
                    fill_attr += f' opacity="{alpha}"'

        # Check for no line
        ln = spPr.find(qn('a:ln'))
        stroke_attr = ''
        if ln is not None and ln.find(qn('a:noFill')) is not None:
            stroke_attr = 'stroke="none"'

        # Build fill-rule
        fill_rule = 'fill-rule="evenodd"' if custGeom is not None else ''

        svg_elements.append(f'  <path {fill_rule} {fill_attr} {stroke_attr} d="{d}"/>')

    # Compute viewBox from shapes
    # Use slide dimensions as fallback
    sw = emu_to_px(prs.slide_width.emu)
    sh = emu_to_px(prs.slide_height.emu)

    # Build SVG
    defs_str = ''
    if defs:
        defs_str = '  <defs>\n    ' + '\n    '.join(defs) + '\n  </defs>\n'

    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {sw} {sh}" width="{sw}" height="{sh}">
{defs_str}{chr(10).join(svg_elements)}
</svg>'''

    with open(svg_path, 'w', encoding='utf-8') as f:
        f.write(svg)
    print(f'Saved: {svg_path}')


if __name__ == '__main__':
    main()
