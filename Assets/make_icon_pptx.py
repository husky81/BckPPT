"""BckPPT 아이콘 SVG → PPTX 변환 스크립트"""
from pptx import Presentation
from pptx.util import Emu, Cm
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree
import copy

# 512 SVG units = 10cm in PPTX
SCALE = Cm(10).emu / 512  # EMU per SVG unit

def svg_to_emu(v):
    return int(v * SCALE)

def make_freeform(slide, points_groups, left, top, width, height, fill_color, fill_alpha=None):
    """Create a freeform shape from SVG-like path data."""
    # Add a freeform shape via XML
    sp_tree = slide.shapes._spTree

    sp = etree.SubElement(sp_tree, qn('p:sp'))
    nvSpPr = etree.SubElement(sp, qn('p:nvSpPr'))
    cNvPr = etree.SubElement(nvSpPr, qn('p:cNvPr'))
    cNvPr.set('id', str(len(slide.shapes) + 10))
    cNvPr.set('name', 'Freeform')
    cNvSpPr = etree.SubElement(nvSpPr, qn('p:cNvSpPr'))
    nvPr = etree.SubElement(nvSpPr, qn('p:nvPr'))

    spPr = etree.SubElement(sp, qn('p:spPr'))

    # Transform
    xfrm = etree.SubElement(spPr, qn('a:xfrm'))
    off = etree.SubElement(xfrm, qn('a:off'))
    off.set('x', str(left))
    off.set('y', str(top))
    ext = etree.SubElement(xfrm, qn('a:ext'))
    ext.set('cx', str(width))
    ext.set('cy', str(height))

    # Custom geometry
    custGeom = etree.SubElement(spPr, qn('a:custGeom'))
    avLst = etree.SubElement(custGeom, qn('a:avLst'))
    gdLst = etree.SubElement(custGeom, qn('a:gdLst'))
    ahLst = etree.SubElement(custGeom, qn('a:ahLst'))
    cxnLst = etree.SubElement(custGeom, qn('a:cxnLst'))

    rect = etree.SubElement(custGeom, qn('a:rect'))
    rect.set('l', '0')
    rect.set('t', '0')
    rect.set('r', str(width))
    rect.set('b', str(height))

    pathLst = etree.SubElement(custGeom, qn('a:pathLst'))
    path_el = etree.SubElement(pathLst, qn('a:path'))
    path_el.set('w', str(width))
    path_el.set('h', str(height))

    for group in points_groups:
        for i, cmd in enumerate(group):
            if cmd[0] == 'M':
                moveTo = etree.SubElement(path_el, qn('a:moveTo'))
                pt = etree.SubElement(moveTo, qn('a:pt'))
                pt.set('x', str(cmd[1]))
                pt.set('y', str(cmd[2]))
            elif cmd[0] == 'L':
                lnTo = etree.SubElement(path_el, qn('a:lnTo'))
                pt = etree.SubElement(lnTo, qn('a:pt'))
                pt.set('x', str(cmd[1]))
                pt.set('y', str(cmd[2]))
            elif cmd[0] == 'C':
                cubicBezTo = etree.SubElement(path_el, qn('a:cubicBezTo'))
                for j in range(3):
                    pt = etree.SubElement(cubicBezTo, qn('a:pt'))
                    pt.set('x', str(cmd[1 + j*2]))
                    pt.set('y', str(cmd[2 + j*2]))
            elif cmd[0] == 'Z':
                etree.SubElement(path_el, qn('a:close'))

    # Fill
    solidFill = etree.SubElement(spPr, qn('a:solidFill'))
    srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgbClr.set('val', fill_color)
    if fill_alpha is not None:
        alpha = etree.SubElement(srgbClr, qn('a:alpha'))
        alpha.set('val', str(int(fill_alpha * 1000)))  # percentage * 1000

    # No line
    ln = etree.SubElement(spPr, qn('a:ln'))
    etree.SubElement(ln, qn('a:noFill'))

    return sp


def make_rounded_rect(slide, left, top, width, height):
    """Create rounded rectangle with gradient fill matching the SVG background."""
    sp_tree = slide.shapes._spTree

    sp = etree.SubElement(sp_tree, qn('p:sp'))
    nvSpPr = etree.SubElement(sp, qn('p:nvSpPr'))
    cNvPr = etree.SubElement(nvSpPr, qn('p:cNvPr'))
    cNvPr.set('id', str(len(slide.shapes) + 10))
    cNvPr.set('name', 'Background')
    cNvSpPr = etree.SubElement(nvSpPr, qn('p:cNvSpPr'))
    nvPr = etree.SubElement(nvSpPr, qn('p:nvPr'))

    spPr = etree.SubElement(sp, qn('p:spPr'))

    # Transform
    xfrm = etree.SubElement(spPr, qn('a:xfrm'))
    off = etree.SubElement(xfrm, qn('a:off'))
    off.set('x', str(left))
    off.set('y', str(top))
    ext = etree.SubElement(xfrm, qn('a:ext'))
    ext.set('cx', str(width))
    ext.set('cy', str(height))

    # Rounded rectangle preset geometry
    prstGeom = etree.SubElement(spPr, qn('a:prstGeom'))
    prstGeom.set('prst', 'roundRect')
    avLst = etree.SubElement(prstGeom, qn('a:avLst'))
    # Corner radius ~90/512 = 17.6%
    gd = etree.SubElement(avLst, qn('a:gd'))
    gd.set('name', 'adj')
    gd.set('fmla', 'val 17578')  # 17.578%

    # Gradient fill (top to bottom)
    gradFill = etree.SubElement(spPr, qn('a:gradFill'))
    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))

    gs1 = etree.SubElement(gsLst, qn('a:gs'))
    gs1.set('pos', '0')
    srgb1 = etree.SubElement(gs1, qn('a:srgbClr'))
    srgb1.set('val', '4BA3F5')

    gs2 = etree.SubElement(gsLst, qn('a:gs'))
    gs2.set('pos', '100000')
    srgb2 = etree.SubElement(gs2, qn('a:srgbClr'))
    srgb2.set('val', '2178E8')

    lin = etree.SubElement(gradFill, qn('a:lin'))
    lin.set('ang', '5400000')  # 90 degrees (top to bottom)
    lin.set('scaled', '1')

    # No line
    ln = etree.SubElement(spPr, qn('a:ln'))
    etree.SubElement(ln, qn('a:noFill'))

    return sp


def main():
    prs = Presentation()
    # Blank slide
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)

    size = svg_to_emu(512)

    # Center on slide
    slide_w = prs.slide_width.emu
    slide_h = prs.slide_height.emu
    origin_x = (slide_w - size) // 2
    origin_y = (slide_h - size) // 2

    # 1) Rounded rectangle background with gradient
    make_rounded_rect(slide, origin_x, origin_y, size, size)

    # 2) White B letter — outer shape + two holes (evenodd)
    # Convert SVG coordinates to EMU relative to shape origin
    def e(v):
        return svg_to_emu(v)

    # All paths relative to the B bounding box
    # B bounding box in SVG: x=130..395, y=80..432
    bx, by = 130, 80
    bw, bh = 395 - 130, 432 - 80  # 265 x 352

    def rel(x, y):
        """Convert SVG coords to EMU coords relative to B bounding box."""
        return (svg_to_emu(x - bx), svg_to_emu(y - by))

    def r(x, y):
        return rel(x, y)

    # Outer path
    outer = [
        ('M', *r(130, 80)),
        ('L', *r(280, 80)),
        ('C', *r(340, 80), *r(380, 100), *r(380, 155)),
        ('C', *r(380, 200), *r(355, 230), *r(310, 242)),
        ('C', *r(365, 254), *r(395, 282), *r(395, 335)),
        ('C', *r(395, 400), *r(355, 432), *r(280, 432)),
        ('L', *r(130, 432)),
        ('Z',),
    ]

    # Upper hole (counter-clockwise for subtraction)
    upper_hole = [
        ('M', *r(210, 128)),
        ('L', *r(265, 128)),
        ('C', *r(295, 128), *r(310, 140), *r(310, 157)),
        ('C', *r(310, 174), *r(295, 186), *r(265, 186)),
        ('L', *r(210, 186)),
        ('Z',),
    ]

    # Lower hole
    lower_hole = [
        ('M', *r(210, 290)),
        ('L', *r(272, 290)),
        ('C', *r(308, 290), *r(325, 304), *r(325, 328)),
        ('C', *r(325, 352), *r(308, 366), *r(272, 366)),
        ('L', *r(210, 366)),
        ('Z',),
    ]

    shape_left = origin_x + svg_to_emu(bx)
    shape_top = origin_y + svg_to_emu(by)
    shape_w = svg_to_emu(bw)
    shape_h = svg_to_emu(bh)

    make_freeform(slide, [outer, upper_hole, lower_hole],
                  shape_left, shape_top, shape_w, shape_h,
                  'FFFFFF')

    out_path = r'C:\Users\bckim\BckSoft\cellbox2\apps\tray\BckPPT\Assets\icon.pptx'
    prs.save(out_path)
    print(f'Saved: {out_path}')


if __name__ == '__main__':
    main()
