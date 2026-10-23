#!/usr/bin/env python3
"""Render and check the actual five-page PPTX, rather than a separate mockup."""
from pathlib import Path
import subprocess
import fitz
from PIL import Image, ImageOps, ImageDraw

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'downloads/presentation_5slides'
PPTX = OUT / 'thz_background_objectives_5slides_revised.pptx'
subprocess.run(['soffice','--headless','--convert-to','pdf','--outdir',str(OUT),str(PPTX)],check=True,cwd=ROOT)
PDF = PPTX.with_suffix('.pdf')
expected = ['研究背景及现状','现有电磁建模方法','现有研究面临的具体挑战','研究目标','研究目标实现面临的困难']
tiles = []
with fitz.open(PDF) as pdf:
    assert len(pdf) == 5, f'Expected five slides, found {len(pdf)}'
    for i,page in enumerate(pdf):
        assert expected[i] in page.get_text(), (i+1,expected[i])
        target = OUT / f'revised_slide_{i+1:02d}.png'
        page.get_pixmap(dpi=144,alpha=False).save(target)
        # The contact sheet is a document preview, not an edited scientific asset.
        tile = Image.open(target).convert('RGB')
        tile.thumbnail((800,450))
        framed=Image.new('RGB',(824,490),'#E6EDF4')
        framed.paste(tile,((824-tile.width)//2,12))
        ImageDraw.Draw(framed).text((16,468),f'{i+1} / 5',fill='#155EA7')
        tiles.append(framed)
sheet=Image.new('RGB',(1648,1470),'white')
for i,tile in enumerate(tiles):sheet.paste(tile,((i%2)*824,(i//2)*490))
sheet.save(OUT/'preview_revised.png')
print('Rendered five pages, Chinese headings present:',OUT)
