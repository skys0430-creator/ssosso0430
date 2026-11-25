import sys,io
from pptx import Presentation
from PIL import Image
src,out,figdir=sys.argv[1:4]
m={'7b7b652e':'bg_cover.png','b6fee876':'bg_content.png','25a6ab58':'bg_mid.png'}
p=Presentation(src)
n=0
for s in p.slides:
    for sh in list(s.shapes):
        if sh.shape_type==13 and sh.name.startswith('bg-') and sh.image.sha1[:8] in m:
            f=figdir+'/'+m[sh.image.sha1[:8]]
            new=s.shapes.add_picture(f,sh.left,sh.top,sh.width,sh.height)
            new.name=sh.name
            sh._element.addnext(new._element)
            sh._element.getparent().remove(sh._element)
            n+=1
p.save(out); print('replaced',n)
