import sys,os,re,hashlib,shutil,zipfile,io
from PIL import Image
src,out,bgdir=sys.argv[1:4]
work=out+'.d'; shutil.rmtree(work,ignore_errors=True)
zipfile.ZipFile(src).extractall(work)
med=os.path.join(work,'ppt','media')
bgmap={'7b7b652e':'bg_cover.png','b6fee876':'bg_content.png','25a6ab58':'bg_mid.png'}
for f in os.listdir(med):
    p=os.path.join(med,f); b=open(p,'rb').read(); h=hashlib.sha1(b).hexdigest()[:8]
    if h in bgmap:
        shutil.copy(os.path.join(bgdir,bgmap[h]),p); print('bg',f,bgmap[h]); continue
    if f.endswith('.png'):
        im=Image.open(p)
        if im.mode=='RGBA' and im.size==(256,256):
            px=im.load(); hit=0; tot=0
            for y in range(0,256,4):
                for x in range(0,256,4):
                    r,g,bb,a=px[x,y]
                    if a>150: tot+=1; hit+= (bb>180 and r<90)
            if tot and hit/tot>0.6:
                new=Image.new('RGBA',im.size,(0xC8,0x15,0x7F,0)); a=im.getchannel('A'); new.putalpha(a)
                new.save(p); print('icon',f)
cmap={'2F5BEA':'C8157F','DCE6FF':'F6DCEB','D6E0FF':'F3D2E5','8FAEFF':'E79AC7','C9D5F2':'EAC9DC','6F8FE8':'D35A9E',
 'C9D7FF':'F3C3DD','E3EAFF':'FBE6F1','EEF3FF':'FDF2F8','5A82F5':'D9369A','7FA2F7':'E37DB5','3C63E0':'B3126F','4F6BD8':'4F6BD8'}
for root,_,fs in os.walk(os.path.join(work,'ppt')):
    for f in fs:
        if f.endswith('.xml') and ('slides' in root or 'charts' in root):
            p=os.path.join(root,f); s=open(p,encoding='utf-8').read(); o=s
            for a,b in cmap.items(): s=re.sub(r'val="%s"'%a,'val="%s"'%b,s,flags=re.I)
            if s!=o: open(p,'w',encoding='utf-8').write(s)
z=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED)
names=zipfile.ZipFile(src).namelist()
for n in names: z.write(os.path.join(work,n),n)
z.close(); print('saved',out)
