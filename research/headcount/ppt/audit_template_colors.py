"""템플릿 슬라이드의 도형 채우기·선·글자 색을 테마 기준으로 실제 RGB로 풀어서 출력한다."""
import sys, zipfile, re, colorsys, collections
from lxml import etree
NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
z=zipfile.ZipFile(sys.argv[1])
SCHEME={'dk1':'000000','lt1':'FFFFFF','dk2':'0E2841','lt2':'E8E8E8','accent1':'156082','accent2':'E97132','accent3':'196B24','accent4':'0F9ED5','accent5':'A02B93','accent6':'4EA72E','hlink':'467886','folHlink':'96607D'}
MAP={'bg1':'lt1','tx1':'dk1','bg2':'lt2','tx2':'dk2'}
PRST={'black':'000000','white':'FFFFFF'}
def hsl_mod(hexv,mods):
    r,g,b=[int(hexv[i:i+2],16)/255 for i in (0,2,4)]
    h,l,s=colorsys.rgb_to_hls(r,g,b)
    for tag,val in mods:
        v=int(val)/100000
        if tag=='lumMod': l*=v
        elif tag=='lumOff': l+=v
    l=max(0,min(1,l)); r,g,b=colorsys.hls_to_rgb(h,l,s)
    return '%02X%02X%02X'%(round(r*255),round(g*255),round(b*255))
def resolve(el):
    if el is None: return None
    for c in el:
        tag=etree.QName(c).localname
        mods=[(etree.QName(m).localname,m.get('val')) for m in c if etree.QName(m).localname in ('lumMod','lumOff')]
        alpha=[m.get('val') for m in c if etree.QName(m).localname=='alpha']
        if tag=='srgbClr': base=c.get('val')
        elif tag=='schemeClr': base=SCHEME[MAP.get(c.get('val'),c.get('val'))]
        elif tag=='sysClr': base=c.get('lastClr')
        elif tag=='prstClr': base=PRST.get(c.get('val'),c.get('val'))
        else: continue
        out=hsl_mod(base,mods) if mods else base
        desc=f"{tag}:{c.get('val')}"+(''.join(f' {t}={v}' for t,v in mods))+(f' alpha={alpha[0]}' if alpha else '')
        return out,desc
    return None
stats=collections.Counter()
for n in sys.argv[2:]:
    x=etree.fromstring(z.read(f'ppt/slides/slide{n}.xml'))
    print(f'===== slide{n}')
    for sp in x.iter('{%s}sp'%NS['p'],'{%s}cxnSp'%NS['p']):
        spPr=sp.find('p:spPr',NS)
        fill=resolve(spPr.find('a:solidFill',NS)) if spPr is not None else None
        ln=spPr.find('a:ln',NS) if spPr is not None else None
        line=resolve(ln.find('a:solidFill',NS)) if ln is not None else None
        geom=spPr.find('a:prstGeom',NS).get('prst') if spPr is not None and spPr.find('a:prstGeom',NS) is not None else ''
        txt=''.join(sp.itertext()).strip()[:28]
        tc=[resolve(r.find('a:solidFill',NS)) for r in sp.iter('{%s}rPr'%NS['a']) if r.find('a:solidFill',NS) is not None]
        tcs=sorted(set(t[0] for t in tc if t))
        sz=sorted(set(r.get('sz') for r in sp.iter('{%s}rPr'%NS['a']) if r.get('sz')))
        if fill or line or tcs:
            print(f"{geom:10s} fill={fill[0]+' ('+fill[1]+')' if fill else '-':38s} line={line[0] if line else '-':7s} text={','.join(tcs) or '-':22s} sz={','.join(sz)} | {txt}")
        for v in [fill,line]+tc:
            if v: stats[(v[0],v[1])]+=1
    # table cells
    for tc_ in x.iter('{%s}tcPr'%NS['a']):
        f=resolve(tc_.find('a:solidFill',NS))
        if f: stats[(f[0],'cell '+f[1])]+=1
print('===== color frequency')
for (c,d),k in stats.most_common(40): print(k,c,d)
