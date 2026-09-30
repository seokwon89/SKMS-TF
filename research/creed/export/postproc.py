import re, sys, zipfile, shutil, os
src, dst = sys.argv[1], sys.argv[2]
tmp='pp_tmp'; shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
zipfile.ZipFile(src).extractall(tmp)
def rw(p, fn):
    fp=os.path.join(tmp,p); s=open(fp,encoding='utf-8').read(); s=fn(s); open(fp,'w',encoding='utf-8').write(s)
rw('word/styles.xml', lambda s: re.sub(r'<w:keepNext /><w:pageBreakBefore />(\s*)<w:keepLines />',r'<w:keepNext />\1<w:keepLines /><w:pageBreakBefore />',s))
sect='<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1" /><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1247" w:right="1191" w:bottom="1247" w:left="1191" w:header="709" w:footer="567" w:gutter="0" /></w:sectPr>'
def doc(d):
    if '<w:sectPr' in d: return re.sub(r'<w:sectPr[^>]*>.*?</w:sectPr>|<w:sectPr[^>]*/>',sect,d,flags=re.S)
    return d.replace('</w:body>',sect+'</w:body>')
rw('word/document.xml', doc)
run='<w:r><w:rPr><w:sz w:val="16"/><w:color w:val="7A8288"/></w:rPr>{}</w:r>'
footer='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'+run.format('<w:t xml:space="preserve">선언에서 작동으로 · 초안 v0.1 · </w:t>')+run.format('<w:fldChar w:fldCharType="begin"/>')+run.format('<w:instrText xml:space="preserve"> PAGE </w:instrText>')+run.format('<w:fldChar w:fldCharType="separate"/>')+run.format('<w:t>1</w:t>')+run.format('<w:fldChar w:fldCharType="end"/>')+'</w:p></w:ftr>'
open(os.path.join(tmp,'word/footer1.xml'),'w',encoding='utf-8').write(footer)
rw('word/_rels/document.xml.rels', lambda r: r if 'rIdFooter1' in r else r.replace('</Relationships>','<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'))
rw('[Content_Types].xml', lambda c: c if 'footer1.xml' in c else c.replace('</Types>','<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'))
def st(s):
    s=s.replace('<w:updateFields w:val="true" />','')
    m=re.search(r'<w:(hdrShapeDefaults|footnotePr|endnotePr|compat|docVars|rsids)\b',s)
    return s[:m.start()]+'<w:updateFields w:val="true" />'+s[m.start():]
rw('word/settings.xml', st)
if os.path.exists(dst): os.remove(dst)
z=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
# [Content_Types].xml first
z.write(os.path.join(tmp,'[Content_Types].xml'),'[Content_Types].xml')
for root,_,files in os.walk(tmp):
    for f in files:
        full=os.path.join(root,f); arc=os.path.relpath(full,tmp)
        if arc=='[Content_Types].xml': continue
        z.write(full,arc)
z.close(); shutil.rmtree(tmp)
