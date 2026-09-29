from docx import Document
from wps_shapes import Canvas, emu
doc = Document()
cv = Canvas(doc, 20.0, 10.0)
cv.rect(emu(1), emu(1), emu(5), emu(3), fill="FFC9C9", line="333333", text="Привет\nтест", size=10, bold=True)
cv.line(emu(6), emu(4), emu(12), emu(8), color="D94A38", w=1.5, arrow_end=True)
cv.poly([(emu(8), emu(1)), (emu(10), emu(3)), (emu(7), emu(3))], closed=True, fill="B7E1CD", line="333")
doc.save('/tmp/test.docx')
import zipfile
z=zipfile.ZipFile('/tmp/test.docx')
x=z.read('word/document.xml').decode('utf8')
i=x.find('<wp:anchor')
print(x[i:i+2500])
