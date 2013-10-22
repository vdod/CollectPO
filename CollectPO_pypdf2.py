#coding:utf-8
import os, re, PyPDF2
from time import time
t = time()
out = ""
#srcfolder 希望能做成可以選擇的
srcfolder = r"."
#srcfolder = r"C:\Users\Aaron\Dropbox\翻譯部文件\01說明文件\TRL\2013-10Oct".decode('utf-8')

#fo 是等一下要用的檔案名稱
fo = os.path.basename(os.path.abspath(srcfolder)) + u".txt"

#從 RegexBuddy 來的
reobj = re.compile(r"(?:.*?No\.\s+)(TRL-\d+)(?:PO Ref)((?:PD|TR|WR)[\d-]+)(?:.*?)(\d+\.\d+)TGP(\d+\.\d{2})(?:.*)",
                   re.IGNORECASE | re.MULTILINE)
reobj2 = re.compile(r"(?:.*?Total Amount)(\d+\.\d{2})(?:.*)", re.IGNORECASE | re.MULTILINE)


def getPDFContent(path):
    content = ""
    # Load PDF into PyPDF2
    pdf = PyPDF2.PdfFileReader(file(path, "rb"))
    pdf.decrypt('')
    # Iterate pages
    for i in range(0, pdf.getNumPages()):
        # Extract text from page and add to content
        content += pdf.getPage(i).extractText() + "/n"
        # Collapse whitespace
    content = " ".join(content.replace(u"/xa0", " ").split())
    return content


#跑目錄下所有檔案
for a, b, c in os.walk(srcfolder.decode('utf-8')):
    for cs in c:
        filename = os.path.join(a, cs)
        if filename.endswith(".pdf"):
            try:
                subject = getPDFContent(filename).encode("utf-8", "xmlcharrefreplace")
                m = reobj.match(subject)
                m_Total = reobj2.match(subject)
                #                print "-" *40
                print "m.group(1) is " + m.group(1)
                print "m.group(2) is " + m.group(2)
                print "m.group(3) is " + m.group(3)
                print "m.group(4) is " + m.group(4)
                print m_Total.group(1)
                m3 = round(float(m_Total.group(1)) / float(m.group(4)), 2)
                out += m.group(1) + "\t" + m.group(2) + "\t" + "\t" + str(m3) + "\t" + "\t" + m.group(4) + "\n"

            except:
                pass
print time() -t


#將結果存到txt檔案中，以便複製貼上到invoice 檔
fo = os.path.join(os.path.abspath(srcfolder), fo)
with open(fo, 'w') as f:
    f.write(out)
