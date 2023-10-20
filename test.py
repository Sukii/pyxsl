import sys
import pyxsl
from zipfile import ZipFile


# applying sequence of xsl
xsls = ["xsl/iden.xsl","xsl/text.xsl"]

# getting docx input file

docx_file = sys.argv[1]
docx_out_file = docx_file.replace(".docx","_output.docx")


xml_file = "word/document.xml"
doc_zip = ZipFile(docx_file, "r")

# apply sequence of xsl on docx to extract xml
out_xml = pyxsl.xsldocx(doc_zip,xml_file,xsls)

# update xml in docx and save
pyxsl.writeModifiedDocx(doc_zip,docx_out_file,xml_file,out_xml)
