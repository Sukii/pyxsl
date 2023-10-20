import sys
from pyxsl import pydocx
from zipfile import ZipFile


# applying sequence of xsl
xsls = ["xsl/iden.xsl","xsl/text.xsl"]

# getting docx input file

docx_file = sys.argv[1]
docx_output_path = docx_file.replace(".docx","_output.docx")


docx = ZipFile(docx_file, "r")


xml_paths = ["word/document.xml"]

# apply sequence of xsl on docx to extract xml
xml_outputs = []
xml_outputs.append(pydocx.xslDocx(docx,xml_paths[0],xsls))

# update xml in docx and save
pydocx.writeModifiedDocx(docx,docx_output_path,xml_paths,xml_outputs)
