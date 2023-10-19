import sys
import io
from zipfile import ZipFile
from saxonche import *

# applying sequence of xsl
xsls = ["xsl/iden.xsl","xsl/text.xsl"]

# getting docx input file
file_name = sys.argv[1]
doc_zip = ZipFile(file_name, 'r')
doc_xml_file = doc_zip.open("word/document.xml",)
doc_xml  = io.TextIOWrapper(doc_xml_file, encoding='utf-8', newline='').read()


#proc = PySaxonProcessor(license=False)
with PySaxonProcessor(license=False) as proc:
    try:
        out_xml = doc_xml
        for i,xsl in enumerate(xsls):
            xsltproc = proc.new_xslt30_processor()
            document = proc.parse_xml(xml_text=out_xml)
            executable = xsltproc.compile_stylesheet(stylesheet_file=xsl)
            out_xml = executable.transform_to_string(xdm_node=document)
            print("output{}: {}".format(i+1,out_xml))
    except PySaxonApiError as err:
        print('Error during function call', err)
