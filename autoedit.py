import sys
import io
from zipfile import ZipFile
from saxonche import *

file_name = sys.argv[1]
doc_zip = ZipFile(file_name, 'r')
doc_xml_file = doc_zip.open("word/document.xml",)
doc_xml  = io.TextIOWrapper(doc_xml_file, encoding='utf-8', newline='').read()

#proc = PySaxonProcessor(license=False)
with PySaxonProcessor(license=False) as proc:
    try:
        xsltproc = proc.new_xslt30_processor()
        document = proc.parse_xml(xml_text=doc_xml)
        executable = xsltproc.compile_stylesheet(stylesheet_file="test.xsl")
        output = executable.transform_to_string(xdm_node=document)
        print(output)
    except PySaxonApiError as err:
        print('Error during function call', err)
