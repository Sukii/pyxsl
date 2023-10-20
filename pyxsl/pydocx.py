import sys
import io
from zipfile import ZipFile
from saxonche import *


def xslDocx(doc_zip,xml_file,xsls):
    doc_xml_file = doc_zip.open(xml_file,"r")
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
                #print("output{}: {}".format(i+1,out_xml))
        except PySaxonApiError as err:
            print('Error during function call', err)
    return out_xml


def writeModifiedDocx(input_docx,docx_output_path,xml_paths,xml_outputs):
    # write output to docx file
    # there is no simple command to replace a file in zip
    # Iterate the input files
    with ZipFile(docx_output_path, mode="w") as output_docx:
        for info in input_docx.infolist():
            # Read input file
            with input_docx.open(info) as infile:
                content = infile.read()
                if info.filename in xml_paths:
                    # copy the modified content
                    i = xml_paths.index(info.filename)
                    output_docx.writestr(info.filename, xml_outputs[i])
                else: # copy the content
                    output_docx.writestr(info.filename, content)


