#!/usr/bin/env python
import zipfile
import datetime
import tempfile
import os
from lxml import etree as ET

# add more as needed
namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}


def get_attr(element, attr):
    index = '{%s}%s' % (namespaces['w'], attr)
    if index not in element.attrib:
        return None
    return element.attrib[index]


def set_text_input(docxml, name, value):
    start_selector = './/w:bookmarkStart[@w:name="%s"]' % name
    bookmark_start = docxml.find(start_selector, namespaces)
    bookmark_id = get_attr(bookmark_start, 'id')
    bookmark_end = docxml.find(
        './/w:bookmarkEnd[@w:id="%s"]' % (bookmark_id), namespaces)

    parent = docxml.find('%s/..' % start_selector, namespaces)
    children = parent.getchildren()
    start_index = children.index(bookmark_start)
    end_index = children.index(bookmark_end)
    for i in range(start_index + 1, end_index):
        node = children[i]
        # print i, node, node.attrib
        # print dir(node)
        # node_rsid = get_attr(node, 'rsidR')
        # if node_rsid == '00E32656':
        #   print node
        text_input = node.find('.//w:t', namespaces)
        if text_input is not None:
            text_input.text = value
            break


def update_zip_in_place(
        input_zipfile,
        output_zipfile_name,
        replaced_file_in_zip_name,
        new_file_data
        ):
    print "updating file", replaced_file_in_zip_name, 'in zipfile'
    print "writing updated file to", output_zipfile_name
    with zipfile.ZipFile(output_zipfile_name, 'w') as zout:
        zout.comment = input_zipfile.comment  # preserve the comment
        # copy all files that are not the one we are replacing
        for item in input_zipfile.infolist():
            if item.filename != replaced_file_in_zip_name:
                zout.writestr(item, input_zipfile.read(item.filename))

    # now add replaced_file_in_zip_name with its new data
    with zipfile.ZipFile(
        output_zipfile_name,
        mode='a',
        compression=zipfile.ZIP_DEFLATED
    ) as cloned_zip_file:
        cloned_zip_file.writestr(replaced_file_in_zip_name, new_file_data)


in_name = 'report_template.docx'
out_name = 'report_template_out.docx'
docx_namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

ET.register_namespace('w', docx_namespace)
input_zip_file = zipfile.ZipFile(in_name, 'r')
xml_text = input_zip_file.read('word/document.xml')
docxml = ET.fromstring(xml_text)

# set_text_input(docxml, 'Text412', 'hi')

xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
xml_body = ET.tostring(docxml)

update_zip_in_place(
    input_zip_file,
    out_name,
    'word/document.xml',
    xml_header + xml_body
)
