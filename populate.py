#!/usr/bin/env python2
import zipfile
import datetime
import tempfile
import os
from lxml import etree as ET
from jellyfish import levenshtein_distance

# add more as needed
namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

def log_warning(warning):
    print "WARNING:", warning

def get_attr(element, attr):
    index = '{%s}%s' % (namespaces['w'], attr)
    if index not in element.attrib:
        return None
    return element.attrib[index]

# Text inputs are a series of w:r elements at the same level bounded by
# w:bookmarkStart and w:bookmarkEnd with common IDs.
def set_text_input(docxml, name, value):
    start_selector = './/w:bookmarkStart[@w:name="%s"]' % name
    bookmark_start = docxml.find(start_selector, namespaces)
    bookmark_id = get_attr(bookmark_start, 'id')
    bookmark_end = docxml.find(
        './/w:bookmarkEnd[@w:id="%s"]' % (bookmark_id), namespaces)

    current_element = bookmark_start
    while (current_element != bookmark_end):
        print "stepping to next element"
        current_element = current_element.getnext()
        # Find any w:t elements and set their values.
        text_input = current_element.find('.//w:t', namespaces)
        if text_input is not None:
            print "updating current text element"
            text_input.text = value
            break

def find_next_element_matching(
        docxml,
        start_selector,
        check_element_lamdba,
        failure_message='Could not find following node matching condition'):
    label_element = docxml.find(start_selector, namespaces)
    if label_element == None:
        raise Exception('Could not find name element with val=%s' % name)

    current_element = label_element
    while current_element != None and check_element_lamdba(current_element):
        current_element = current_element.getnext()

    if current_element == None:
        raise Exception(failure_message)

    return current_element


# Checkbox elements are w:checkBox element that are preceeded by a w:name with val
# equal to the checkbox name
#
# isChecked depends on the presence of an <w:checked> field inside this checkBox element
def check_checkbox(docxml, name, should_be_checked):
    start_selector = './/w:name[@w:val="%s"]' % name
    checkbox_tag_name = "{%s}checkBox" % namespaces['w']

    checkbox_element = find_next_element_matching(
        docxml,
        start_selector,
        lambda element: element.tag != checkbox_tag_name,
        failure_message="Could not find w:checkBox following a w:name with val=%s" % name
        )

    check_element = checkbox_element.find('w:checked', namespaces);
    if should_be_checked and check_element is None:
        new_check_element = ET.Element('{%s}checked' % namespaces['w'])
        checkbox_element.append(new_check_element)
    elif not should_be_checked and check_element is not None:
        check_element.getparent().remove(check_element)

def levenshtein_distance_nounicode(a, b):
    return levenshtein_distance(a.decode(encoding='UTF-8'), b.decode(encoding='UTF-8'))

def get_best_option_index(options, requested):
    best_option = min(options, key=lambda option: levenshtein_distance_nounicode(requested, option))
    if levenshtein_distance_nounicode(requested, best_option) != 0:
        log_warning(
            "Using best match '%s' for requested option '%s'" % (
                best_option,
                requested,
                )
        )

    return options.index(best_option)

# Dropdown elements are a w:ddList element that are preceeded by a w:name with val
# equal to the dropdown name
#
# the selection depends on the presence of a <w:result w:val="$INDEX"/> field inside this checkBox element
def select_option(docxml, name, option_to_select):
    start_selector = './/w:name[@w:val="%s"]' % name
    dropdownlist_element_name = "{%s}ddList" % namespaces['w']
    val_attribute_name = '{%s}val' % namespaces['w']

    dropdown_element = find_next_element_matching(
        docxml,
        start_selector,
        lambda element: element.tag != dropdownlist_element_name,
        failure_message="Could not find w:ddList following a w:name with val=%s" % name
        )

    dropdown_option_elements = dropdown_element.findall(
        '{%s}listEntry' % namespaces['w']
        )
    dropdown_options = [
        dropdown_option_element.attrib[val_attribute_name]
        for dropdown_option_element
        in dropdown_option_elements
        ]

    index = get_best_option_index(dropdown_options, option_to_select)
    result_element = dropdown_element.find('{%s}result' % namespaces['w'])
    if result_element is None:
        result_element = ET.Element('{%s}result' % namespaces['w'])
        dropdown_element.append(result_element)
    result_element.attrib[val_attribute_name] = str(index)

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

set_text_input(docxml, 'Text414', 'hi')
check_checkbox(docxml, 'Check70', False)
select_option(docxml, 'Dropdown21', 'NCR(s) closed')

xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
xml_body = ET.tostring(docxml)

update_zip_in_place(
    input_zip_file,
    out_name,
    'word/document.xml',
    xml_header + xml_body
)
