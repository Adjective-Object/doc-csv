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


def find_element_following_bookmark(
        docxml,
        start_selector,
        check_element_lamdba,
        failure_message='Could not find following node matching condition'):
    """
    Finds the first sibling node matching check_element_lambda that follows the
    first node found matching start_selector
    """
    label_element = docxml.find(start_selector, namespaces)
    if label_element is None:
        raise Exception('Could not find name element with val=%s' % start_selector)

    current_element = label_element
    while (
        current_element is not None and
        check_element_lamdba(current_element)
    ):
        current_element = current_element.getnext()

    if current_element is None:
        raise Exception(failure_message)

    return current_element


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


class Field(object):
    def __init__(self, field_name, field_value):
        self.field_name = field_name
        self.field_value = field_value

    def __repr__(self):
        return "%s(%s, %s)" % (
            type(self).__name__,
            self.field_name.__repr__(),
            self.field_value.__repr__()
            )


class TextField(Field):
    def __init__(self, field_name, field_value):
        super(TextField, self).__init__(field_name, field_value)

    def apply(self, docxml):
        TextField.set_text_input(
            docxml,
            self.field_name,
            self.field_value
        )

    @staticmethod
    def set_text_input(docxml, name, value):
        """
        Text inputs are a series of w:r elements at the same level bounded by

        w:bookmarkStart and w:bookmarkEnd with common IDs.
        """
        start_selector = './/w:bookmarkStart[@w:name="%s"]' % name
        bookmark_start = docxml.find(start_selector, namespaces)
        bookmark_id = get_attr(bookmark_start, 'id')
        bookmark_end = docxml.find(
            './/w:bookmarkEnd[@w:id="%s"]' % (bookmark_id), namespaces)

        current_element = bookmark_start
        while (current_element != bookmark_end):
            # print "stepping to next element"
            current_element = current_element.getnext()
            # Find any w:t elements and set their values.
            text_input = current_element.find('.//w:t', namespaces)
            if text_input is not None:
                # print "updating current text element"
                text_input.text = value
                break


class CheckboxField(Field):
    def __init__(self, field_name, field_value):
        super(CheckboxField, self).__init__(
            field_name,
            CheckboxField.transform_value(field_value)
        )

    @staticmethod
    def validate_value(field_value_string):
        return CheckboxField.transform_value(field_value_string) is not None

    @staticmethod
    def transform_value(field_value_string):
        field_value_lower = field_value_string.lower()
        is_field_value_truthy = field_value_lower in [
            'true', '1', 't', 'yes', 'y'
        ]
        is_field_value_fasly = field_value_lower in [
            'false', '0', 'f', 'no', 'n'
        ]
        if (not is_field_value_truthy) and (not is_field_value_fasly):
            return None

        return is_field_value_truthy

    def apply(self, docxml):
        CheckboxField.check_checkbox(
            docxml,
            self.field_name,
            self.field_value
        )

    @staticmethod
    def check_checkbox(
        docxml,
        name,
        should_be_checked
    ):
        """
        Checkbox elements are w:checkBox element that are preceeded by a w:name
        with val equal to the checkbox name

        isChecked depends on the presence of an <w:checked> field inside this
        checkBox element
        """
        start_selector = './/w:name[@w:val="%s"]' % name
        checkbox_tag_name = "{%s}checkBox" % namespaces['w']

        checkbox_element = find_element_following_bookmark(
            docxml,
            start_selector,
            lambda element: element.tag != checkbox_tag_name,
            failure_message=(
                "Could not find w:checkBox following a w:name with val=%s" % name
            )
        )

        check_element = checkbox_element.find('w:checked', namespaces)
        if should_be_checked and check_element is None:
            new_check_element = ET.Element('{%s}checked' % namespaces['w'])
            checkbox_element.append(new_check_element)
        elif not should_be_checked and check_element is not None:
            check_element.getparent().remove(check_element)


class DropdownField(Field):
    def __init__(self, field_name, field_value):
        super(DropdownField, self).__init__(field_name, field_value)

    def apply(self, docxml):
        DropdownField.select_option(docxml, self.field_name, self.field_value)

    @staticmethod
    def levenshtein_distance_nounicode(a, b):
        return levenshtein_distance(
            a.decode(encoding='UTF-8'),
            b.decode(encoding='UTF-8')
        )

    @staticmethod
    def get_best_option_index(options, requested):
        best_option = min(
            options,
            key=lambda option: DropdownField.levenshtein_distance_nounicode(
                requested, option)
        )

        distance_to_best_option = DropdownField.levenshtein_distance_nounicode(
            requested,
            best_option
            )

        if distance_to_best_option != 0:
            log_warning(
                "Using best match '%s' for requested option '%s'" % (
                    best_option,
                    requested,
                )
            )

        return options.index(best_option)

    @staticmethod
    def select_option(docxml, name, option_to_select):
        """
        Dropdown elements are a w:ddList element that are preceeded by a w:name
        with val equal to the dropdown name

        the selection depends on the presence of a <w:result w:val="$INDEX"/>
        field inside this checkBox element
        """
        start_selector = './/w:name[@w:val="%s"]' % name
        dropdownlist_element_name = "{%s}ddList" % namespaces['w']
        val_attribute_name = '{%s}val' % namespaces['w']

        dropdown_element = find_element_following_bookmark(
            docxml,
            start_selector,
            lambda element: element.tag != dropdownlist_element_name,
            failure_message=(
                "Could not find w:ddList following a w:name with val=%s" % name
            )
        )

        dropdown_option_elements = dropdown_element.findall(
            '{%s}listEntry' % namespaces['w']
        )
        dropdown_options = [
            dropdown_option_element.attrib[val_attribute_name]
            for dropdown_option_element
            in dropdown_option_elements
        ]

        index = DropdownField.get_best_option_index(
            dropdown_options, option_to_select)
        result_element = dropdown_element.find('{%s}result' % namespaces['w'])
        if result_element is None:
            result_element = ET.Element('{%s}result' % namespaces['w'])
            dropdown_element.append(result_element)
        result_element.attrib[val_attribute_name] = str(index)


def set_document_fields(in_name, out_name, fields):
    # in_name = 'report_template.docx'
    # out_name = 'report_template_out.docx'

    input_zip_file = zipfile.ZipFile(in_name, 'r')
    xml_text = input_zip_file.read('word/document.xml')

    docx_namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    ET.register_namespace('w', docx_namespace)
    docxml = ET.fromstring(xml_text)

    for field in fields:
        field.apply(docxml)

    xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    xml_body = ET.tostring(docxml)

    update_zip_in_place(
        input_zip_file,
        out_name,
        'word/document.xml',
        xml_header + xml_body
    )
