#!/usr/bin/env python2
import codecs
import csv
from field_values import CheckboxField, DropdownField, TextField


def validation_warning(error_string):
    print("VALIDATION WARNING: %s" % error_string)


def validate_csv_row(row_dict):
    expected_columns = ['field_type', 'field_name', 'field_value']
    missing_expected_columns = filter(
        lambda column: column not in row_dict.keys(),
        expected_columns,
        )
    if len(missing_expected_columns) != 0:
        validation_warning(
            "csv row %s is missing required keys %s" % (
                row_dict,
                missing_expected_columns
                )
            )
        return False

    expected_field_types = [
        'text', 'dropdown', 'checkbox'
    ]
    field_type = row_dict['field_type'].lower()
    if field_type not in expected_field_types:
        validation_warning(
            "field_type %s is not one of exected values %s" % (
                row_dict,
                expected_field_types
                )
            )
        return False

    return True


def get_fields_from_file(csvfile):
    fields = []
    reader = csv.DictReader(csvfile)
    for row in reader:
        if not validate_csv_row(row):
            continue

        field_type = row["field_type"].lower()
        field_name = row['field_name']
        field_value = row['field_value']

        if field_type == "text":
            fields.append(TextField(field_name, field_value))

        elif field_type == "dropdown":
            fields.append(DropdownField(field_name, field_value))

        elif field_type == "checkbox":
            if not CheckboxField.validate_value(field_value):
                validation_warning(
                    'Invalid value "%s" for checkbox field "%s"' % (
                        row['field_value'],
                        row['field_name'],
                    ))
                continue
            fields.append(CheckboxField(field_name, field_value))
        else:
            print('Unexpected field_type "%s" encountered in row %s' % (
                field_type,
                row
                )
            )
            exit(1)
    return fields


def get_fields(csvfilename):
    with codecs.open(csvfilename, 'rb', encoding='ascii', errors='ignore') as csvfile:
        return get_fields_from_file(csvfile)

if __name__ == "__main__":
    print get_fields("data.csv")
