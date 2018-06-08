#!/usr/bin/env python2
import getopt
import sys
import os
import time
from csv_parse import get_fields
from field_values import set_document_fields

long_opts = [
    'help',
    'output='
]

short_opts = [
    opt[0] + (':' if opt.endswith('=') else '') for opt in long_opts
]


def convert_short_to_long_opt(key, value):
    if len(key) != 2:  # Keys of form '-k'
        return (key, value)
    else:
        key = filter(
            lambda long_option_name: long_option_name[0] == key[1],
            long_opts
        )[0]
        if key.endswith('='):
            key = key[:-1]
        return '--'+key, value


def get_opt_dict(argv):
    optlist, args = getopt.getopt(argv, ''.join(short_opts), long_opts)
    optlist = [convert_short_to_long_opt(k, v) for k, v in optlist]
    return dict(optlist), args


def get_default_output_docx():
    return "output_%s.docx" % time.strftime(
        "%Y-%m-%d.%H.%M.%S"
    )

def print_help():
    print "usage: %s <template_docx> <csv_file>" % sys.argv[0]
    print """
    options:
        -h/--help   : display this helptext
        -o/--output : set the output path (defaults to output_<current_time>.dox)
    """


if __name__ == "__main__":
    opts, args = get_opt_dict(sys.argv[1:])
    if '--help' in opts.keys():
        print_help()
        exit(0)

    if len(args) != 2:
        print_help()
        exit()

    docx_template_path = args[0]
    output_docx_path = (
        opts['--output']
        if '--output' in opts.keys()
        else get_default_output_docx()
        )
    csv_data_path = args[1]

    if not os.path.isfile(docx_template_path):
        print "Template file %s does not exist" % docx_template_path
        print_help()
        exit(1)

    if not os.path.isfile(csv_data_path):
        print "CSV file %s does not exist" % csv_data_path
        print_help()
        exit(1)

    print "templating %s with data from %s to %s" % (
        docx_template_path,
        csv_data_path,
        output_docx_path
        )

    fields = get_fields(csv_data_path)
    set_document_fields(docx_template_path, output_docx_path, fields)
