#!/usr/bin/env python
# -*- coding: utf-8 mode: python -*-
#
# Copyright (c) 2014, Peter Hillerström <peter.hillerstrom@gmail.com>
# All rights reserved. This software is licensed under 3-clause BSD license.
#
# For the full copyright and license information, please view the LICENSE
# file that was distributed with this source code.

import docopt
import json
import os
import re
import sys
import yaml
import xlrd

from mmap import mmap, ACCESS_READ


supported_formats=['csv', 'tsv', 'yaml', 'json']


if sys.version_info >= (3, 0, 0):
    unicode = str


def cell_value(cell):
    ctype, value = repr(cell).split(':', 1)
    if 'empty' == ctype:
        return u''
    elif 'text' == ctype:
        return unicode(value[1:-1]) if sys.version_info >= (3, 0, 0) else unicode(value[2:-1])
    elif 'number' == ctype:
        if float(value) % 1 == 0:
            return int(float(value))
        else:
            return float(value)
    elif 'xldate' == ctype:
        return
    elif 'bool' == ctype:
        return bool(value)
    elif 'error' == ctype:
        return u'ERR: {}'.format(value)
    elif 'blank' == ctype:
        return None
    else:
        return unicode(value)


def read_data(xls):
    data = {}
    data['sheets'] = []

    for sheet in xls.sheets():
        table = {'name': sheet.name, 'data': []}
        for row_nr in range(sheet.nrows):
            row = []
            for col_nr in range(sheet.ncols):
                row.append(cell_value(sheet.cell(row_nr, col_nr)))
            table['data'].append(row)
        data['sheets'].append(table)
    return data


def convert_format(sheet, fmt='tsv'):
    if 'yaml' == fmt:
        return yaml.dump(data)
    elif 'json' == fmt:
        return json.dumps(data, indent=4)
    elif 'tsv' == fmt:
        return u'\n'.join([
            '\t'.join(map(unicode, row))
            for row in sheet['data']
            ])
    elif 'csv' == fmt:
        return u'\n'.join([
            ', '.join(map(lambda s: '"' + unicode(s).replace('"', '""') + '"', row))
            for row in sheet['data']
            ])
    else:
        err_message = 'Unsupported format: {}'.format(fmt)
        raise ValueError(err_message)


def colour_error(message):
    red = '\x1b[91m'
    normal = '\x1b[0m'
    return red + message + normal


def write_file(data, outfile):
    with open(outfile, 'wb') as f:
        f.write(data)
        sys.stderr.write('{}: written'.format(outfile))


def convert(excel_file, fmt='tsv', output='print'):
    try:
        with open(excel_file, 'rb') as f:
            xls = xlrd.open_workbook(
                file_contents=mmap(f.fileno(), 0, access=ACCESS_READ)
                )
        # xls = xlrd.open_workbook(excel_file)
        data = read_data(xls)
        path = os.path.splitext(excel_file)[0]

        for idx, sheet in enumerate(data['sheets']):
            if output == 'print':
                if len(data['sheets']) > 1:
                    if idx != 0:
                        sys.stdout.write(u'\x0c\n')
                    sys.stdout.write(u'# {} #\n'.format(sheet['name']))
                sys.stdout.write(convert_format(sheet, fmt=fmt))
            else:
                if len(data['sheets']) > 1:
                    outfile = path + '-' + sheet['name'] + '.' + fmt
                else:
                    outfile = path + '.' + fmt
                write_file(convert_format(sheet, fmt=fmt), outfile)
            sys.stdout.write('\n')

    except xlrd.biffh.XLRDError as err:
        sys.stderr.write(excel_file + ': ' + colour_error('Error: ' + err.message))
    except (IOError, ValueError) as err:
        sys.stderr.write(colour_error('Error: ') + str(err) + '\n')
        sys.exit(2)


__doc__ = """{command}

Usage:
    {command} [-f fmt|--format=fmt] [-w|--write]
    {command} [-f fmt|--format=fmt] [-w|--write] <excels>...
    {command} -h | --help

Arguments:
    <excels>  Excel spreadsheet(s) to be converted

Options:
    -f fmt|--format=fmt  Output file format, one of {supported_formats}
    -w|--write   Write to files
    --caps  convert the text to upper case
""".format(
    command=os.path.basename(sys.argv[0]),
    supported_formats=supported_formats
    )


def main():
    try:
        args = docopt.docopt(__doc__)
    except docopt.DocoptExit as err:
        print(err)
        sys.exit(1)

    if args['--help']:
        print(__doc__)
        sys.exit(0)

    output = 'files' if args['--write'] or args['-w'] else 'print'

    fmt = args['--format'] or args['-f']
    if fmt and not fmt in supported_formats:
        print('Unsupported format: {}'.format(fmt))
        sys.exit(1)

    for excel_file in args['<excels>'] or map(lambda s: s.strip(), sys.stdin.readlines()):
        convert(excel_file, fmt, output)


if __name__ == '__main__':
    main()
