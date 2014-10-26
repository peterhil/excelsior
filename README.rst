Excelsior
=========

Excelsior is a tool to convert Excel spreadsheets into TSV, CSV, Json or Yaml.
Reads all sheets of the spreasheets.

Uses `xlrd <http://www.python-excel.org/>`_ for reading the Excel
files, and thus supports the new `Open Office XML file
format <https://en.wikipedia.org/wiki/Office_Open_XML>`_ (.xlsx
extension).

Supported output formats are `tab separated values
(.tsv) <http://www.cs.tut.fi/~jkorpela/TSV.html>`_, `comma separated
values (.csv) <https://en.wikipedia.org/wiki/Comma-separated_values>`_,
Yaml and JSON.

Installation
============

.. code-block:: bash

    $ pip install excelsior

Usage
=====

By default outputs into standard output, and separate sheets are separated by
a `form feed <https://en.wikipedia.org/wiki/Page_break#Form_feed>`_ and new
line characters (``\x0c\n``), followed by a header line of the form ``# Sheet
name #\n``.

When writing onto files with the ``-w`` option, no such characters or headers
are written.

Output TSV:
-----------

.. code-block:: bash

    $ excelsior -f tsv excel.xlsx

Convert into TSV and write to files:
------------------------------------

.. code-block:: bash

    $ excelsior -w -f tsv excel.xlsx another-excel.xls

This will save the output into ``<filename>.tsv``, if the spreasheet has only  
one sheet, or ``<filename>-<sheet>.tsv`` if it has multiple sheets.

You can also pipe in the filenames (separated by newlines):

.. code-block:: bash

    $ echo "ds140-bauxi.xlsx\nds140-alumi.xlsx" | excelsior -w -f tsv
    ds140-bauxi-Bauxite.tsv: written
    ds140-bauxi-Alumina.tsv: written
    ds140-alumi.tsv: written

Show help:
----------

.. code-block:: bash

    $ excelsior -h
