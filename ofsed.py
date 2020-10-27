
import re
import sys
import zipfile
import argparse
import typing as t

import xml.dom.minidom as minidom

from pathlib import Path

sys.path.append('python_coreutils')

from python_coreutils.coreutils.sed import sed_substitute

def get_contents(path_to_doc: Path, docx: bool) -> minidom.Document:
    with zipfile.ZipFile(str(path_to_doc), 'r') as doc:
        content_file_path = 'word/document.xml' if docx else 'content.xml'
        with doc.open(content_file_path) as content:
            xml = content.read()
            dom = minidom.parseString(xml)
            dom.normalize()
            return dom


def set_contents(path_to_doc: Path, docx: bool, contents: t.Iterable[str]):
    with zipfile.ZipFile(str(path_to_doc), 'r') as doc:
        content_file_path = 'word/document.xml' if docx else 'content.xml'
        with doc.open(content_file_path) as content:
            xml = content.writelines(contents)


def sed_content(cmd: str,
                path_to_doc: Path,
                docx: bool,
                raw: bool = False):
    dom = get_contents(path_to_doc, docx)

    lines_to_substitute = []
    if docx:
        for text in dom.getElementsByTagName('w:t'):
            for line in text.childNodes:
                lines_to_substitute.append((line, line.nodeValue))
    else:
        for text in dom.getElementsByTagName('text:p'):
            for span in text.childNodes:
                if span.nodeValue:
                    lines_to_substitute.append((span, span.nodeValue))
                for line in span.childNodes:
                    lines_to_substitute.append((line, line.nodeValue))

    lines_to_substitute = list(filter(lambda t: t[1], lines_to_substitute))
    nodes = list(map(lambda t: t[0], lines_to_substitute))
    lines = list(map(lambda t: t[1], lines_to_substitute))

    substituted_lines = sed_substitute(cmd, lines)
    for i, subs_line in enumerate(substituted_lines):
        nodes[i].nodeValue = subs_line

    set_contents(path_to_doc, docx, dom.toxml())


parser = argparse.ArgumentParser(description='Program to "sed" .docx and .odt files')
parser.add_argument('-i', '--inplace', action='store_true', help='change files inplace')
parser.add_argument('-r', '--raw', action='store_true', help='use ofsed on raw xml files')
parser.add_argument('CMD', help='command to execute')
parser.add_argument('PATH', nargs='+', help='path to document file')

args = parser.parse_args()

if not hasattr(args, 'inplace') or not args.inplace:
    print('This program changes file contents inplace!')
    print('Be careful with that! If you want this nag dissapear on start')
    print('launch program with "-i" flag')
    input('Are you conscious about that? [Press <Enter> or abort]')
    input('Last chance to turn back [Press <Enter> or abort]')

for path in args.PATH:
    path_to_doc = Path(path)
    extension = path_to_doc.suffix[1:]
    
    sed_content(args.CMD,
                path_to_doc,
                extension == 'docx',
                args.raw)
