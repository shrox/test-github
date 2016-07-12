import base64
import shutil
import sys

from lxml import etree
from copy import deepcopy
from zipfile import ZipFile
from StringIO import StringIO


def mimetype(zip_file):
    # TODO Change according to type of open document
    zip_file.writestr("mimetype", "application/vnd.oasis.opendocument.text")


def parse_fodt(filename):
    filename = open(filename, "r")
    fodt_tree = etree.parse(filename)
    fodt_root = fodt_tree.getroot()
    fodt_namespaces = fodt_root.nsmap
    return (fodt_root, fodt_namespaces)


def write_split_to_zip(zip_file, files_dictionary):
    for filename in files_dictionary:
        zip_file.writestr(filename, files_dictionary[filename])


def split_file(fodt_root, fodt_namespaces, manifest):
    tag_dict = {
        'meta': ['meta'],
        'settings': ['settings'],
        'settings': ['settings'],
        'scripts': ['content'],
        'font-face-decls': ['content', 'styles'],
        'styles': ['styles'],
        'automatic-styles': ['content', 'styles'],
        'master-styles': ['styles'],
        'body': ['content']
    }


    documents_processed = {}
    split_files = {}

    for child in fodt_root:
        tag = child.tag.split('}')[1]

        for xml_filename in tag_dict[tag]:
            document = documents_processed.get(xml_filename)

            if document is None:
                document = etree.Element(
                    ('{%s}document-%s' %
                     (fodt_namespaces['office'], xml_filename)),
                    nsmap=fodt_namespaces)
                documents_processed[xml_filename] = document

            document.append(deepcopy(child))
            document_string = etree.tostring(
                document, encoding='UTF-8', xml_declaration=True, pretty_print=True)


            split_files[xml_filename + '.xml'] = document_string
            # zip_file.writestr("%s.xml" % (xml_filename), document_string)

            # Write to manifest object
            # manifest.add_manifest_entry("%s.xml" % (xml_filename))
            manifest.add_manifest_entry("%s.xml" % (xml_filename))

    return split_files




