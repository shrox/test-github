import os
import sys
import shutil
import base64
import mimetypes

from lxml import etree
from magic import Magic
from magic import MAGIC_MIME_TYPE
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


def decode_images_to_zip(zip_file, document, fodt_namespaces, manifest):
    ''' Will decode images and write them to zip '''

    image_number = 0

    all_binary_data = document.xpath(
        "//draw:image/office:binary-data/text()",
        namespaces=fodt_namespaces)

    for binary_data in all_binary_data:
        # Decode image using base64 module
        image = base64.b64decode(binary_data)
        
        # Identify mime to identify extension 
        with Magic(flags=MAGIC_MIME_TYPE) as m:
            mime = m.id_buffer(image)

        image_name = "Pictures/image%s%s" % (str(image_number), mimetypes.guess_extension(mime))

        zip_file.writestr(image_name, image)

        node = document.xpath(
            "//draw:image", namespaces=fodt_namespaces)[image_number]
        node.attrib["{%s}href" % (fodt_namespaces['xlink'])] = image_name
        node.attrib["{%s}simple" % (fodt_namespaces['xlink'])] = "simple"
        node.attrib["{%s}show" % (fodt_namespaces['xlink'])] = "embed"
        node.attrib["{%s}actuate" % (fodt_namespaces['xlink'])] = "onLoad"

        image_number += 1

        # Delete binary data
        elem = binary_data.getparent()
        elem.getparent().remove(elem)

        # Write to manifest object
        manifest.add_manifest_entry(image_name)

    document_string = etree.tostring(
        document, encoding='UTF-8', xml_declaration=True)

    return document_string


def split_file_to_zip(zip_file, fodt_root, fodt_namespaces, manifest):
    ''' Splits files and writes them to zip'''

    tag2file = {
        'meta': ['meta'],
        'settings': ['settings'],
        'scripts': ['content'],
        'font-face-decls': ['content', 'styles'],
        'styles': ['styles'],
        'automatic-styles': ['content', 'styles'],
        'master-styles': ['styles'],
        'body': ['content']
    }

    documents_processed = {
        'meta': etree.Element(
                    ('{%s}document-%s' %
                     (fodt_namespaces['office'], 'meta')),
                    nsmap=fodt_namespaces),
        'settings': etree.Element(
                    ('{%s}document-%s' %
                     (fodt_namespaces['office'], 'settings')),
                    nsmap=fodt_namespaces),
        'content': etree.Element(
                    ('{%s}document-%s' %
                     (fodt_namespaces['office'], 'content')),
                    nsmap=fodt_namespaces),
        'styles': etree.Element(
                    ('{%s}document-%s' %
                     (fodt_namespaces['office'], 'styles')),
                    nsmap=fodt_namespaces)
    }

    for child in fodt_root:
        tag = etree.QName(child).localname
        for xml_filename in tag2file[tag]:
            document = documents_processed.get(xml_filename)

            document.append(deepcopy(child))
            
            # Specified document ends only with one of the following tags
            if tag in ['meta', 'settings', 'master-styles', 'body']:
                # Convert to string while checking for images
                document_string = decode_images_to_zip(
                    zip_file, document, fodt_namespaces, manifest)

                zip_file.writestr("%s.xml" % (xml_filename), document_string)

                # Write to manifest object
                manifest.add_manifest_entry("%s.xml" % (xml_filename))


class Manifest(object):

    ''' Class to handle manifest.xml in META-INF folder'''

    def __init__(self, fodt_root, fodt_namespaces):
        self.manifest_namespace = {
            "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"}

        self.document = etree.Element(
            ("{%s}manifest" % (self.manifest_namespace["manifest"])),
            nsmap=self.manifest_namespace)
        self.document.attrib["{%s}version" %
                             (self.manifest_namespace["manifest"])] = fodt_root.xpath("//@office:version", namespaces=fodt_namespaces)[0]

        # Will be in every manifest.xml
        manifest_entry = etree.SubElement(self.document,
                                          "{%s}file-entry" % (self.manifest_namespace["manifest"]))
        manifest_entry.attrib["{%s}full-path" %
                              (self.manifest_namespace["manifest"])] = "/"
        manifest_entry.attrib["{%s}version" %
                              (self.manifest_namespace["manifest"])] = "1.2"
        manifest_entry.attrib["{%s}media-type" % 
                                (self.manifest_namespace["manifest"])] = "application/vnd.oasis.opendocument.text"

    def add_manifest_entry(self, file_path):
        manifest_entry = etree.SubElement(self.document, "{%s}file-entry" % 
                                            (self.manifest_namespace["manifest"]))

        manifest_entry.attrib["{%s}full-path" %
                              (self.manifest_namespace["manifest"])] = file_path

        file_name = os.path.split(file_path)[1]

        manifest_entry.attrib[
            "{%s}media-type" % (self.manifest_namespace["manifest"])] = mimetypes.guess_type(file_name)[0]

    def write_to_zip(self, zip_file, manifest): 
        manifest_string = etree.tostring(
            manifest.document, encoding='UTF-8', xml_declaration=True, pretty_print=True)
        zip_file.writestr("META-INF/manifest.xml", manifest_string)


def convert(filename):
    fodt_root, fodt_namespaces = parse_fodt(filename)
    output_odt = StringIO()
    zip_file = ZipFile(output_odt, "w")

    mimetype(zip_file)

    # Can add option to have something else as odt_filename
    odt_filename = filename

    manifest = Manifest(fodt_root, fodt_namespaces)
    split_file_to_zip(
        zip_file, fodt_root, fodt_namespaces, manifest)
    manifest.write_to_zip(zip_file, manifest)

    zip_file.close()
    output_odt.seek(0)

    with open("%s.odt" % (odt_filename.split(".")[0]), "wb") as odt:
        shutil.copyfileobj(output_odt, odt)


if __name__ == "__main__":
    convert(sys.argv[1])
