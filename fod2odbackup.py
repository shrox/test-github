import os
import sys
import shutil
import base64
import mimetypes
import magic

from lxml import etree
from copy import deepcopy
from zipfile import ZipFile
from StringIO import StringIO


def mimetype(zip_file):
    # TODO Change according to type of open document
    zip_file.writestr("mimetype", "application/vnd.oasis.opendocument.text")


def parse_fod(filename):
    filename = open(filename, "r")
    fod_tree = etree.parse(filename)
    fod_root = fod_tree.getroot()
    fod_namespaces = fod_root.nsmap
    return (fod_root, fod_namespaces)


def decode_images_to_zip(zip_file, document, fod_namespaces, manifest):
    ''' 
        Args:
            param zip_file: zip file to write images to
            param document: document input, with or without images
            param fod_namespaces: all namespaces in the fod document
            param manifest: manifest instance to write to manifest.xml
    '''

    image_number = 0

    all_binary_data = document.xpath(
        "//draw:image/office:binary-data/text()",
        namespaces={'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', 'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'})

    with magic.Magic(flags=magic.MAGIC_MIME_TYPE) as m:
        for binary_data in all_binary_data:
            # Decode image using base64 module
            image = base64.b64decode(binary_data)
            
            # Identify mime to identify extension 
            mime = m.id_buffer(image)

            image_name = "Pictures/image%s%s" % (image_number, mimetypes.guess_extension(mime))

            zip_file.writestr(image_name, image)

            node = document.xpath(
                "//draw:image", namespaces=fod_namespaces)[image_number]
            node.attrib["{%s}href" % (fod_namespaces['xlink'])] = image_name
            node.attrib["{%s}simple" % (fod_namespaces['xlink'])] = "simple"
            node.attrib["{%s}show" % (fod_namespaces['xlink'])] = "embed"
            node.attrib["{%s}actuate" % (fod_namespaces['xlink'])] = "onLoad"

            image_number += 1

            # Delete binary data
            elem = binary_data.getparent()
            elem.getparent().remove(elem)

            # Write to manifest object
            manifest.add_manifest_entry(image_name)


def split_file_to_zip(zip_file, fod_root, fod_namespaces, manifest):
    ''' 
        Args:
            param zip_file: zip file to write images to
            param manifest: to write file locations to manifest.xml

        FOD will be split to smaller files in accordance to their tags and written to zip
        
    '''

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
                     (fod_namespaces['office'], 'meta')),
                    nsmap=fod_namespaces),
        'settings': etree.Element(
                    ('{%s}document-%s' %
                     (fod_namespaces['office'], 'settings')),
                    nsmap=fod_namespaces),
        'content': etree.Element(
                    ('{%s}document-%s' %
                     (fod_namespaces['office'], 'content')),
                    nsmap=fod_namespaces),
        'styles': etree.Element(
                    ('{%s}document-%s' %
                     (fod_namespaces['office'], 'styles')),
                    nsmap=fod_namespaces)
    }

    for child in fod_root:
        tag = etree.QName(child).localname
        for xml_filename in tag2file[tag]:
            document = documents_processed[xml_filename]

            document.append(deepcopy(child))
            
            # Specified document ends only with one of the following tags
            if tag in ['meta', 'settings', 'master-styles', 'body']:
                decode_images_to_zip(
                    zip_file, document, fod_namespaces, manifest)

                document_string = etree.tostring(
                     document, encoding='UTF-8', xml_declaration=True)

                zip_file.writestr("%s.xml" % (xml_filename), document_string)

                # Write to manifest object
                manifest.add_manifest_entry("%s.xml" % (xml_filename))


class Manifest(object):

    ''' Class to handle manifest.xml in META-INF folder'''

    def __init__(self, fod_root, fod_namespaces):
        self.manifest_namespace = {
            "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"}

        self.document = etree.Element(
            ("{%s}manifest" % (self.manifest_namespace["manifest"])),
            nsmap=self.manifest_namespace)

        self.document.attrib["{%s}version" %
                             (self.manifest_namespace["manifest"])] = fod_root.xpath("//@office:version", namespaces=fod_namespaces)[0]

        # Will be in every manifest.xml
        # manifest_entry = etree.SubElement(self.document,
        #                                   "{%s}file-entry" % (self.manifest_namespace["manifest"]))
        # manifest_entry.attrib["{%s}full-path" %
        #                       (self.manifest_namespace["manifest"])] = "/"
        # manifest_entry.attrib["{%s}media-type" % 
        #                         (self.manifest_namespace["manifest"])] = "application/vnd.oasis.opendocument.text"

        self.add_manifest_entry('/')

    def add_manifest_entry(self, file_path):
        manifest_entry = etree.SubElement(self.document, "{%s}file-entry" % 
                                            (self.manifest_namespace["manifest"]))

        manifest_entry.attrib["{%s}full-path" %
                              (self.manifest_namespace["manifest"])] = file_path

        file_name = os.path.basename(file_path)
        # If condition for when filename is '/'
        if file_name == '':
            manifest_entry.attrib[
            "{%s}media-type" % (self.manifest_namespace["manifest"])] = 'application/vnd.oasis.opendocument.text'
        else:
            manifest_entry.attrib[
                "{%s}media-type" % (self.manifest_namespace["manifest"])] = mimetypes.guess_type(file_name)[0]

    def write_to_zip(self, zip_file, manifest): 
        manifest_string = etree.tostring(
            manifest.document, encoding='UTF-8', xml_declaration=True, pretty_print=True)
        zip_file.writestr("META-INF/manifest.xml", manifest_string)


def convert(filename):
    fod_root, fod_namespaces = parse_fod(filename)
    output_odt = StringIO()
    zip_file = ZipFile(output_odt, "w")

    mimetype(zip_file)

    # Can add option to have something else as odt_filename
    odt_filename = filename

    manifest = Manifest(fod_root, fod_namespaces)
    split_file_to_zip(
        zip_file, fod_root, fod_namespaces, manifest)
    manifest.write_to_zip(zip_file, manifest)

    zip_file.close()
    output_odt.seek(0)

    with open("%s.odt" % (odt_filename.split(".")[0]), "wb") as odt:
        shutil.copyfileobj(output_odt, odt)


if __name__ == "__main__":
    convert(sys.argv[1])