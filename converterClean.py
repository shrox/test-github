import os

from lxml import etree
from binary2image import Binary2Image
from copy import deepcopy
from StringIO import StringIO

# Global Variables
fodt_root = []
fodt_namespaces = []
documents_processed = {}

def make_folders():
    os.mkdir("Pictures")


class OpenParseFodt():
    "A class to open and parse an FODT file"

    def __init__(self, filename):
        self.filename = open(filename, "r")

    def parse(self):
        global fodt_root
        global fodt_namespaces

        fodt_tree = etree.parse(self.filename)
        fodt_root = fodt_tree.getroot()
        fodt_namespaces = fodt_root.nsmap


class FileSplit():
    "A class that splits the FODT into the different XML files"

    def __init__(self, filename):
        self.tag_dict = {}
        self.tag_dict['meta'] = ['meta']
        self.tag_dict['settings'] = ['settings']
        self.tag_dict['scripts'] = ['content']
        self.tag_dict['font-face-decls'] = ['content', 'styles']
        self.tag_dict['styles'] = ['styles']
        self.tag_dict['automatic-styles'] = ['content', 'styles']
        self.tag_dict['master-styles'] = ['styles']
        self.tag_dict['body'] = ['content']

    def split(self):
        global documents_processed

        for child in fodt_root:
            tag = child.tag.split('}')[1]

            for xml_filename in tag_dict[tag]:
                document = documents_processed.get(xml_filename)

                if document is None:
                    document = etree.Element(
                        ('{'+fodt_namespaces['office']+'}' + 'document-' + xml_filename),
                                         nsmap=fodt_namespaces)
                    documents_processed[xml_filename] = document

                document.append(deepcopy(child))
                document_string = etree.tostring(
                    document, encoding='UTF-8', xml_declaration=True)
                xml_file = open(xml_filename + ".xml", "w+")
                xml_file.write(document_string)
                xml_file.close()


class HandleImages():
    """Class used to to handle images : convert to images, 
                delete binary content and link to images"""

    def __init__(self):
        content_file = open("content.xml", "r")
        self.content_string = content_file.read()
        self.content_root = etree.fromstring(self.content_string)
        content_file.close()

    def convert_binary_link(self):
        image_number = 0

        try:
            while True:
                binary_data = self.content_root.xpath(
                        "//draw:image/office:binary-data/text()", 
                                     namespaces=fodt_namespaces)[image_number]

                image_name = "Pictures/image" + str(image_number) + ".jpg"
                Binary2Image(binary_data, image_name).convert2image()

                node = self.content_root.xpath(
                           "//draw:image", namespaces=fodt_namespaces)[x]
                node.attrib[fodt_namespaces['office'] + "href"] = image_name
                node.attrib[fodt_namespaces['office'] + "simple"] = "simple"
                node.attrib[fodt_namespaces['office'] + "show"] = "embed"
                node.attrib[fodt_namespaces['office'] + "actuate"] = "onLoad"


                image_number = image_number + 1

        except IndexError:
            pass

        def delete_binary_data(self):
            for elem in self.content_root.xpath(
                "//draw:image/office:binary-data", namespaces=fodt_root.nsmap):
                elem.getparent().remove(elem)

        def write_to_disk():
            content_file = open("content.xml", "w")
            content_file.write(etree.tostring(
                  documents_processed["body"], encoding='UTF-8', 
                                                xml_declaration=True))


    





        
        