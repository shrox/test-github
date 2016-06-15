import os
import base64

from lxml import etree
from copy import deepcopy
from StringIO import StringIO

# Global Variables
fodt_root = []
fodt_namespaces = {}

def make_folders():
    os.mkdir("Pictures")
    os.mkdir("Thumbnails")
    os.mkdir("META-INF")


class Binary2Image():
    '''This class will convert the binary data
                        to an image'''

    
    def __init__(self, binary_data, image_name):
        self.binary_data = binary_data
        self.image_name = image_name

    def convert2image(self):
        image = base64.b64decode(self.binary_data)

        with open(self.image_name, 'w+') as img:
            img.write(image)


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
        documents_processed = {}
        for child in fodt_root:
            tag = child.tag.split('}')[1]

            for xml_filename in self.tag_dict[tag]:
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

    def handle_images(self):
        image_number = 0

        try:
            while True:
                binary_data = self.content_root.xpath(
                        "//draw:image/office:binary-data/text()", 
                                     namespaces=fodt_namespaces)[image_number]

                image_name = "Pictures/image" + str(image_number) + ".jpg"

                Binary2Image(binary_data, image_name).convert2image()


                node = self.content_root.xpath(
                        "//draw:image", namespaces=fodt_namespaces)[image_number]
                node.attrib["{" + fodt_namespaces['xlink'] + "}" + "href"] = image_name
                node.attrib["{" + fodt_namespaces['xlink'] + "}" + "simple"] = "simple"
                node.attrib["{" + fodt_namespaces['xlink'] + "}" + "show"] = "embed"
                node.attrib["{" + fodt_namespaces['xlink'] + "}" + "actuate"] = "onLoad"

                image_number = image_number + 1

        except IndexError:
            pass

        # Delete  binary data
        for elem in self.content_root.xpath(
                "//draw:image/office:binary-data", namespaces=fodt_root.nsmap):
                elem.getparent().remove(elem)

        # Write to disk
        content_file = open("content.xml", "w")
        content_file.write(etree.tostring(
                  self.content_root, encoding='UTF-8', 
                                                xml_declaration=True))

        # def delete_binary_data(self):
        #     print "HandleImages delete_binary_data"
        #     for elem in self.content_root.xpath(
        #         "//draw:image/office:binary-data", namespaces=fodt_root.nsmap):
        #         elem.getparent().remove(elem)

        # def write_to_disk():
        #     print "HandleImages write_to_disk"
        #     content_file = open("content.xml", "w")
        #     content_file.write(etree.tostring(
        #           self.content_root, encoding='UTF-8', 
        #                                         xml_declaration=True))


def main():
    make_folders()

    fodt_filename = raw_input("Enter name of FODT file: ")
    OpenParseFodt(fodt_filename).parse()
    FileSplit(fodt_filename).split()
    HandleImages().handle_images()


if __name__ == "__main__":main()