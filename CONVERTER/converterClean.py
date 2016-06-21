import os
import base64
import shutil

from lxml import etree
from copy import deepcopy
from zipfile import ZipFile
from StringIO import StringIO

# Working with creating a memory zip 
output_odt = StringIO()
zip_file = ZipFile(output_odt, "w")

# Global Variables
fodt_root = []
fodt_namespaces = {}

# def make_folders():  #function will probably go
#     os.mkdir("Pictures")
#     os.mkdir("Thumbnails")
#     os.mkdir("META-INF")

def mimetype():
    global zip_file
    zip_file.writestr("mimetype", "application/vnd.oasis.opendocument.text")

    # mimetype = open("mimetype", "w")
    # mimetype.write("application/vnd.oasis.opendocument.text")
    # mimetype.close()

class Binary2Image():
    '''This class will convert the binary data
                        to an image'''

    
    def __init__(self, binary_data, image_name):
        self.binary_data = binary_data
        self.image_name = image_name

    def convert2image(self):
        global zip_file
        image = base64.b64decode(self.binary_data)
        zip_file.writestr(self.image_name, image)

        # with open(self.image_name, 'w+') as img:
            # img.write(image)


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
        global zip_file
        documents_processed = {}
        for child in fodt_root:
            tag = child.tag.split('}')[1]

            for xml_filename in self.tag_dict[tag]:
                document = documents_processed.get(xml_filename)

                if document is None:
                    document = etree.Element(
                        ('{' + fodt_namespaces['office'] + '}' + 'document-' + xml_filename),
                                nsmap=fodt_namespaces)
                    documents_processed[xml_filename] = document

                document.append(deepcopy(child))
                document_string = etree.tostring(
                    document, encoding='UTF-8', xml_declaration=True)

                zip_file.writestr(xml_filename + ".xml", document_string)

                # xml_file = open(xml_filename + ".xml", "w+")
                # xml_file.write(document_string)
                # xml_file.close()


class HandleImages():
    """Class used to to handle images : convert to images, 
                delete binary content and link to images"""

    def __init__(self):
        global zip_file
        self.content_string = zip_file.read("content.xml")
        # content_file = open("content.xml", "r")
        # self.content_string = content_file.read()
        self.content_root = etree.fromstring(self.content_string)
        # content_file.close()

    def handle_images(self):
        global zip_file
        image_number = 0

        # possible bad code, improve
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
        content_newstring = etree.tostring(
                  self.content_root, encoding='UTF-8', 
                                                xml_declaration=True)

        zip_file.writestr("content.xml", content_newstring)
        # content_file = open("content.xml", "w")
        # content_file.write(content_newstring)
        # content_file.close()


class Manifest():
    ''' Class to create manifest.xml in META-INF folder '''
    
    def __init__(self):
        # Might need to add more possible extensions
        self.extension_dict = {}
        self.extension_dict['xml'] = 'text/xml'
        self.extension_dict['jpg'] = 'image/jpeg'
        self.extension_dict['jpeg'] = 'image/jpeg'
        self.extension_dict['png'] = 'image/png'
        self.extension_dict['rdf'] = 'application/rdf+xml'
        self.extension_dict[''] = 'application/binary'
        # Following 2 only while testing
        self.extension_dict['fodt'] = 'REMOVE THIS'
        self.extension_dict['py'] = 'REMOVE THIS'

    def make_manifest(self):
        global zip_file
        manifest_namespace = {"manifest":"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"}
        document = etree.Element(
                        ("{" + manifest_namespace["manifest"] + "}" + "manifest"),
                                        nsmap=manifest_namespace)
        document.attrib["{" + manifest_namespace['manifest'] + "}" + "version"] = "1.2"

        # Will come in every manifest.xml
        manifest_entry = etree.SubElement(document, 
            "{" + manifest_namespace["manifest"] + "}" + "file-entry")
        manifest_entry.attrib["{" + manifest_namespace['manifest'] + "}" + "full-path"] = "/"
        manifest_entry.attrib["{" + manifest_namespace['manifest'] + "}" + "version"] = "1.2"                               
        manifest_entry.attrib["{" + manifest_namespace['manifest'] + "}" + "media-type"] = "application/vnd.oasis.opendocument.text"

        # Will vary according to content
        manifest_file_path = zip_file.namelist()

        # for root, dirs, files in os.walk("."):
        #     for name in files:
        #         manifest_file_path.append(os.path.join(root, name))

        # file_number = 0
        # while file_number < len(manifest_file_path):
        #     manifest_file_path[file_number] = manifest_file_path[file_number].replace("./", "")
        #     file_number += 1

        for file_path in manifest_file_path:
            manifest_entry = etree.SubElement(document, 
                "{" + manifest_namespace["manifest"] + "}" + "file-entry")
            manifest_entry.attrib["{" + manifest_namespace['manifest'] + "}" + "full-path"] = file_path

            file_name = file_path.split('/')
            file_name = file_name[len(file_path.split('/')) - 1]
            file_extension = file_name.split('.')
            if len(file_extension) == 2:
                file_extension = file_extension[1]
            elif len(file_extension) == 1:
                file_extension = ""

            manifest_entry.attrib["{" + manifest_namespace['manifest'] + "}" + "media-type"] = self.extension_dict.get(file_extension)


        document_string = etree.tostring(
                document, encoding='UTF-8', xml_declaration=True, pretty_print=True)


        zip_file.writestr("META-INF/manifest" + ".xml", document_string)


        # xml_file = open("META-INF/manifest" + ".xml", "w+")
        # xml_file.write(document_string)
        # xml_file.close()


class FODT2ODT():
    def __init__(self, filename):
        self.filename = filename

    def convert(self):
        # make_folders()
        mimetype()
        OpenParseFodt(self.filename).parse()
        FileSplit(self.filename).split()
        HandleImages().handle_images()
        Manifest().make_manifest()


def main():
    global zip_file
    # make_folders()
    mimetype()
    fodt_filename = raw_input("Enter name of FODT file: ")
    OpenParseFodt(fodt_filename).parse()
    FileSplit(fodt_filename).split()
    HandleImages().handle_images()
    Manifest().make_manifest()
    zip_file.close()
    output_odt.seek(0)
    with open(fodt_filename.split(".")[0] + ".odt", "w") as odt:
        shutil.copyfileobj(output_odt, odt)



if __name__ == "__main__":main()