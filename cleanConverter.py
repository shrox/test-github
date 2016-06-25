import base64
import shutil

from lxml import etree
from copy import deepcopy
from zipfile import ZipFile
from StringIO import StringIO


def mimetype(zip_file):
    zip_file.writestr("mimetype", "application/vnd.oasis.opendocument.text")

def binary2image(zip_file, binary_data, image_name):
    image = base64.b64decode(binary_data)
    zip_file.writestr(image_name, image)

def parse_fodt(filename):
    filename = open(filename, "r")

    fodt_tree = etree.parse(filename)
    fodt_root = fodt_tree.getroot()
    fodt_namespaces = fodt_root.nsmap

    return (fodt_root, fodt_namespaces)

def split_file(zip_file, fodt_root, fodt_namespaces):
    tag_dict = {}
    tag_dict['meta'] = ['meta']
    tag_dict['settings'] = ['settings']
    tag_dict['scripts'] = ['content']
    tag_dict['font-face-decls'] = ['content', 'styles']
    tag_dict['styles'] = ['styles']
    tag_dict['automatic-styles'] = ['content', 'styles']
    tag_dict['master-styles'] = ['styles']
    tag_dict['body'] = ['content']

    documents_processed = {}
    for child in fodt_root:
        tag = child.tag.split('}')[1]

        for xml_filename in tag_dict[tag]:
            document = documents_processed.get(xml_filename)

            if document is None:
                document = etree.Element(
                        ('{%s}document-%s' % (fodt_namespaces['office'], xml_filename)),
                                nsmap=fodt_namespaces)
                documents_processed[xml_filename] = document

            document.append(deepcopy(child))
            document_string = etree.tostring(
                    document, encoding='UTF-8', xml_declaration=True, pretty_print=True)

            zip_file.writestr("%s.xml" % (xml_filename), document_string)

def handle_images(zip_file, fodt_root, fodt_namespaces):
    content_string = zip_file.read("content.xml")
    content_root = etree.fromstring(content_string)

    image_number = 0

    # possible bad code, improve
    try:
        while True:
            binary_data = content_root.xpath(
                "//draw:image/office:binary-data/text()",
                namespaces=fodt_namespaces)[image_number]

            image_name = "Pictures/image" + str(image_number) + ".jpg"

            binary2image(zip_file, binary_data, image_name)

            node = content_root.xpath(
                "//draw:image", namespaces=fodt_namespaces)[image_number]
            node.attrib["{%s}href" % (fodt_namespaces['xlink'])] = image_name
            node.attrib["{%s}simple" % (fodt_namespaces['xlink'])] = "simple"
            node.attrib["{%s}show" % (fodt_namespaces['xlink'])] = "embed"
            node.attrib["{%s}actuate" % (fodt_namespaces['xlink'])] = "onLoad"

            image_number = image_number + 1

    except IndexError:
        pass

    # Delete  binary data
    for elem in content_root.xpath(
        "//draw:image/office:binary-data", namespaces=fodt_root.nsmap):
            elem.getparent().remove(elem)

    # Write to disk
    content_newstring = etree.tostring(
                   content_root, encoding='UTF-8', 
                    xml_declaration=True)

    zip_file.writestr("content.xml", content_newstring)

def make_manifest(zip_file):
    # Might need to add more possible extensions for Configurations2 folder, etc.
    extension_dict = {}
    extension_dict['xml'] = 'text/xml'
    extension_dict['jpg'] = 'image/jpeg'
    extension_dict['jpeg'] = 'image/jpeg'
    extension_dict['png'] = 'image/png'
    extension_dict['rdf'] = 'application/rdf+xml'
    extension_dict[''] = 'application/binary'

    manifest_namespace = {"manifest":"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"}
    document = etree.Element(
                ("{%s}manifest" % (manifest_namespace["manifest"])),
                                        nsmap=manifest_namespace)
    document.attrib["{%s}version" % (manifest_namespace["manifest"])] = "1.2"

    # Will come in every manifest.xml
    manifest_entry = etree.SubElement(document,
        "{%s}file-entry" % (manifest_namespace["manifest"]))
    manifest_entry.attrib["{%s}full-path" % (manifest_namespace["manifest"])] = "/"
    manifest_entry.attrib["{%s}version" % (manifest_namespace["manifest"])] = "1.2"                               
    manifest_entry.attrib["{%s}media-type" % (manifest_namespace["manifest"])] = "application/vnd.oasis.opendocument.text"

    # Will vary according to content
    # need to avoid repetition of locations in manifest
    manifest_file_path = zip_file.namelist()

    for file_path in manifest_file_path:
        manifest_entry = etree.SubElement(document,
            "{%s}file-entry" % (manifest_namespace["manifest"]))

        manifest_entry.attrib["{%s}full-path" % (manifest_namespace["manifest"])] = file_path

        file_name = file_path.split('/')
        file_name = file_name[len(file_path.split('/')) - 1]
        file_extension = file_name.split('.')
        if len(file_extension) == 2:
            file_extension = file_extension[1]
        elif len(file_extension) == 1:
            file_extension = ""

        manifest_entry.attrib["{%s}media-type" % (manifest_namespace["manifest"])] = extension_dict.get(file_extension)

    document_string = etree.tostring(
                document, encoding='UTF-8', xml_declaration=True, pretty_print=True)
    zip_file.writestr("META-INF/manifest" + ".xml", document_string)


class FODT2ODT():
    def __init__(self, filename):
        self.filename = filename

    def convert(self):
        output_odt = StringIO()
        zip_file = ZipFile(output_odt, "w")

        mimetype()

        fodt_filename = self.filename
        fodt_root, fodt_namespaces = parse_fodt(fodt_filename)

        split_file(zip_file, fodt_root, fodt_namespaces)
        handle_images(zip_file, fodt_root, fodt_namespaces)
        make_manifest(zip_file)

        zip_file.close()
        output_odt.seek(0)
        
        with open("%s.odt" % (fodt_filename.split(".")[0]), "w") as odt:
            shutil.copyfileobj(output_odt, odt)

def main():
    output_odt = StringIO()
    zip_file = ZipFile(output_odt, "w")

    mimetype(zip_file)

    fodt_filename = raw_input("Enter name of FODT file: ")
    fodt_root, fodt_namespaces = parse_fodt(fodt_filename)

    split_file(zip_file, fodt_root, fodt_namespaces)
    handle_images(zip_file, fodt_root, fodt_namespaces)
    make_manifest(zip_file)

    zip_file.close()
    output_odt.seek(0)

    with open("%s.odt" % (fodt_filename.split(".")[0]), "w") as odt:
            shutil.copyfileobj(output_odt, odt)

if __name__ == "__main__":main()




