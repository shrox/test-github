import os
import lxml
from lxml import etree
from binary2image import Binary2Image

from copy import deepcopy
from StringIO import StringIO


filename = open('textandpic.fodt', "r+")


fodt_tree = lxml.etree.parse(filename)

fodt_root = fodt_tree.getroot()

office_ns = fodt_root.nsmap['office']
print "OFFICE_NS", office_ns


tag_dict = {}
tag_dict['meta'] = ['meta']
tag_dict['settings'] = ['settings']
tag_dict['scripts'] = ['content']
tag_dict['font-face-decls'] = ['content', 'styles']
tag_dict['styles'] = ['styles']
tag_dict['automatic-styles'] = ['content', 'styles']
tag_dict['master-styles'] = ['styles']
tag_dict['body'] = ['content']

print tag_dict

documents_processed = {}

for child in fodt_root:
    tag = child.tag.split('}')[1]
    print "TAG", tag

    for file_name in tag_dict[tag]:
        print 'FILE_NAME', file_name

        document = documents_processed.get(file_name)

        if document is None:

            document = lxml.etree.Element(('{'+office_ns+'}' + 'document-' + file_name),
                                          nsmap=fodt_root.nsmap)
            print 'DOCUMENT', document

            documents_processed[file_name] = document

        document.append(deepcopy(child))

        document_string = lxml.etree.tostring(
            document, encoding='UTF-8', xml_declaration=True)

        file = open(file_name+".xml", "w+")
        file.write(document_string)

# Following code is to locate image tag in content.xml
file = open("content.xml", 'r')
content_string = file.read()


content_root = etree.fromstring(content_string)

print "TEST2"
# """bad code, need suggestions to improve"""
try:
    os.mkdir("Pictures")
    x = 0
    # Using fodt_root.nsmap for content.xml (below) is not wrong but feels out
    # of places
    while True:
        binary_data = content_root.xpath(
            "//draw:image/office:binary-data/text()", namespaces=fodt_root.nsmap)[x]
        image_name = 'Pictures/image' + str(x) + '.jpg'
        Binary2Image(binary_data, image_name).convert2image()

        # add xlink
        node = content_root.xpath(
            "//draw:image", namespaces=fodt_root.nsmap)[x]
        node.attrib['{http://www.w3.org/1999/xlink}href'] = image_name
        node.attrib['{http://www.w3.org/1999/xlink}simple'] = "simple"
        node.attrib['{http://www.w3.org/1999/xlink}show'] = "embed"
        node.attrib['{http://www.w3.org/1999/xlink}actuate'] = "onLoad"

        x = x + 1
except IndexError:
    pass


# delete binary data
for elem in content_root.xpath("//draw:image/office:binary-data", namespaces=fodt_root.nsmap):
    elem.getparent().remove(elem)

file = open("content.xml", 'w')
file.write(etree.tostring(
    document, encoding='UTF-8', xml_declaration=True))
