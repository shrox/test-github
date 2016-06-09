import os
import lxml
from lxml import etree
from binary2image import binary2image

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


for child in fodt_root:
    tag = child.tag.split('}')[1]
    print tag

    document = lxml.etree.Element(('{'+office_ns+'}'+'document-'+tag),
                                  nsmap=fodt_root.nsmap)

    document.append(deepcopy(child))

    document_string = lxml.etree.tostring(
        document, encoding='UTF-8', xml_declaration=True)
    
    file = open(tag_dict[tag][0]+".xml", "w+")
    file.write(document_string)

    if len(tag_dict[tag]) == 2:
        file = open(tag_dict[tag][1]+".xml", "w+")
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
    while True:
        binary_data = content_root.xpath("//draw:image/office:binary-data/text()", namespaces={"draw":"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", 
                                                                           "office":"urn:oasis:names:tc:opendocument:xmlns:office:1.0"})[x]
        image_name = 'Pictures/image' + str(x) + '.jpg'
        binary2image(binary_data, image_name).convert2image()

        x = x + 1
except:
    pass


#delete binary data
for elem in content_root.xpath("//draw:image/office:binary-data", namespaces={"draw":"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", 
                                                                           "office":"urn:oasis:names:tc:opendocument:xmlns:office:1.0"}):
    elem.getparent().remove(elem)

file = open("content.xml", 'w')
file.write(etree.tostring(content_root))

#add xlink
print "TEST1"
content_root.xpath("//draw:image", namespaces={"draw":"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"})[0].set('xlink:href', 'test2')
print "TEST", etree.tostring(content_root)


# xlink:href="Pictures/1000000000000132000001C1B172F5629754D00D.jpg"
#  xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"

# content_root.xpath("//draw:image", namespaces={"draw":"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"})[0].attrib['xlink:href'] = 'image'.attrib["xlink:type"] = "simple".attrib["xlink:show"] = "embed".attrib[
#                                                     "xlink:actuate"] = "onLoad"