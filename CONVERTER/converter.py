from lxml import etree
import lxml

from copy import deepcopy
from StringIO import StringIO

# #Extract namespaces and put them into xml files
# XMLfile = open('textandpic.fodt', "r+")
# doc = XMLfile.read()
# tree = etree.fromstring(doc)

# # For meta.xml
# meta.xml = open('meta.xml', 'w')
# page = etree.E

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
