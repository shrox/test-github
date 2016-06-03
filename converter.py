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
tag2files = {
			'{%s}meta' % office_ns: ('meta',),
        	'{%s}settings' % office_ns: ('settings',),
       		'{%s}scripts' % office_ns: ('content',),
        	'{%s}font-face-decls' % office_ns: ('content', 'styles'),
        	'{%s}styles' % office_ns: ('styles',),
        	'{%s}automatic-styles' % office_ns: ('content', 'styles'),
        	'{%s}master-styles' % office_ns: ('styles',),
        	'{%s}body' % office_ns: ('content',),
        	}

print "START TAG2FILES"

for item in tag2files:
	print item

print "END TAG2FILES"

# print tag2files
for child in fodt_root:
	tag = child.tag.split('}')[1]
	print child.tag.split('}')[1]

	document = lxml.etree.Element(
		'{%s}document-%s' % (office_ns, tag),
		nsmap=fodt_root.nsmap)

	document.append(deepcopy(child))

	document_string = lxml.etree.tostring(document, encoding='UTF-8', xml_declaration=True)


	file = open(tag+".xml", "w+")
	file.write(document_string)

