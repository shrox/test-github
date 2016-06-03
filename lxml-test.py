from lxml import etree
from io import StringIO, BytesIO

# xml = '<a xmlns="test"><b xmlns="test"/></a>'
# print "xml:", xml
# print
# root = etree.fromstring(xml)
# print "etree.tostring(root) :", etree.tostring(root)
# print
# tree = etree.parse(StringIO(xml.decode('utf-8')))
# print "etree.tostring(tree.getroot()):", etree.tostring(tree.getroot())


# fo = open("textandpic.fodt", "r+")

# doc = etree.parse(fo)

# print doc.xpath('//form')

# for df in doc.xpath('//form'):
# 	print "sex", df


# Following code is to access namespace
fo = open('textandpic.fodt', "r+") 

doc = fo.read()

root = etree.fromstring(doc)




print "xmlns:office=" + '"'+root.nsmap['office']+'"'


#To access regular attributes
print root.attrib.get('version')
# print root.attrib.set()
print "keys", root.attrib.keys()
print "values", root.attrib.values()
print "items", root.attrib.items()

# for ns in sorted(root.attrib.items()):
#     print ns

# Testing if I can create XML
