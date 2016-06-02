from lxml import etree
import lxml

from StringIO import StringIO

#Extract namespaces and put them into xml files
XMLfile = open('textandpic.fodt', "r+")
doc = XMLfile.read()
tree = etree.fromstring(doc)

# For meta.xml
meta.xml = open('meta.xml', 'w')
page = etree.E