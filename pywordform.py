"""
pywordform

a python module to parse a Microsoft Word form in docx format, and
extract all field values with their tags into a dictionary.

Project website: http://www.decalage.info/python/pywordform

Copyright (c) 2012, Philippe Lagadec (http://www.decalage.info)
All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

 * Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
 * Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

__version__ = '0.03'

#------------------------------------------------------------------------------
# CHANGELOG:
# 2012-02-17 v0.01 PL: - first version
# 2012-04-19 v0.02 PL: - added support for multiline text fields
# 2021-11-17 v0.03 FB: - made it work with Python 3.9
#------------------------------------------------------------------------------
#TODO:
# - recognize date fields and extract fulldate
# - test docm files, add support if needed
# - support legacy fields
# - CSV output (option)
# - more advanced parser returning a list of field objects: keep order, and
#   get fields with no tag, extract other attributes such as title
# - add possibility to modify field values and write docx back to disk

#------------------------------------------------------------------------------
import zipfile, sys
import xml.etree.ElementTree as ET

# namespace for word XML tags:
NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# XML tags in Word forms:
TAG_FIELD         = NS_W+'sdt'
TAG_FIELDPROP     = NS_W+'sdtPr'
TAG_FIELDTAG      = NS_W+'tag'
ATTR_FIELDTAGVAL  = NS_W+'val'
TAG_FIELD_CONTENT = NS_W+'sdtContent'
TAG_RUN           = NS_W+'r'
TAG_TEXT          = NS_W+'t'
TAG_BREAK         = NS_W+'br'


def parse_multiline (field_content_elem):
    """
    parse an XML element (<w:sdtContent>) containing a single or multiline text
    field
    """
    value = ''
    # iterate over all children elements
    for elem in field_content_elem.iter():
        # extract text:
        if elem.tag == TAG_TEXT:
            value += elem.text
        # and line breaks:
        elif elem.tag == TAG_BREAK:
            value += '\n'
    return value


def parse_form(filename):
    fields = {}
    zfile = zipfile.ZipFile(filename)
    form = zfile.read('word/document.xml')
    xmlroot = ET.fromstring(form)
    for field in xmlroot.iter(TAG_FIELD):
        field_tag = field.find(TAG_FIELDPROP+'/'+TAG_FIELDTAG)
        if field_tag is not None:
            tag = field_tag.get(ATTR_FIELDTAGVAL, None)
##            field_value = field.find(TAG_FIELD_CONTENT+'/'+TAG_RUN+'/'+TAG_TEXT)
##            if field_value is not None:
##                value = field_value.text
            field_content = field.find(TAG_FIELD_CONTENT)
            if field_content is not None:
                value = parse_multiline(field_content)
                fields[tag] = value
    zfile.close()
    return fields


if __name__ == '__main__':
    fields = parse_form(sys.argv[1])
    for tag, value in sorted(fields.items()):
        print('%s = "%s"' % (tag, value))
