# encoding: utf-8

"""
Custom element classes related to the header part
"""

from docx.oxml.shared import nsmap, OxmlBaseElement, qn


class CT_Header(OxmlBaseElement):
    """
    A ``<w:???>`` element, representing a header definition
    """
    @property
    def pPr(self):
        return self.find(qn('w:pPr'))


class CT_Headers(OxmlBaseElement):
    """
    ``<w:???>`` element, the root element of a styles part, i.e.
    styles.xml
    """
    def header_having_headerId(self, headerId):
        """
        Return the ``<w:???>`` child element having ``headerId`` attribute
        matching *headerId*.
        """
        #FIX THIS \/
        xpath = './w:style[@w:styleId="%s"]' % styleId
        try:
            return self.xpath(xpath, namespaces=nsmap)[0]
        except IndexError:
            raise KeyError('no <w:style> element with styleId %d' % styleId)

    @property
    def header_lst(self):
        """
        List of <w:???> child elements.
        """
        return self.findall(qn('w:???'))
