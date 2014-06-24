# encoding: utf-8

"""
Custom element classes related to the header part
"""

from docx.oxml.shared import nsmap, OxmlBaseElement, qn


class CT_Header(OxmlBaseElement):
    """
    A ``<w:hdr>`` element, representing a header definition
    """
    @property
    def pPr(self):
        return self.find(qn('w:pPr'))


    def header_having_headerId(self, headerId):
        """
        Return the ``<w:hdr>`` child element having ``headerId`` attribute
        matching *headerId*.
        """
        xpath = './w:hdr[@w:hdrId="%s"]' % headerId
        try:
            return self.xpath(xpath, namespaces=nsmap)[0]
        except IndexError:
            raise KeyError('no <w:style> element with hdrId %d' % headerId)

    @property
    def header_lst(self):
        """
        List of <w:hdr> child elements.
        """
        return self.findall(qn('w:hdr'))

    @classmethod
    def new(cls, hdr_id, abstractHdr_id):
        """
        Return a new ``<w:hdr>`` element having numId of *num_id* and having
        a ``<w:abstractNumId>`` child with val attribute set to
        *abstractNum_id*.
        """
        header = OxmlElement('w:hdr')
        header.hdrId = hdr_id
        abstractHdrId = CT_DecimalNumber.new(
            'w:abstractHdrId', abstractHdr_id
        )
        nhdr.append(abstractHdrId)
        return header