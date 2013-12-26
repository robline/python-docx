# encoding: utf-8

"""
Document parts such as _Document, and closely related classes.
"""

from docx.enum.shape import WD_INLINE_SHAPE
from docx.opc.oxml import serialize_part_xml
from docx.opc.package import Part
from docx.oxml.shared import nsmap, oxml_fromstring
from docx.shared import lazyproperty
from docx.table import Table
from docx.text import Paragraph


class _Document(Part):
    """
    Main document part of a WordprocessingML (WML) package, aka a .docx file.
    """
    def __init__(self, partname, content_type, document_elm, package):
        super(_Document, self).__init__(
            partname, content_type, package=package
        )
        self._element = document_elm

    @property
    def blob(self):
        return serialize_part_xml(self._element)

    @property
    def body(self):
        """
        The |_Body| instance containing the content for this document.
        """
        return _Body(self._element.body)

    @lazyproperty
    def inline_shapes(self):
        """
        The |InlineShapes| instance containing the inline shapes in the
        document.
        """
        return InlineShapes(self._element.body)

    @staticmethod
    def load(partname, content_type, blob, package):
        document_elm = oxml_fromstring(blob)
        document = _Document(partname, content_type, document_elm, package)
        return document


class _Body(object):
    """
    Proxy for ``<w:body>`` element in this document, having primarily a
    container role.
    """
    def __init__(self, body_elm):
        super(_Body, self).__init__()
        self._body = body_elm

    def add_paragraph(self):
        """
        Return a paragraph newly added to the end of body content.
        """
        p = self._body.add_p()
        return Paragraph(p)

    def add_table(self, rows, cols):
        """
        Return a table having *rows* rows and *cols* cols, newly appended to
        the main document story.
        """
        tbl = self._body.add_tbl()
        table = Table(tbl)
        for i in range(cols):
            table.columns.add()
        for i in range(rows):
            table.rows.add()
        return table

    def clear_content(self):
        """
        Return this |_Body| instance after clearing it of all content.
        Section properties for the main document story, if present, are
        preserved.
        """
        self._body.clear_content()
        return self

    @property
    def paragraphs(self):
        return [Paragraph(p) for p in self._body.p_lst]

    @property
    def tables(self):
        """
        A sequence containing all the tables in the document, in the order
        they appear.
        """
        return [Table(tbl) for tbl in self._body.tbl_lst]


class InlineShape(object):
    """
    Proxy for an ``<wp:inline>`` element, representing the container for an
    inline graphical object.
    """
    def __init__(self, inline):
        super(InlineShape, self).__init__()
        self._inline = inline

    @property
    def type(self):
        graphicData = self._inline.graphic.graphicData
        uri = graphicData.uri
        if uri == nsmap['pic']:
            blip = graphicData.pic.blipFill.blip
            if blip.link is not None:
                return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap['c']:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap['dgm']:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED


class InlineShapes(object):
    """
    Sequence of |InlineShape| instances, supporting len(), iteration, and
    indexed access.
    """
    def __init__(self, body_elm):
        super(InlineShapes, self).__init__()
        self._body = body_elm

    def __getitem__(self, idx):
        """
        Provide indexed access, e.g. 'inline_shapes[idx]'
        """
        try:
            inline = self._inline_lst[idx]
        except IndexError:
            msg = "inline shape index [%d] out of range" % idx
            raise IndexError(msg)
        return InlineShape(inline)

    def __iter__(self):
        return (InlineShape(inline) for inline in self._inline_lst)

    def __len__(self):
        return len(self._inline_lst)

    def add_picture(self, image_path_or_stream):
        """
        Add the image at *image_path_or_stream* to the document at its native
        size. The picture is placed inline in a new paragraph at the end of
        the document.
        """

    @property
    def _inline_lst(self):
        body = self._body
        xpath = './w:p/w:r/w:drawing/wp:inline'
        return body.xpath(xpath, namespaces=nsmap)
