# encoding: utf-8

"""
Provides HeaderPart and related objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..opc.package import Part
from ..oxml.shared import oxml_fromstring
from ..shared import lazyproperty


class HeaderPart(Part):
    """
    Proxy for the styles.xml part containing style definitions for a document
    or glossary.
    """
    def __init__(self, partname, content_type, element, package):
        super(HeaderPart, self).__init__(
            partname, content_type, element=element, package=package
        )

    @classmethod
    def load(cls, partname, content_type, blob, package):
        """
        Provides PartFactory interface for loading a styles part from a WML
        package.
        """
        styles_elm = oxml_fromstring(blob)
        styles_part = cls(partname, content_type, styles_elm, package)
        return styles_part

    @classmethod
    def new(cls):
        """
        Return newly created empty header part, containing only the root
        ``<w:hdr>`` element.
        """
        raise NotImplementedError

    @lazyproperty
    def header(self):
        """
        The |_Header| instance containing the header (<w:hdr> element
        proxies) for this header part.
        """
        return _Header(self._element)


class _Header(object):
    """
    FIX THIS >>> Collection of |_Header| instances corresponding to the ``<w:hdr>``
    elements in a header part.
    """
    def __init__(self, header_elm):
        super(_Header, self).__init__()
        self._header_elm = header_elm

    def __len__(self):
        return len(self._styles_elm.header_lst)