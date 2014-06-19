# encoding: utf-8

"""
The |Section| object and related proxy classes.
"""

from __future__ import absolute_import, print_function, unicode_literals

from .shared import lazyproperty, write_only_property
from .text import Paragraph


class Section(object):
    """
    Proxy class for a WordprocessingML ``<w:sectPr>`` element.
    """
    def __init__(self, sectPr):
        super(Section, self).__init__()
        self._sectPr = sectPr

    def add_header(self):
        """
        Return a |_Header| instance.
        """
        return "header"

    @lazyproperty
    def headers(self):
        """
        |_Header| instance.
        """
        return _Header(self._sectPr)

class _Header(object):
    """
    Header
    """
    def __init__(self, sectPr):
        super(_Header, self).__init__()
        self._sectPr = sectPr
