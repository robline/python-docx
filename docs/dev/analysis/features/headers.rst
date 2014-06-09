
Headers
=======

Headers are text, graphics or data (such as page number, date, document title, and so on) that
can appear at the top of each page in a word document. A header appears in the top margin above 
the main document content on the page. Headers are applied by specifying the headers for all pages 
in a particular section of a document. Within each section of a document there can be up to three 
different types of headers:

* First page header
* Odd page header
* Even page header

First page headers specify a unique header which shall appear on the first page of a
section. Odd page headers specify a unique header which shall appear on all odd
numbered pages for a given section. Even page headers specify a unique header which
shall appear on all even numbered pages in a given section.


Differences between a document without and with a header
--------------------------------------------------------

If you create a default document and save it (let's call that test.docx), then add a header to it like so:

    This is a header.   x of xx
    
the following changes will occur in the package:

1) A part called header1.xml will be added to the package with the following pathname:
    /word/header1.xml

2) A new relationship is specified at word/_rels/document.xml.rels:

    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
    *<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml" />*
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml" />
    ...
    <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml" />
    </Relationships>

3) Within the <w:sectPr> element of document.xml, there will be a new element called headerReference:

    <w:sectPr>
        *<w:headerReference w:type="default" r:id="rId2"/>*
        <w:type w:val="nextPage"/>
        <w:pgSz w:w="12240" w:h="15840"/>
        ...
    </w:sectPr>

Structure of header1.xml
------------------------

    <?xml version="1.0" encoding="UTF-8"?>
    <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
        xmlns:o="urn:schemas-microsoft-com:office:office" 
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
        xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" 
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Header" />
            <w:rPr />
        </w:pPr>
        <w:r>
            <w:rPr />
            <w:t xml:space="preserve">This is a header.  </w:t>
        </w:r>
        <w:r>
            <w:rPr />
            <w:fldChar w:fldCharType="begin" />
        </w:r>
        <w:r>
            <w:instrText>PAGE</w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate" />
        </w:r>
        <w:r>
            <w:t>1</w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end" />
        </w:r>
        <w:r>
            <w:rPr />
            <w:t xml:space="preserve"> of </w:t>
        </w:r>
        <w:r>
            <w:rPr />
            <w:fldChar w:fldCharType="begin" />
        </w:r>
        <w:r>
            <w:instrText>NUMPAGES</w:instrText>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="separate" />
        </w:r>
        <w:r>
            <w:t>1</w:t>
        </w:r>
        <w:r>
            <w:fldChar w:fldCharType="end" />
        </w:r>
    </w:p>
    </w:hdr>

Different Even/Odd Page Headers and Footers
-------------------------------------------

The evenAndOddHeaders element specifies whether sections in the document shall have different headers and 
footers for even and odd pages (an odd page header/footer and an even page header/footer).
If the val attribute is set to True, then each section in the document shall use an odd page header for all odd
numbered pages in the section, and an even page header for all even numbered pages in the section (counting
from the starting value of page numbering for the parent section to determine if the first page is even or odd, as
specified with the start attribute on the pgNumType element). If the val attribute is set to False, then all pages
in a section shall use the odd page header.

    <w:hdr>
        <w:p>
            <w:r>
                <w:t>First</w:t>
            </w:r>
        </w:p>
    </w:hdr>
    
Even page header part:

    <w:hdr>
        <w:p>
            <w:r>
                <w:t>Even</w:t>
            </w:r>
        </w:p>
    </w:hdr>
    
Odd page header part:

    <w:hdr>
        <w:p>
            <w:r>
                <w:t>Odd</w:t>
            </w:r>
        </w:p>
    </w:hdr>


Candidate protocol -- document.add_header()
-------------------------------------

Pending...


Relevant sections in the ISO Spec
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* 17.2 Main Document Story
* 17.10 Headers and Footers
* 17.10.4 hdr (Header)
* 17.10.5 headerReference (Header Reference)
* 17.10.6 titlePg (Different First Page Headers and Footers)

