"""
This type stub file was generated by pyright.
"""

from pptx.oxml.xmlchemy import BaseOxmlElement

"""Custom element classes for text-"""
class CT_RegularTextRun(BaseOxmlElement):
    """`a:r` custom element class"""
    rPr = ...
    t = ...
    @property
    def text(self): # -> str:
        """(unicode) str containing text of"""
        ...
    
    @text.setter
    def text(self, str): # -> None:
        """*str* is unicode value to replac"""
        ...
    


class CT_TextBody(BaseOxmlElement):
    """`p:txBody` custom element class."""
    bodyPr = ...
    p = ...
    def clear_content(self): # -> None:
        """Remove all `a:p` children, but l"""
        ...
    
    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` element """
        ...
    
    @property
    def is_empty(self): # -> bool:
        """True if only a single empty `a:p"""
        ...
    
    @classmethod
    def new(cls): # -> Any:
        """
        Return a new ``<p:txBod"""
        ...
    
    @classmethod
    def new_a_txBody(cls): # -> Any:
        """
        Return a new ``<a:txBod"""
        ...
    
    @classmethod
    def new_p_txBody(cls): # -> Any:
        """
        Return a new ``<p:txBod"""
        ...
    
    @classmethod
    def new_txPr(cls): # -> Any:
        """
        Return a ``<c:txPr>`` e"""
        ...
    
    def unclear_content(self): # -> None:
        """Ensure p:txBody has at least one"""
        ...
    


class CT_TextBodyProperties(BaseOxmlElement):
    """
    <a:bodyPr> custom element c"""
    eg_textAutoFit = ...
    lIns = ...
    tIns = ...
    rIns = ...
    bIns = ...
    anchor = ...
    wrap = ...
    @property
    def autofit(self): # -> Literal[0, 2, 1] | None:
        """
        The autofit setting for"""
        ...
    
    @autofit.setter
    def autofit(self, value): # -> None:
        ...
    


class CT_TextCharacterProperties(BaseOxmlElement):
    """`a:rPr, a:defRPr, and `a:endPara"""
    eg_fillProperties = ...
    latin = ...
    hlinkClick = ...
    lang = ...
    sz = ...
    b = ...
    i = ...
    u = ...
    def add_hlinkClick(self, rId):
        """
        Add an <a:hlinkClick> c"""
        ...
    


class CT_TextField(BaseOxmlElement):
    """
    <a:fld> field element, for """
    rPr = ...
    t = ...
    @property
    def text(self): # -> str:
        """
        The text of the ``<a:t>"""
        ...
    


class CT_TextFont(BaseOxmlElement):
    """
    Custom element class for <a"""
    typeface = ...


class CT_TextLineBreak(BaseOxmlElement):
    """`a:br` line break element"""
    rPr = ...
    @property
    def text(self): # -> Literal['\u000b']:
        """Unconditionally a single vertica"""
        ...
    


class CT_TextNormalAutofit(BaseOxmlElement):
    """
    <a:normAutofit> element spe"""
    fontScale = ...


class CT_TextParagraph(BaseOxmlElement):
    """`a:p` custom element class"""
    pPr = ...
    r = ...
    br = ...
    endParaRPr = ...
    def add_br(self):
        """
        Return a newly appended"""
        ...
    
    def add_r(self, text=...):
        """
        Return a newly appended"""
        ...
    
    def append_text(self, text): # -> None:
        """Append `a:r` and `a:br` elements"""
        ...
    
    @property
    def content_children(self): # -> tuple[Unknown, ...]:
        """Sequence containing text-contain"""
        ...
    
    @property
    def text(self): # -> str:
        """str text contained in this parag"""
        ...
    


class CT_TextParagraphProperties(BaseOxmlElement):
    """
    <a:pPr> custom element clas"""
    _tag_seq = ...
    lnSpc = ...
    spcBef = ...
    spcAft = ...
    defRPr = ...
    lvl = ...
    algn = ...
    @property
    def line_spacing(self): # -> None:
        """
        The spacing between bas"""
        ...
    
    @line_spacing.setter
    def line_spacing(self, value): # -> None:
        ...
    
    @property
    def space_after(self): # -> None:
        """
        The EMU equivalent of t"""
        ...
    
    @space_after.setter
    def space_after(self, value): # -> None:
        ...
    
    @property
    def space_before(self): # -> None:
        """
        The EMU equivalent of t"""
        ...
    
    @space_before.setter
    def space_before(self, value): # -> None:
        ...
    


class CT_TextSpacing(BaseOxmlElement):
    """
    Used for <a:lnSpc>, <a:spcB"""
    spcPct = ...
    spcPts = ...
    def set_spcPct(self, value): # -> None:
        """
        Set spacing to *value* """
        ...
    
    def set_spcPts(self, value): # -> None:
        """
        Set spacing to *value* """
        ...
    


class CT_TextSpacingPercent(BaseOxmlElement):
    """
    <a:spcPct> element, specify"""
    val = ...


class CT_TextSpacingPoint(BaseOxmlElement):
    """
    <a:spcPts> element, specify"""
    val = ...


