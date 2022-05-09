"""
This type stub file was generated by pyright.
"""

from pptx.shapes.base import BaseShape
from pptx.shared import ParentedElementProxy

"""Graphic Frame shape and related """
class GraphicFrame(BaseShape):
    """Container shape for table, chart"""
    @property
    def chart(self):
        """The |Chart| object containing th"""
        ...
    
    @property
    def chart_part(self):
        """The |ChartPart| object containin"""
        ...
    
    @property
    def has_chart(self):
        """|True| if this graphic frame con"""
        ...
    
    @property
    def has_table(self):
        """|True| if this graphic frame con"""
        ...
    
    @property
    def ole_format(self): # -> _OleFormat:
        """Optional _OleFormat object for t"""
        ...
    
    @property
    def shadow(self):
        """Unconditionally raises |NotImple"""
        ...
    
    @property
    def shape_type(self): # -> None:
        """Optional member of `MSO_SHAPE_TY"""
        ...
    
    @property
    def table(self): # -> Table:
        """
        The |Table| object cont"""
        ...
    


class _OleFormat(ParentedElementProxy):
    """Provides attributes on an embedd"""
    def __init__(self, graphicData, parent) -> None:
        ...
    
    @property
    def blob(self):
        """Optional bytes of OLE object, su"""
        ...
    
    @property
    def prog_id(self):
        """str "progId" attribute of this e"""
        ...
    
    @property
    def show_as_icon(self):
        """True when OLE object should appe"""
        ...
    


