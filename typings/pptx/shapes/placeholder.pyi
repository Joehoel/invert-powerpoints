"""
This type stub file was generated by pyright.
"""

from pptx.shapes.autoshape import Shape
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.picture import Picture

"""Placeholder-related objects.

Sp"""
class _InheritsDimensions:
    """
    Mixin class that provides i"""
    @property
    def height(self): # -> Any | None:
        """
        The effective height of"""
        ...
    
    @height.setter
    def height(self, value): # -> None:
        ...
    
    @property
    def left(self): # -> Any | None:
        """
        The effective left of t"""
        ...
    
    @left.setter
    def left(self, value): # -> None:
        ...
    
    @property
    def shape_type(self):
        """
        Member of :ref:`MsoShap"""
        ...
    
    @property
    def top(self): # -> Any | None:
        """
        The effective top of th"""
        ...
    
    @top.setter
    def top(self, value): # -> None:
        ...
    
    @property
    def width(self): # -> Any | None:
        """
        The effective width of """
        ...
    
    @width.setter
    def width(self, value): # -> None:
        ...
    


class _BaseSlidePlaceholder(_InheritsDimensions, Shape):
    """Base class for placeholders on s"""
    @property
    def is_placeholder(self): # -> Literal[True]:
        """
        Boolean indicating whet"""
        ...
    
    @property
    def shape_type(self):
        """
        Member of :ref:`MsoShap"""
        ...
    


class BasePlaceholder(Shape):
    """
    NOTE: This class is depreca"""
    @property
    def idx(self):
        """
        Integer placeholder 'id"""
        ...
    
    @property
    def orient(self):
        """
        Placeholder orientation"""
        ...
    
    @property
    def ph_type(self):
        """
        Placeholder type, e.g. """
        ...
    
    @property
    def sz(self):
        """
        Placeholder 'sz' attrib"""
        ...
    


class LayoutPlaceholder(_InheritsDimensions, Shape):
    """
    Placeholder shape on a slid"""
    ...


class MasterPlaceholder(BasePlaceholder):
    """
    Placeholder shape on a slid"""
    ...


class NotesSlidePlaceholder(_InheritsDimensions, Shape):
    """
    Placeholder shape on a note"""
    ...


class SlidePlaceholder(_BaseSlidePlaceholder):
    """
    Placeholder shape on a slid"""
    ...


class ChartPlaceholder(_BaseSlidePlaceholder):
    """Placeholder shape that can only """
    def insert_chart(self, chart_type, chart_data): # -> PlaceholderGraphicFrame:
        """
        Return a |PlaceholderGr"""
        ...
    


class PicturePlaceholder(_BaseSlidePlaceholder):
    """Placeholder shape that can only """
    def insert_picture(self, image_file): # -> PlaceholderPicture:
        """Return a |PlaceholderPicture| ob"""
        ...
    


class PlaceholderGraphicFrame(GraphicFrame):
    """
    Placeholder shape populated"""
    @property
    def is_placeholder(self): # -> Literal[True]:
        """
        Boolean indicating whet"""
        ...
    


class PlaceholderPicture(_InheritsDimensions, Picture):
    """
    Placeholder shape populated"""
    ...


class TablePlaceholder(_BaseSlidePlaceholder):
    """Placeholder shape that can only """
    def insert_table(self, rows, cols): # -> PlaceholderGraphicFrame:
        """Return |PlaceholderGraphicFrame|"""
        ...
    


