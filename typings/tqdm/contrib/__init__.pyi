"""
This type stub file was generated by pyright.
"""

import sys
from functools import wraps
from ..auto import tqdm as tqdm_auto
from ..std import tqdm
from ..utils import ObjectWrapper

"""
Thin wrappers around common fun"""
__author__ = ...
__all__ = ['tenumerate', 'tzip', 'tmap']
class DummyTqdmFile(ObjectWrapper):
    """Dummy file-like that will write """
    def __init__(self, wrapped) -> None:
        ...
    
    def write(self, x, nolock=...): # -> None:
        ...
    
    def __del__(self): # -> None:
        ...
    


def builtin_iterable(func): # -> (*args: Unknown, **kwargs: Unknown) -> list[Unknown]:
    """Wraps `func()` output in a `list"""
    ...

def tenumerate(iterable, start=..., total=..., tqdm_class=..., **tqdm_kwargs): # -> tqdm | enumerate[Unknown]:
    """
    Equivalent of `numpy.ndenum"""
    ...

@builtin_iterable
def tzip(iter1, *iter2plus, **tqdm_kwargs): # -> Generator[tuple[Unknown], None, None]:
    """
    Equivalent of builtin `zip`"""
    ...

@builtin_iterable
def tmap(function, *sequences, **tqdm_kwargs): # -> Generator[Unknown, None, None]:
    """
    Equivalent of builtin `map`"""
    ...

