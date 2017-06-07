# -*- coding: utf-8 -*-
"""
PyVisio visDocuments - Visio Document manipulation library

See docstring for class VisDocument for usage
"""
#TODO docstring

__author__ = 'Ivo Velcovsky'
__email__ = 'velcovsky@email.cz'
__copyright__ = "Copyright (c) 2015"
__license__ = "MIT"
__status__ = "Development"

from .visCOM import *
from .documents import *
from .stencils import *
from .shapes import *

if __name__ == "__main__":
    import doctest
    doctest.testmod()
