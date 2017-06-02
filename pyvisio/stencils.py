# -*- coding: utf-8 -*-
"""PyVisio stencils - Visio Stencil manipulation module"""

import logging
import os
from visCOM import visCOMobject as vCOM
from visCOM import visCOMconstants as vC
from visCOM import com_error

logging.info("stencils loaded...")
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


class VisStencils(object):
    """
    === VisStencils ===

    Introduction
    ------------

    This library is part of PyVisio package. It's purpose is loading the stencil files and manipulating the master
    shapes. VisStencils could be either used for programmatic generating of Visio files or for exporting shapes from
    Visio stencils to other formats (only those supported by Visio itself).

    Type of objects:
    Master Object - https://msdn.microsoft.com/en-us/library/office/ff765439.aspx

    Shape Object - https://msdn.microsoft.com/EN-US/library/office/ff768546.aspx

    Usage
    -----

    For testing we'll use low level DOM functions to be independent of other VisABC objects, no need here to explain
    these as we'll create high level Python API later on.

    This class relies on visCOM which is responsible for creating instance of Visio.Application COM object.

    Doctests will open new Visio window and draw some shapes, just to have some visual debug/test output.

    So let's create new Visio document and get the active/first page handle
    >>> vCOM.Visible = True
    >>> test_visio = vCOM.Documents.Add("")
    >>> test_visio_page = vCOM.ActiveDocument.Pages.Item(1)

    and now we could start with small tutorial...

    1. Creating instance

        1. With no parameters

    By default empty instance is created

    >>> vS = VisStencils()

    For quick checking what shapes are loaded you could print VisStencils object, __str__ function returns list
    like output

    >>> print vS #doctest
    []

    At this moment object instance is quite useless but you can add stencil file later on.

        2. With stencil as param

    You could create instance and load one or more stencils when creating instance of VisStencils

    Note: You don't need to add full path if stencil is in Visio's path, but you have to include the file extension

    >>> vS = VisStencils("Computers and Monitors.vss", "periph_m.vss")

    You already know you could print the VisStencils object...

    >>> print vS #doctest:+ELLIPSIS
    ['...

    This time we got list of loaded master shapes.

    2. Adding stencils to the object

    When you need to load additional stencil, just use method VisStencils.add_stencil() with name of stencil file
    including the extension. For the stencil files located in Visio path you can use just filename.

    >>> vS.add_stencil("Computers and Monitors.vss")


    But of course you could use as well full path
    >>> vS.add_stencil(os.getcwd() + os.path.normpath("/data/network_connector.vss"))


    When you accidentally use wrong filename it will throw com_error exception

    >>> vS.add_stencil("wrong name.vss") #doctest:+ELLIPSIS
    Traceback (most recent call last):
    ...
    com_error:...

    and log lines will be added
    ERROR:__main__:Loading of stencil file failed for: wrong name.vss
    ERROR:__main__:COM Exception: ...

    3. Available stencils

    I've implemented __iter__ function to return list of available shapes so you could simply use 'in' to get info
    if the shape you want to use is available

    >>> "Server" in vS
    True

    >>> "NonExistingShape" in vS
    False

    4. Get the stencil master shape object for further usage

        1. Using VisStencils["ObjectName"]

    Stencil itself is quite useless without using the masters/shapes in the Visio pages, so real usage is to simply
    pick the master shape object from stencil for further usage

    I've implemented __getitem__ function so we could treat the VisStencils as dictionary of master shapes and
    ask for master shape directly by it's name as key

    >>> my_shape = vS["Server"]

    And to do finally something interesting we'll drop master shape we've just picked into the page we've created
    at the beginning of our tutorial.

    >>> test_visio_object = test_visio_page.Drop(my_shape, 1.0, 1.0)

    You should see the new Visio file with Server icon

    When you ask for non existing master shape you'll KeyError exception.

    >>> my_shape2 = vS["NonExistingShape"]
    Traceback (most recent call last):
    ...
    KeyError: 'NonExistingShape'

    and log line will be added
    ERROR:__main__:Non existing shape: NonExistingShape

        2. Using VisStencils.get method

    In some cases the names of master shapes are same in different stencil files i. e. 'Dynamic connector', but
    we count with this and our VisStencil object keeps internally dict with all the master shape names and
    their mapping back to the stencil from where they come from. LIFO applies in the mapping, so last master shape
    added will be first returned. In case user would load other stencil with same master shape name it's safe to
    use VisStencils["DuplicatedShapeName"] and will get what she/he/it probably expects...the last loaded shape.

    In specific cases, it might be useful to have possibility to explicitly ask for master shape with specific index
    for these purposes we have VisStencils.get()

    There is 1 mandatory parameter - Master Shape Name
    There is 1 optional parameter - Index of Master Shape in the internal dict, with default index = 0

    For non existing shape VisStencil.get() raises KeyError same way like the first shape access method

    >>> my_shape3 = vS.get('NonExistingShape')
    Traceback (most recent call last):
    ...
    KeyError: 'NonExistingShape'

    and again line in log
    ERROR:__main__:Non existing shape: NonExistingShape

    When VisStencils.get() is called with wrong Index result is raised ValueError

    >>> my_shape4 = vS.get('PC',10)
    Traceback (most recent call last):
    ...
    ValueError: 10

    and as expected, another log line
    ERROR:__main__:Bad shape index: 10

    When called with correct parameters it will return the the object.

    >>> my_shape5 = vS.get('Dynamic connector',0)
    >>> my_shape6 = vS.get('Dynamic connector',1)

    Let's drop both objects to the page... just to see them.

    Note: actually 'Dynamic connector' is contained in quite lot of standard MS Visio Stencils and it's really
    same object with same properties in all the stencils...so a bit stupid example :-) But we just try to explain
    how the shape index is working.

    >>> test_visio_object5 = test_visio_page.Drop(my_shape5, 3.0, 3.0)
    >>> test_visio_object6 = test_visio_page.Drop(my_shape6, 5.0, 5.0)

    You should see new connectors object created on page

    5. Other available functions

    Sometimes it is useful to know some details about the master shape, for this purpose we've added function
    get_info() which returns the dictionary with some shape parameters: shapesCount, width, objectType, size and height.
    Values are in Inches.

    >>> info = vS.get_info("PC")
    >>> print info["width"]
    0.984251968504
    >>> print info["height"]
    0.984251968504
    >>> print info["size"]
    (0.9842519685039419, 0.984251968503937)

    This info could be used for example when you calculate placement of shapes on page.

    Another feature is master shape export() function, which is using Visio export API and could be used as following:

    >>> vS.export(os.getcwd() + "\data\stencil.svg","PC")

    This example will store PC shape in current directory as stencil.svg. Export file formats are those which are
    available in your version of Visio.

    >>> vS.export(os.getcwd() + "\data\stencil.png","PC")

    As you can see you can export to bitmap as well.

    For debugging purposes there is as well implemented __repr__ function.

    >>> print vS.__repr__() #doctest:+ELLIPSIS
    <object VisStencils("Computers and Monitors", "periph_m"...

    That is shortly all

    Now you should manually close Visio with created demo diagram.
    """

    def __init__(self, *args):
        """
        Create instance and optionaly load stencil file(s)

        :type args: list of Str
        """
        self.__stencils = {}
        self.__shapes = {}
        #self.add_stencil("Basic Shapes.vss")
        for stencil in args:
            self.add_stencil(stencil)

    def __str__(self):
        """
        Return list style string representing list of loaded shapes

        :return: Str
        """
        if len(self.__shapes) > 0:
            return "['" + "', '".join(self.__shapes.keys()) + "']"
        else:
            return "[]"

    def __repr__(self):
        """
        Return python style representation of object

        :return: Str
        """
        return "<object VisStencils(\"{0}\") at {1}>".format("\", \"".join(self.__stencils), hex(id(self)))

    def __iter__(self):
        """
        Iterate over loaded shapes

        :return: iterator/generator
        """
        for shape in self.__shapes:
            yield shape

    def __getitem__(self, shape):
        """
        Return COM object representing the shape with index 0.

        :param shape: Str
        :return: COM object - Shape
        """
        return self.get(shape, 0)

    def get(self, shape, index=0):
        """
        Return COM object representing the shape by it's index (remember LIFO applies).

        If object doesn't exist or index is out of range raise exception

        :param shape: Str
        :param index: Int
        :return: COM object - Shape
        """
        if shape in self.__shapes:
            if len(self.__shapes[shape]) > index >= 0:
                return self.__stencils[self.__shapes[shape][index]].Masters.ItemU(shape)
            else:
                logger.error("Bad shape index: %s", index)
                raise ValueError(index)
        else:
            logger.error("Non existing shape: %s", shape)
            raise KeyError(shape)

    def get_info(self, shape, index=0):
        """
        Return dict with master shape properties

        :param shape: Str - Name of Shape
        :param index: Int - Index of Shape
        :return: tuple ( width, height )
        """
        master = self.get(shape, index)
        info_dict = dict(objectType=vC.visObjTypeUnknown, shapesCount=0, width=0, height=0, size=(0, 0))

        if master.ObjectType == vC.visObjTypeMaster and master.Shapes.Count == 1:
            info_dict["objectType"] = master.ObjectType
            info_dict["shapesCount"] = master.Shapes.Count
            info_dict["width"] = master.Shapes.Item(1).Cells("Width").ResultIU
            info_dict["height"] = master.Shapes.Item(1).Cells("Height").ResultIU
            info_dict["size"] = (lambda x: (x[2] - x[0], x[3] - x[1]))(self.get(shape, index).BoundingBox(vC.visBBoxUprightWH, 0.0, 0.0, 0.0, 0.0))
        else:
            logger.warning("Object type is not 'master', results unpredictable")
        return info_dict

    def export(self, filename, shape, index=0):
        """
        Export shape to file

        https://msdn.microsoft.com/en-us/library/office/ff769225%28v=office.14%29.aspx
        http://goo.gl/Ob7eed

        :param shape: Str
        :param index: Int
        :param filename: Str full name including file extension
        :return:
        """
        visAppSettings = vCOM.Application.Settings
        #print ApplicationSettings.RasterExportTransparencyColor
        #16777215 = white
        visAppSettings.RasterExportUseTransparencyColor = -1
        master = self.get(shape, index)
        try:
            master.Shapes.Item(1).Export(filename)
        except com_error as reason:
            logger.error("Export failed: %s", reason)
            raise

    def add_stencil(self, filename):
        """
        Load the Visio stencil file, add available shapes into the list

        Note: LIFO applies, so when there are 2 shapes with same name from 2 different stencils, the one added last will
        be used when returning the shape object.

        :param filename: name of stencil file including the extension
        :return: bool: True if succeeded, False if not

        """
        assert isinstance(filename, str), "filename must be a string"
        key = os.path.basename(os.path.splitext(filename)[0])

        if key not in self.__stencils:
            try:
                self.__stencils[key] = vCOM.Documents.OpenEx(filename, vC.visOpenDocked)
                #TODO visOpenDocked
            except com_error as reason:
                logger.error("Loading of stencil file failed for: %s", filename)
                logger.error("COM Exception: %s", reason)
                raise
            for shape in self.__stencils[key].Masters:
                if shape.NameU not in self.__shapes:
                    self.__shapes[shape.NameU] = [key]
                else:
                    self.__shapes[shape.NameU].insert(0, key)


if __name__ == "__main__":
    import doctest
    doctest.testmod()
