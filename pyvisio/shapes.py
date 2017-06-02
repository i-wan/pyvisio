# -*- coding: utf-8 -*-
"""PyVisio shapes - Visio Shape manipulation module"""

import logging
import os
from visCOM import visCOMobject as vCOM
from visCOM import visCOMconstants as vC
from stencils import VisStencils
from documents import VisDocument

logging.info("shapes loaded...")
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
#TODO implement correct module wide logging
#  See http://www.gossamer-threads.com/lists/python/python/863016


"""
TODO
    1. vlozit shape
2. zjisit parametry - ...
3. zmenit parametry -
    barva,
    velikost,
    pozice,
    text,
    data/cell
    1d
Pokracovat ZDE:4. smazat shape

1. vlozit konektor (a,b)
2. zjistit parametry
3. zmenit parametry
4. smazat konektor

a. grupping - kvuli pozicovani
b. vzdalenost objektu


"""


class VisShape(object):
    """
    === VisShape ===

    Introduction
    ------------

    This library is part of PyVisio package. It's purpose is handling of new Visio shape objects.

    Type of objects:

    Shape object - https://msdn.microsoft.com/EN-US/library/office/ff768546.aspx

    Cell object - https://msdn.microsoft.com/en-us/library/office/ff765137.aspx

    Usage
    -----

    To simplify manipulation with Visio shapes VisShape object is translating some functionality to more pythonic.

    1. Preparing document and loading stencils

    We'll create empty document first
    #>>> doc = VisDocument(os.getcwd() + "\data\sample_connectors.vsd",rw=True)
    >>> doc = VisDocument("")

    New document already contain Page-1 by default and it's index is 1 (and NameU should be 'Page-1')

    >>> for i in doc:
    ...     print i
    Page-1

    Indexing in Visio starts with 1 so our first (and only) page is available under index 1

    >>> #page = doc['Page-1]
    >>> page1 = doc[1]

    So we've our page where we want to draw something. Next step is to load some stencils to have shapes available
    for use

    >>> stencils = VisStencils('Basic Shapes.vss', "periph_m.vss")

    Let's use page AutoSize property, that will grow the page when we'll add shapes outside of the page.

    Note: Do not worry when page will not grow when adding shapes, seems that autosize will trigger only when
    user is interacting with Visio window so once you interact with document (add/move/remove shapes), page should
    resize.

    >>> page1.AutoSize = True

    And now we are ready to work with shapes in our new document

    2. Creating the shape

    Shape instance is created at the moment of adding it to the drawing page.

    When you omit the x and y params shape is located to position x=0, y=0

    >>> circle = VisShape(stencils['Circle'], page=page1)

    Circle was created in our document and it's center is located to (0,0).

    Yes, this is important point: Visio shape coordinates are related to shape center.
    Another important point: Visio universal units are Inches so coordinates are in Inches too
    Next important point: Page starts with (0,0) in bottom left corner, but it's possible to use negative values.

    Let's create another shape and place it's center to coordinates 1.5 inches for both x and y
    >>> rectangle = VisShape(stencils['Rectangle'], page=page1, x=1.5, y=1.5)

    3. Getting/Setting x, y, width and height

    Object coordinates are available as VisShape.x and VisShape.y

    >>> print "Circle coordinates - x: {0}, y: {1}".format(circle.x, circle.y)
    Circle coordinates - x: 0.0, y: 0.0

    Object width and height are available as VisShape.w property and VisShape.h property

    >>> print "Circle width: {0}, height: {1}".format(circle.w, circle.h)
    Circle width: 1.57480314961, height: 1.57480314961

    To get coordinates of box surrounding object (imagine rectangle which contain shape) use the VisShape.coords
    property (read only)

    >>> circle_coordinates = circle.coords
    >>> print circle_coordinates
    (-0.7874015748031495, -0.7874015748031495, 0.7874015748031495, 0.7874015748031495)

    Set (x0, y0, x1, y1) is returned

    >>> print (circle_coordinates[2] - circle_coordinates[0])
    1.57480314961

    You can see that bouncing box is a bit bigger than object itself as it's as well considering the line thickness

    Of course you could move with object and define new width or height by setting appropriate VisShape property.

    Note: Shape is dynamic object and values could be not just values but as well formulas so before setting properties
    check first in ShapeSheet if your object is not having these values calculated. This is valid virtually for every
    shape property. It's important point as changing properties could completely change shape behavior.

    To change single coordinate use VisShape.x or VisShape.y

    >>> circle.x = 2.5
    >>> circle.y = 2.5
    >>> print "Circle coordinates - x: {0}, y: {1}".format(circle.x, circle.y)
    Circle coordinates - x: 2.5, y: 2.5

    To move object in one step you could use VisShape.moveto(x,y) function

    >>> circle.moveto(3, 3)
    >>> print "Circle coordinates - x: {0}, y: {1}".format(circle.x, circle.y)
    Circle coordinates - x: 3.0, y: 3.0

    To change object width or height set new VisObject.w and/or VisObject.h
    >>> print "Rectangle width: {0}, height: {1}".format(rectangle.w, rectangle.h)
    Rectangle width: 1.57480314961, height: 1.1811023622
    >>> rectangle.w = 5
    >>> rectangle.h = 1
    >>> print "Rectangle width: {0}, height: {1}".format(rectangle.w, rectangle.h)
    Rectangle width: 5.0, height: 1.0

    4. Adding text

    Adding text is usually very simple just use VisShape.text property.

    >>> circle.text = 'Circle Text'
    >>> print circle.text
    Circle Text

    Note: Internally Visio.Shape.Text property is used and in normal situation that's perfectly OK, but in case your
    shape is having Text property calculated (based on Data1..3 or Prop or other Cells this might break standard
    shape behavior so again check ShapeSheet first (Hint: google for visio developer mode)

    5. Changing line color and background fill color

    As already mentioned before, every property of Visio Shape object is possibly function/formula so changing
    line color or fill color is not so straightforward. You can get or set current "LineColor" and "Fillforegnd"
    formula by accessing or setting VisShape.linecolor and VisShape.fillcolor, but you need to know the format how
    Visio is defining colors

    Getting current linecolor, getting current fillcolor
    >>> print circle.linecolor
    THEMEVAL()
    >>> print circle.fillcolor
    THEMEVAL()

    What we received is basically string containing formula defining color, in this case it says:
    Get the line and fill color based on current theme

    If you know how to define color formulas you can set it yourself for both linecolor and fillcolor

    >>> circle.linecolor = "=THEMEGUARD(RGB(255,255,0))"
    >>> print circle.linecolor
    THEMEGUARD(RGB(255,255,0))
    >>> circle.fillcolor = "=THEMEGUARD(THEMEVAL(\\"AccentColor4\\"))"
    >>> print circle.fillcolor
    THEMEGUARD(THEMEVAL("AccentColor4"))

    Note: Formula is starting with '='

    If really insist on changing the color to some RGB value you can use VisShape.setfillcolor() function
    >>> circle.setlinecolor(0,255,0)
    >>> print circle.linecolor
    RGB(0,255,0)
    >>> circle.setfillcolor(255,0,0)
    >>> print circle.fillcolor
    RGB(255,0,0)

    6. Other available properties/functions

    Seee source code and MS Visio documentation for details

    #TODO ma to cenu u zakladniho objektu? oned se hodi jen pro connector jinakasi moc ne
    >>> print circle.oneD
    False

    >>> print circle.cell("LineWeight")
    THEMEVAL("LineWeight",0.24 pt)

    >>> circle.setcell("LineWeight","5 pt", preformated=False)

    7. Deleting object from page

    By using VisShape.delete() you'll delete object from page. Internally it's using Shape.DeleteEx COM method so you
    can use constants available in that function if needed.

    >>> #rectangle.delete(vC.visDeleteHealConnectors)
    >>> rectangle.delete()

    Rectangle was removed from page.

    After VisShape.delete() you should remove VisShape object as it's useless. See what will happen when you'll
    try to use it.

    >>> rectangle.x  #doctest:+ELLIPSIS
    Traceback (most recent call last):
    ...
    AttributeError: 'NoneType' object...


    """
    def __init__(self, stencil, page=None, x=0.0, y=0.0):
        if page is None:
            page = vCOM.Application.ActiveWindow.Page
        #TODO when stencil is opened active window will be stencil and not document we opened before
        #  so it's not easy to know which page you want to add the shape too
        self._shape = page.Drop(stencil, float(x), float(y))

    @property
    def x(self):
        """
        Get X position
        """
        return self._shape.CellsU("PinX")

    @x.setter
    def x(self, xval):
        """
        Set X position
        """
        self._shape.CellsU("PinX").FormulaU = float(xval)

    @property
    def y(self):
        """
        Get Y position
        """
        return self._shape.CellsU("PinY")

    @y.setter
    def y(self, yval):
        """
        Set Y position
        """
        self._shape.CellsU("PinY").FormulaU = float(yval)

    @property
    def coords(self):
        """
        Get Y position
        """
        return self._shape.BoundingBox(vC.visBBoxUprightWH, 0.0, 0.0, 0.0, 0.0)

    def moveto(self, xnew, ynew):
        """
        Move to new position
        """
        self._shape.SetCenter(float(xnew), float(ynew))

    @property
    def w(self):
        """
        """
        #TODO bug returned is visio object to get float = float(str( cellsu ))
        return self._shape.CellsU("Width")

    @w.setter
    def w(self, wval):
        """
        Set width
        """
        self._shape.CellsU("Width").FormulaU = float(wval)

    @property
    def h(self):
        """
        """
        return self._shape.CellsU("Height")

    @h.setter
    def h(self, hval):
        """
        Set height
        """
        self._shape.CellsU("Height").FormulaU = float(hval)

    @property
    def text(self):
        """
        """
        return self._shape.Text

    @text.setter
    def text(self, textval):
        """
        Set text
        """
        self._shape.Text = unicode(textval)

    @property
    def linecolor(self):
        return self._shape.CellsU("LineColor").FormulaU

    @linecolor.setter
    def linecolor(self, fillformula):
        self._shape.CellsU("LineColor").FormulaU = unicode(fillformula)

    def setlinecolor(self, red=0, green=0, blue=0):
        """
        Set line color to RGB value

        Note: Surprisingly both "RGB(R,G,B)" and "=RGB(R,G,B)" works, even thou in theory only '=' version should.
        """
        self.linecolor = "RGB({0},{1},{2})".format(red, green, blue)

    @property
    def fillcolor(self):
        return self._shape.CellsU("Fillforegnd").FormulaU

    @fillcolor.setter
    def fillcolor(self, fillformula):
        self._shape.CellsU("Fillforegnd").FormulaU = unicode(fillformula)

    def setfillcolor(self, red=0, green=0, blue=0):
        """
        Set background fill color to RGB value

        Note: Surprisingly both "RGB(R,G,B)" and "=RGB(R,G,B)" works, even thou in theory only '=' version should.
        """
        self.fillcolor = "RGB({0},{1},{2})".format(red, green, blue)

    @property
    def oneD(self):
        """
        """
        return False if self._shape.OneD == 0 else True

    def cell(self, cellid):
        """
        Get cell value
        :param cellid:
        :return:
        """
        if not self._shape.CellExistsU(cellid, 0):
            raise AttributeError
        return self._shape.CellsU(cellid).FormulaU

    def setcell(self, cellid, cellvalue, preformated=False):
        if not self._shape.CellExistsU(cellid, 0):
            raise AttributeError
        if preformated:
            self._shape.CellsU(cellid).FormulaU = cellvalue
        else:
            self._shape.CellsU(cellid).FormulaU = "=\"{0}\"".format(cellvalue)

    def delete(self, flags=vC.visDeleteNormal):
        """
        See https://msdn.microsoft.com/EN-US/library/office/ff768633.aspx
        """
        self._shape.DeleteEx(flags)
        self._shape = None


class VisConnector(VisShape):
    """
    === VisConnector ===

    Introduction
    ------------

    This library is part of PyVisio package. It's purpose is handling of new Visio shape objects/connector.

    Usage
    -----

    To simplify manipulation with Visio shapes VisShape object is translating some functionality to more pythonic.

    1. Preparing document and loading stencils

    >>> doc = VisDocument("")
    >>> page2 = doc[doc.add()]
    >>> stencils = VisStencils('Basic Shapes.vss', "periph_m.vss")

    In our examples we'll use Data_Connect connector from Square Mile Systems available for free
    Link: http://www.squaremilesystems.com/products/sms-visio-utils/#download

    >>> stencils.add_stencil(os.getcwd() + os.path.normpath("/data/network_connector.vss"))

    2. Adding some shapes
    >>> bshape = stencils["Square"]
    >>> box1 = VisShape(bshape,page2, 1, 1)
    >>> box1.text = "Box1"
    >>> box2 = VisShape(bshape,page2, 1, 5)
    >>> box2.text = "Box2"
    >>> box3 = VisShape(bshape,page2, 5, 5)
    >>> box3.text = "Box3"
    >>> box4 = VisShape(bshape,page2, 5, 1)
    >>> box4.text = "Box4"

    3. Creating connector between the shapes on page

    Connector is just another VisShape object, it's basically shape with special properties related to connection.

    Either you could use 'default' dynamic connector and specify just source and destination of connector

    >>> conn1 = VisConnector(box1,box2)

    This will create connector between box1 and box2

    When creating the link you can already specify the link text

    >>> conn2 = VisConnector(box2,box3,text="Link box2 to box3")

    This will create connector between box2 and box3 with text

    If you want to use custom connector you need to specify the master shape from loaded stencil.

    >>> conn3 = VisConnector(box3,box4,stencils["Data_Connect"])

    This will create connector from defined master shape.

    You're maybe asking why to use custom connector. Reason is, that custom connector could have some fancy
    features which are not available in standard dynamic connector. In case of Data_Connect the killer feature
    is description on both ends of the shape/connector. By default description is visible and initial text is 'A'
    for start description and 'B' for end description. To simplify work with Data_Connect we've already
    implemented custom text property where start description and end description could be set by setting the
    text property with ';' as separator between start and end text. Following example shows how to use this
    feature.

    >>> conn4 = VisConnector(box4,box1,stencils["Data_Connect"],text="to Box1;to Box4")

    As VisConnector is just another VisShape object you could use same properties and functions

    Setting the text property (custom connector)

    >>> conn3.text = "to Box4;to Box3"
    >>> conn3.text = "abcd"

    Getting the X position of connector
    >>> print conn1.x
    1.0

    Getting the width of connector

    >>> print conn3.w
    0.25

    ...etc. See VisShape for further details

    >>> conn4.route_style(style=0)
    todo https://msdn.microsoft.com/EN-US/library/office/ff768312.aspx
    Determines the routing style and direction for a selected connector on the drawing page.
    ShapeRouteStyle( 1- visLORouteRightAngle | 2 - visLORouteStraight | 16 - visLORouteCenterToCenter | 9 - visLORouteNetwork

    ConLineRouteExt( 0 - default, 1 - straight, 2 - curved)
    ConLineRouteExt  CellsSRC

    """

    def __init__(self, shapeA, shapeB, connector=None, placementDir=vC.visAutoConnectDirNone, text=None):
        """
        shape.AutoConnect returns nothing (logical would be to return object of that connection)
        shape.FromConnect is collection/list of Connect objects which suprprisingly as well do not contain
        connect shape IDs/Names
        """
        conn_list = [i.FromSheet.ID for i in shapeA._shape.FromConnects]
        shapeA._shape.AutoConnect(shapeB._shape, placementDir, connector)
        conn_ID = list(set([i.FromSheet.ID for i in shapeA._shape.FromConnects]) - set(conn_list))[0]
        self._shape = shapeA._shape.ContainingPage.Shapes.ItemFromID(conn_ID)
        self._type = self._shape.NameU.split(".")[0]
        if text is not None:
            self.text = text

    @property
    def text(self):
        """
        """
        if self._type == "Data_Connect":
            descA = self._shape.Cells("Prop.Port_A_Port_Name").ResultStr(vC.visNone)
            descB = self._shape.Cells("Prop.Port_B_Port_Name").ResultStr(vC.visNone)
            return "{0};{1}".format(descA, descB)
        else:
            return self._shape.Text

    @text.setter
    def text(self, textval):
        """
        Set text
        """
        if self._type == "Data_Connect":
            # log warning
            if ';' in textval:
                descA, descB = textval.split(";")
            else:
                descA = textval
                descB = ""
            self._shape.Cells("Prop.Port_A_Port_Name").FormulaU = "=\"{0}\"".format(descA)
            self._shape.Cells("Prop.Port_B_Port_Name").FormulaU = "=\"{0}\"".format(descB)
        else:
            self._shape.Text = unicode(textval)

    def route_style(self, style=0):
        self._shape.setcell("ShapeRouteStyle", style, preformated=False)

if __name__ == "__main__":
    import doctest
    doctest.testmod()
