# -*- coding: utf-8 -*-
"""PyVisio documents - Visio Document manipulation module"""

import logging
from visCOM import visCOMobject as vCOM
#from visCOM import visCOMconstants as vC
from visCOM import com_error

logging.info("documents loaded...")
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


class VisDocument(object):
    """
    === VisDocument ===

    Introduction
    ------------

    This library is part of PyVisio package. It's purpose is handling of new Visio documents.

    Type of objects:

    Document object - https://msdn.microsoft.com/en-us/library/ff765575.aspx

    Pages object - https://msdn.microsoft.com/EN-US/library/office/ff766165.aspx

    Page object - https://msdn.microsoft.com/EN-US/library/office/ff767035.aspx

    Usage
    -----

    Let's briefly go through the available functions.

    1. Creating/Opening document

        1. With no parameters

    New empty instance of Visio document will be created with default parameters

    >>> MyVisioDocument = VisDocument()
    >>> print MyVisioDocument #doctest:+ELLIPSIS
    <object VisDocument("") at...

        2. With filename as template

    Filename will be used as template for new document, all content of filename will be copied to new document.
    Basically it's opening new Visio from template.

    We assume current working directory contains the data directory with samples.

    >>> import os
    >>> MyVisioDocument = VisDocument(os.getcwd() + "\data\sample_document.vsd")
    >>> print MyVisioDocument #doctest:+ELLIPSIS
    <object VisDocument("...

        3. With filename and read-write flag set to true

    This is will load existing file for editing.

    Note: At this moment there are no plans to support editing Visio files as there is no business need for that,
    maybe in some next version :-)

    >>> MyVisioDocument = VisDocument(os.getcwd() + "\data\sample_document.vsd",rw=True)
    >>> print MyVisioDocument #doctest:+ELLIPSIS
    <object VisDocument("...

    of course if filename will be non existing document or there will be some problem loading the file exception
    will be raised

    >>> MyFailedVisioDocument = VisDocument("DoNotExist.vsd") #doctest:+ELLIPSIS
    Traceback (most recent call last):
    ...
    com_error: (...

    and some log lines will be written

    .. code-block: bash

    ERROR:__main__:Loading of document failed for: DoNotExist.vsd
    ERROR:__main__:COM Exception: (...

    2. Reading and modifying (basic meta) properties

    As side effect of current implementation, you can access any of the document object properties, but if you really
    need to access those you should consider extending VisDocument object or creating your own module.

    Meta properties are those available in the Properties dialog box in Visio: Title, Subject, Creator (Author),
    Manager, Company, Categories, Tags, Comments,...

    It's possible to access these properties by their name as defined in Document object (see links) same way like
    accessing VisDocument object properties.

    >>> print MyVisioDocument.Company
    Sample Company

    >>> print MyVisioDocument.Creator
    Sample Author

    You'll get exception, if property doesn't exist and you want to change it

    >>> print MyVisioDocument.NonExistingProperty
    Traceback (most recent call last):
    ...
    AttributeError: NonExistingProperty

    and log entry

    .. code-block: bash

    ERROR:__main__:Property does not exist: NonExistingProperty

    It is possible to set these properties using VisDocument.setProperty(name, value) function

    >>> MyVisioDocument.setProperty("Creator", "Myself")
    >>> print MyVisioDocument.Creator
    Myself

    You'll get exception, if property doesn't exist and you want to change it

    >>> MyVisioDocument.setProperty("Creature","Someone")
    Traceback (most recent call last):
    ...
    AttributeError: Creature

    and log entry

    .. code-block: bash

    ERROR:__main__:Property does not exist: NonExistingProperty

    3. Working with pages

    Getting number of pages in document

    >>> print len(MyVisioDocument)
    1

    Returned value is number of pages in Visio document

    Adding new page to the document

    >>> print MyVisioDocument.add()
    2

    Returned value is ID of the created page

    Adding named page to the document

    >>> print MyVisioDocument.add("TestPage")
    3

    Getting info about currently active page
    #TODO bug returned is not str but page object, when used for print it prints page.Name/NameU

    >>> print MyVisioDocument.active_page
    TestPage

    Returned is string containing the universal Name of page

    Changing active page using page name

    >>> MyVisioDocument.active_page = "Page-2"
    >>> print MyVisioDocument.active_page
    Page-2

    Iterating the page names

    >>> for i in MyVisioDocument:
    ...     print i
    Page-1
    Page-2
    TestPage

    Note: Internally Page.NameU property is used so localized names might be different. As I've no non-English version
    of Visio available, I can't test if this works correctly. Feedback from users welcome :-)

    Testing existence of page by name

    >>> if "TestPage" in MyVisioDocument:
    ...     print "TestPage exists"
    TestPage exists

    4. Accessing the pages


    To get Page object you could use VisDocument[page index or page name] returned object is Visio Page object

        4.1 By page Index

    >>> Page = MyVisioDocument[3]
    >>> print Page.NameU
    TestPage


        4.2 By page Name

    >>> Page = MyVisioDocument["TestPage"]
    >>> print Page.NameU
    TestPage

    In case wrong ID or page name will be used, exception will arrive

    >>> Page = MyVisioDocument["UnknownPage"] #doctest:+ELLIPSIS
    Traceback (most recent call last):
    ...
    com_error: (...


    Of Course page object itself is quite useless thing, but stay tuned, we'll add more functionality later on in other
    modules.

    And finally...

    4. Saving / Discarding the changes

    Important note: New document can't be saved by save() function, instead use saveAs(filename).

    So first let's see what will happen when we save file with no name.

    >>> MyVisioDocument2 = VisDocument(os.getcwd() + "\data\sample_document.vsd")
    >>> MyVisioDocument2.save()  #doctest:+ELLIPSIS
    Traceback (most recent call last):
    ...
    com_error: (...

    and of course some log line

    | ERROR:__main__:Failed to save document to file

    Correct usage of saveAs(filename). Let's save our example as sample_document1

    >>> MyVisioDocument2.saveAs(os.getcwd() + "\data\sample_document1.vsd")

    From now on, we could use save() safely

    >>> MyVisioDocument2.save()

    See, this time all went with no error

    To print document on default printer with default setting use the printDefault() method

    >>> #MyVisioDocument.printDefault() #  uncomment to kill the tree and see printout

    In case something will go wrong you'll get exception and log lines.

    To export to PDF use export() function. Default setting is following: Format=PDF, Intent=Print, Print=AllPages.
    See export() function docstring for more info. There are much more options available in Visio API, but let's stick
    for now to defaults.

    Note: This function might be unavailable in some older Visio versions.

    >>> MyVisioDocument.export(os.getcwd() + "\data\sample_document1.pdf")

    In case something will go wrong, you'll get exception and log lines.

    and finally after done with your work, you can close the file with close() function.

    By default Visio will ask what to do with changes you made if there are unsaved changes, if you do not save your
    changes intentionally you might call close(nosave=True) or simply close(True)

    >>> MyVisioDocument.close(True)

    We'll keep visio open with one of example documents we've created at the beginning, but in real program you'd
    use visCOMobject.Quit() method to close Visio instance.


    """

    def __init__(self, filename="", rw=False):
        """
        Create or open Visio file

        See https://msdn.microsoft.com/en-us/library/ff766868%28v=office.14%29.aspx

        :param filename: string - filename or ""
        :param rw: bool - if document should be opened as read-write
        """
        self._filename = filename
        try:
            if rw == False:
                self._document = vCOM.Documents.Add(filename)
            else:
                self._document = vCOM.Documents.Open(filename)
        except com_error as reason:
            logger.error("Loading of document failed for: %s", filename)
            logger.error("COM Exception: %s", reason)
            raise
        self._pages = self._document.Pages

    def __str__(self):
        """
        Return list style string representing list of loaded shapes

        :return: Str
        """
        return self.__repr__()

    def __repr__(self):
        """
        Return python style representation of object

        :return: Str
        """
        return "<object VisDocument(\"{0}\") at {1}>".format(self._filename, hex(id(self)))

    def __getattr__(self, item):
        """
        Get COM document object property

        This is kind of dirty trick (or genial idea...hard to say), __getattr__ will be called only in case other
        methods of getting property will fail. So simply said if python can't find property in VisDocument object
        it will call this function and we can inject the code which returns the Visio Document object properties

        :param item:
        :return: item value
        """
        if hasattr(self._document, item):
            return getattr(self._document, item)
        else:
            logger.error("Property does not exist: %s", item)
            raise AttributeError(item)

    def __len__(self):
        """
        Get number of pages in document

        :return: Int - Number of pages
        """
        return self._pages.Count

    def setProperty(self, name, value):
        """

        :param name: Str - Name of property
        :param value: - Value of property
        """
        if hasattr(self._document, name):
            setattr(self._document, name, value)
        else:
            logger.error("Property does not exist: %s", name)
            raise AttributeError(name)

    def add(self, name=None):
        """
        Add new page to the document

        Note: We set the AutoSize property to grow page if content will not fit to page

        :param name: Str - name of page
        :return: Int - Index of page in pages list
        """
        try:
            new_page = self._pages.Add()
        except com_error as reason:
            logger.error("Page adding failed for page: %s",  name)
            logger.error("COM Exception: %s", reason)
            raise
        new_page.AutoSize = True
        if name is not None:
            new_page.NameU = name
        return new_page.Index

    #TODO move to VisioApp (visCOM) class
    @property
    def active_page(self):
        """
        Get active page name

        :return: Str - name of active page or None
        """
        #TODO Application.ActivePage vs. ActiveWindow.Page?
        return vCOM.Application.ActiveWindow.Page

    @active_page.setter
    def active_page(self, name=None):
        """
        Set active page

        :param name: Str - name of page to focus
        """
        #TODO this is active window active page not the document active page
        #  either move away or create VisApp object where app related settings will be present
        vCOM.Application.ActiveWindow.Page = name

    def __iter__(self):
        """
        Iterate over document page names

        :return: iterator/generator
        """
        for page in self._pages:
            yield page.NameU

    def __getitem__(self, page):
        """
        Return COM object representing the page
        if page is Int then returned page will be page with Index=page
        If page is Str then returned page will be page with Name=page

        :param page: Int or Str
        :return: COM object - Page
        """
        try:
            return self._pages.ItemU(page)
        except com_error as reason:
            logger.error("Getting page '%s' failed", page)
            logger.error("COM Exception: %s", reason)
            raise

    def save(self):
        """
        Save file

        When file doesn't have yet filename raise exception (you have to use saveAs() first)
        """
        try:
            self._document.Save()
        except com_error as reason:
            logger.error("Failed to save document to file")
            logger.error("COM Exception: %s", reason)
            raise

    def saveAs(self, filename):
        """
        Save As filename

        :param filename:
        :return:
        """
        try:
            self._document.SaveAs(filename)
        except com_error as reason:
            logger.error("Failed to save document to file: %s", filename)
            logger.error("COM Exception: %s", reason)
            raise

    def printDefault(self):
        """
        Print the document on default printer

        :return:
        """
        try:
            self._document.Print()
        except com_error as reason:
            logger.error("Failed to print document")
            logger.error("COM Exception: %s", reason)
            raise

    def export(self, filename):
        """
        Export document to PDF

        :param filename: Str
        :return:
        """
        try:
            self._document.ExportAsFixedFormat(FixedFormat=1, OutputFileName=filename, Intent=1, PrintRange=0)
        except com_error as reason:
            logger.error("Failed to print document")
            logger.error("COM Exception: %s", reason)
            raise

    def close(self, nosave=False):
        """

        :param nosave: bool - True if document should be closed with no dialog about saving changes
        :return:
        """
        AlertResponse = vCOM.AlertResponse
        if nosave:
            vCOM.AlertResponse = 7
        try:
            self._document.Close()
        except com_error as reason:
            logger.error("Failed to print document")
            logger.error("COM Exception: %s", reason)
            raise
        vCOM.AlertResponse = AlertResponse


def crop_page(page):
    """
    Center drawing and crop page
    :param page: page object
    :return:
    """
    page.CenterDrawing()
    page.ResizeToFitContents()


if __name__ == "__main__":
    import doctest
    doctest.testmod()
