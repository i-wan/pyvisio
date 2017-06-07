# -*- coding: utf-8 -*-
"""PyVisio visCOM - Visio COM object loader"""

import logging
import win32com.client

if win32com.client.gencache.is_readonly:
    win32com.client.gencache.is_readonly = False
    win32com.client.gencache.Rebuild()

from win32com.client.gencache import EnsureDispatch
from win32com.client import constants as visCOMconstants
from pythoncom import com_error

logging.info("visCOM loaded...")
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

visCOMobject = None

#TODO move to class possible singleton
try:
    visCOMobject = EnsureDispatch("Visio.Application")
except com_error as details:
    logger.error("Exception: {0}".format(details[1]))
    logger.error("It was not possible to initiate Visio COM object! Exiting...")
    raise

visCOMobject.Visible = True

#  https://msdn.microsoft.com/en-us/library/office/ff767782.aspx
#  visCOMobject.AlertResponse = 7
