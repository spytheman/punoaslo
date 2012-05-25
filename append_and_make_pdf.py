#!/usr/bin/env python

import os
from os.path import abspath, isfile, splitext

import sys
import getopt

import socket  # only needed on win32-OOo3.0.0
import uno
import unohelper

from com.sun.star.uno import Exception,RuntimeException
from com.sun.star.task import ErrorCodeIOException
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

class DocumentConversionException(Exception):
    def __init__(self, message):
        print "Exception: %s" % message

def _toProperties(dict):
    props = []
    for key in dict:
        prop = PropertyValue()
        prop.Name = key
        prop.Value = dict[key]
        props.append(prop)
    return tuple(props)
    

def processfile(sourcefile, footerfile, resultfile):
    print "Converting '%(s)s' with footer '%(f)s' , to resulting PDF: '%(r)s'" % {"s":sourcefile , "f":footerfile , "r":resultfile}
    port=2002
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
    try:
        context = resolver.resolve( "uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port )
    except NoConnectException:
        raise DocumentConversionException, "failed to connect to OpenOffice.org on port %s" % port
    desktop = context.ServiceManager.createInstanceWithContext( "com.sun.star.frame.Desktop", context )
    
    sourceurl = uno.systemPathToFileUrl(os.path.abspath(sourcefile))
    footerurl = uno.systemPathToFileUrl(os.path.abspath(footerfile))
    targeturl = uno.systemPathToFileUrl(os.path.abspath(resultfile))
    docsource = desktop.loadComponentFromURL(sourceurl, "_blank", 0, _toProperties({ "Hidden": True, "FilterName": "<All formats>" }))
    docsource.Text.getEnd().insertDocumentFromURL(footerurl, _toProperties({ "Hidden": True }))
    ##docsource.Text.getEnd().setString("\nThis is the end")
    ##docsource.Text.getEnd().setString("\nFinish")
    docsource.storeToURL(targeturl, _toProperties({ "FilterName": "writer_pdf_Export" }))
    docsource.close(True)    
    context.ServiceManager

def usage():
    print "append_and_make_pdf.py --source source.odt --footer footer.odt --pdf result.pdf"
    print "     --help  : shows this usage text"


if __name__ == "__main__":
    sourcefile="source.odt"
    footerfile="footer.odt"
    resultfile="result.pdf"
    try:                                
        opts, args = getopt.getopt(sys.argv[1:], "hs:f:p:", ["help", "source=", "footer=", "pdf="])
        for opt, arg in opts:
            if opt in ("-h", "--help"):
                usage()
                sys.exit()
            elif opt in ("-s", "--source"):
                sourcefile = arg
            elif opt in ("-f", "--footer"):
                footerfile = arg
            elif opt in ("-p", "--pdf"):
                resultfile = arg
        processfile(sourcefile, footerfile, resultfile)
    except getopt.GetoptError:           
        usage()                          
        sys.exit(2) 
