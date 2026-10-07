import sys
import uno
from com.sun.star.beans import PropertyValue

def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

src, dst = sys.argv[1:3]
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
ctx = resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), '_blank', 0, (prop('Hidden', True),))
if doc is None: raise RuntimeError('Unable to open presentation')
doc.storeToURL(uno.systemPathToFileUrl(dst), (prop('FilterName', 'impress_pdf_Export'),))
doc.close(True)
print(dst)
