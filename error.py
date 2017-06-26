import sys

class ColumnError(Exception):

    def __init__(self,value,what):
        self.what    = str( what )
        self.value   = value

    def __str__(self):
        return str( self.what + ",column:" + str(self.value) )


class RowError(Exception):

    def __init__(self,message,value):
        self.message = message
        self.value   = value

    def __str__(self):
        return str( self.message + " row:" + str(self.value) )

class SheetError(Exception):

    def __init__(self,message,value):
        self.message = message
        self.value   = value

    def __str__(self):
        return str( (self.message,"sheet:" + str(self.value)) )

if sys.version_info[0] == 3:
    def raise_ex(et,trace):
        if et.__traceback__ is not trace:
            raise et.with_traceback( trace )
        raise et
else:
    exec( "def raise_ex(et,trace):\n    raise et,None,trace\n" )