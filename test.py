class PythonUtilities:
    _public_methods = ['split_string']
    _reg_progid_ = "PythonUtilities"
    _reg_clsid_ = 1
    def split_string(self,val, item=None):
        if item!= None: item = str(item)
        return str(val).split(item)
#commit test
import pythoncom
print(pythoncom.CreateGuid())