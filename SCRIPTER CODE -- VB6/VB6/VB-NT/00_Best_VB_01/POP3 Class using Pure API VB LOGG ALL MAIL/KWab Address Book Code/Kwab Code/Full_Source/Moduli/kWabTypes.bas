Attribute VB_Name = "modkWabTypes"
Option Explicit

Public Enum kPropTypes
   PT_UNSPECIFIED = 0&  '(Reserved for interface use) type doesn't matter to caller
   PT_NULL = 1&         'NULL property value
   PT_SHORT = 2&        'Signed 16-bit value
   PT_LONG = 3&         'Signed 32-bit value
   PT_FLOAT = 4&        '4-byte floating point
   PT_DOUBLE = 5&       'Floating point double
   PT_BOOLEAN = 11&     '16-bit boolean (non-zero true)
   PT_CURRENCY = 6&     'Signed 64-bit int (decimal w/ 4 digits right of decimal pt)
   PT_APPTIME = 7&      'Application time
   PT_SYSTIME = 64&     '64-bit int w/ number of 100ns periods since Jan 1,1601
   PT_STRING8 = 30&     'Null terminated 8-bit character string
   PT_BINARY = 258&     'Uninterpreted (counted byte array)
   PT_UNICODE = 31&     'Null terminated Unicode string
   PT_CLSID = 72&       'OLE GUID
   PT_LONGLONG = 20&    '8-byte signed integer
   PT_MV_SHORT = (&H1000& Or PT_SHORT)
   PT_MV_LONG = (&H1000& Or PT_LONG)
   PT_MV_FLOAT = (&H1000& Or PT_FLOAT)
   PT_MV_DOUBLE = (&H1000& Or PT_DOUBLE)
   PT_MV_CURRENCY = (&H1000& Or PT_CURRENCY)
   PT_MV_APPTIME = (&H1000& Or PT_APPTIME)
   PT_MV_SYSTIME = (&H1000& Or PT_SYSTIME)
   PT_MV_BINARY = (&H1000& Or PT_BINARY)
   PT_MV_STRING8 = (&H1000& Or PT_STRING8)
   PT_MV_UNICODE = (&H1000& Or PT_UNICODE)
   PT_MV_CLSID = (&H1000& Or PT_CLSID)
   PT_MV_LONGLONG = (&H1000& Or PT_LONGLONG)
   PT_ERROR = 10&    '32-bit error value
   PT_OBJECT = 13&   'Embedded object in a property
End Enum

Public Function kWabPropTypeName(ByVal Code As kPropTypes) As String
   Select Case Code
      Case PT_UNSPECIFIED
         kWabPropTypeName = "Reserved type"
      Case PT_NULL
         kWabPropTypeName = "NULL"
      Case PT_SHORT
         kWabPropTypeName = "Short"
      Case PT_LONG
         kWabPropTypeName = "Long"
      Case PT_FLOAT
         kWabPropTypeName = "Float"
      Case PT_DOUBLE
         kWabPropTypeName = "Double"
      Case PT_BOOLEAN
         kWabPropTypeName = "Boolean"
      Case PT_CURRENCY
         kWabPropTypeName = "Currency"
      Case PT_APPTIME
         kWabPropTypeName = "Application time"
      Case PT_SYSTIME
         kWabPropTypeName = "SysTime"
      Case PT_STRING8
         kWabPropTypeName = "Z-string (8bit)"
      Case PT_BINARY
         kWabPropTypeName = "Counted byte array"
      Case PT_UNICODE
         kWabPropTypeName = "Z-string (Unicode)"
      Case PT_CLSID
         kWabPropTypeName = "CLSID (OLE GUID)"
      Case PT_LONGLONG
         kWabPropTypeName = "LONGLONG"
      Case PT_MV_LONG
         kWabPropTypeName = "MV Long"
      Case PT_MV_SHORT
         kWabPropTypeName = "MV Short"
      Case PT_MV_FLOAT
         kWabPropTypeName = "MV Float"
      Case PT_MV_DOUBLE
         kWabPropTypeName = "MV Double"
      Case PT_MV_CURRENCY
         kWabPropTypeName = "MV Currency"
      Case PT_MV_APPTIME
         kWabPropTypeName = "MV Application time"
      Case PT_MV_SYSTIME
         kWabPropTypeName = "MV SysTime"
      Case PT_MV_BINARY
         kWabPropTypeName = "MV Counted byte array"
      Case PT_MV_STRING8
         kWabPropTypeName = "MV Z-string (8bit)"
      Case PT_MV_UNICODE
         kWabPropTypeName = "MV Z-string (Unicode)"
      Case PT_MV_CLSID
         kWabPropTypeName = "MV CLSID (OLE GUID)"
      Case PT_MV_LONGLONG
         kWabPropTypeName = "MV LONGLONG"
      Case PT_ERROR
         kWabPropTypeName = "32-bit error value"
      Case PT_OBJECT
         kWabPropTypeName = "Embedded object"
      Case Else
         kWabPropTypeName = "<unknown code>"
   End Select
End Function

