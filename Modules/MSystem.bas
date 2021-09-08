Attribute VB_Name = "ModSystem"
Option Explicit
'Public Enum vbVarType
'  vbEmpty = 0       'Empty (uninitialized)
'  vbNull = 1        'Null (no valid data)
'  vbInteger = 2     'Integer
'  vbLong = 3        'Long integer
'  vbSingle = 4      'Single-precision floating-point number
'  vbDouble = 5      'Double-precision floating-point number
'  vbCurrency = 6    'Currency value
'  vbDate = 7        'Date value
'  vbString = 8      'String
'  vbObject = 9      'Object
'  vbError = 10      'Error value
'  vbBoolean = 11    'Boolean value
'  vbVariant = 12    'Variant (used only with arrays of variants)
'  vbDataObject = 13 'A data access object
'  vbDecimal = 14    'Decimal value
'  vbByte = 17       'Byte value
'  vbArray = 8192    'Array
'End Enum
Public Enum TypeCode
  TypeCode_Empty = 0
  TypeCode_Object = 1
  TypeCode_DBNull = 2
  TypeCode_Boolean = 3
  TypeCode_Char = 4
  TypeCode_SByte = 5
  TypeCode_Byte = 6
  TypeCode_Int16 = 7
  TypeCode_UInt16 = 8
  TypeCode_Int32 = 9
  TypeCode_UInt32 = 10
  TypeCode_Int64 = 11
  TypeCode_UInt64 = 12
  TypeCode_Single = 13
  TypeCode_Double = 14
  TypeCode_Decimal = 15
  TypeCode_DateTime = 16
  TypeCode_String = 18
End Enum

