VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements ICloneable
'Implements ICollection
'Implements IEnumerable
'Implements IList
Private Type SABound
  cElements  As Long        'Anzahl der Elemente die je Dimension erlaubt sind
  lLBound    As Long        'Unterster Index der Dimension
End Type
Private mArr '() 'As Long ' As Variant '()
Private mpArr As Long 'Pointer to the Array
Private mpSA  As Long 'Pointer of the SafeArray-Structure
'Private Type SafeArray
Private mcDims       As Integer     ' 2 Anzahl der Dimensionen
Private mfFeatures   As Integer     ' 2 different Flags
Private mcbElements  As Long        ' 4 Anzahl an Bytes per Element
Private mcLocks      As Long        ' 4 Anzahl der Locks, wenn auf Array zugegriffen wird
Private mpvData      As Long        ' 4 Pointer zum Speicheranfang der Daten
Private mArrBounds(3) As SABound    'Speicher für 3 reservieren
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (aPtr() As Any) As Long
Private Declare Sub cpymem Lib "kernel32.dll" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal nBytes As Long)

Private Sub Class_Initialize()
  ReDim mArr(0)
End Sub
Private Sub Class_Terminate()
  If mpArr Then
    Dim pSA As Long: pSA = DeRef(mpArr)
    Dim cDim As Integer: cDim = mcDims 'GetCDim(pSA)
    'den TempSA wieder zurückkopieren,
    'um einen VB-Absturz zu vermeiden
    'Call CopySA(VarPtr(mTempSA), pSA, 16 + cDim * 8)
  End If
End Sub
'Wenn man das Array in einem Variant übergibt,
'dann läßt sicher der Typ des Arrays ermitteln
Friend Sub New_(pVarPtrArr As Long)
  mpArr = pVarPtrArr
  mpSA = DeRef(mpArr)
  'den gesamten SA-Descriptor in die Klasse kopieren
  Call cpymem(mcDims, ByVal mpSA, 16)
  'ReDim mArrBounds(mcDims)
End Sub

Public Property Let Value(index As Long, VarVal As Variant)
  mArr(index) = VarVal
End Property
Public Property Get Value(index As Long) As Variant
  Value = mArr(index)
End Property

Private Function UnsignedAdd(start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If start And &H80000000 Then 'Start < 0
      UnsignedAdd = start + Incr
   ElseIf (start Or &H80000000) < -Incr Then
      UnsignedAdd = start + Incr
   Else
      UnsignedAdd = (start + &H80000000) + (Incr + &H80000000)
   End If
   
End Function
Private Function DeRef(pArr As Long) As Long
  Call cpymem(DeRef, ByVal pArr, 4)
End Function

Public Property Get pArray() As Long
  pArray = mpArr
End Property
Public Function ToString() As String
'keine Ahnung was hier alles abgeht
'muß ma zuerst mal cdim lesen
    'Todo
End Function



'Public Shared Sub Clear(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer)
Public Sub Clear(ByVal sarray As AArray, ByVal index As Long, ByVal Length As Long)
'Legt einen Bereich von Elementen des System.Array je nach Elementtyp auf 0, false oder null fest.
Dim i As Long
  For i = index To Length
  'vielleicht auch RtlZeroMemory
    sarray(i) = vbNull 'othing
  Next
End Sub

'Public Shared Sub Copy(ByVal sourceArray As SystemArray, ByVal sourceIndex As Integer, ByVal destinationArray As SystemArray, ByVal destinationIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal sourceArray As SystemArray, ByVal sourceIndex As Long,    ByVal destinationArray As SystemArray, ByVal destinationIndex As Long, ByVal length As Long)
'Public Shared Sub Copy(ByVal sourceArray As SystemArray, ByVal destinationArray As SystemArray, ByVal length As Integer)
'Public Shared Sub Copy(ByVal sourceArray As SystemArray, ByVal destinationArray As SystemArray, ByVal length As Long)
Public Sub Copy(ByVal sourceArray As AArray, Optional ByVal sourceIndex As Long = 0, Optional ByVal destinationArray As AArray, Optional ByVal destinationIndex As Long, Optional ByVal Length As Long)
    'Todo
End Sub

Private Sub ICollection_CopyTo(ByVal sarray As AArray, ByVal index As Long)
  Call CopyTo(sarray, index)
End Sub
'Public Overridable Sub CopyTo(ByVal array As System.Array, ByVal index As Integer)
'Public Overridable Sub CopyTo(ByVal array As System.Array, ByVal index As Long)
Public Sub CopyTo(ByVal destinationArray As AArray, ByVal destinationIndex As Long)
    'Todo
End Sub

'Public Shared Sub Reverse(ByVal array As System.Array)
'Public Shared Sub Reverse(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer)
Public Sub Reverse(ByVal sarray As AArray, Optional ByVal index As Long, Optional ByVal Length As Long)
  'ungetestet
  Call MArray.Reverse(sarray.pArray, index, Length)
End Sub

'Public Sub SetValue(ByVal value As Object, ByVal index As Integer)
'Public Sub SetValue(ByVal value As Object, ByVal index1 As Integer, ByVal index2 As Integer)
'Public Sub SetValue(ByVal value As Object, ByVal index1 As Integer, ByVal index2 As Integer, ByVal index3 As Integer)
'Public Sub SetValue(ByVal value As Object, ByVal index As Long)
'Public Sub SetValue(ByVal value As Object, ByVal index1 As Long, ByVal index2 As Long)
'Public Sub SetValue(ByVal value As Object, ByVal index1 As Long, ByVal index2 As Long, ByVal index3 As Long)
'Public Sub SetValue(ByVal value As Object, ByVal ParamArray indices() As Integer)
'Public Sub SetValue(ByVal value As Object, ByVal ParamArray indices() As Long)
Public Sub SetValue(ByVal ObjValue As Variant, ByVal index1 As Long, Optional ByVal index2 As Long, Optional ByVal index3 As Long)
    'Todo
End Sub
Public Sub SetValueA(ByVal ObjValue As Object, ParamArray indices())
    'Todo
End Sub

'Public Shared Sub Sort(ByVal array As System.Array)
'Public Shared Sub Sort(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer)
'Public Shared Sub Sort(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer, ByVal comparer As System.Collections.IComparer)
'Public Shared Sub Sort(ByVal array As System.Array, ByVal comparer As System.Collections.IComparer)

'Public Shared Sub Sort(ByVal keys As System.Array, ByVal items As System.Array)
'Public Shared Sub Sort(ByVal keys As System.Array, ByVal items As System.Array, ByVal index As Integer, ByVal length As Integer)
'Public Shared Sub Sort(ByVal keys As System.Array, ByVal items As System.Array, ByVal index As Integer, ByVal length As Integer, ByVal comparer As System.Collections.IComparer)
'Public Shared Sub Sort(ByVal keys As System.Array, ByVal items As System.Array, ByVal comparer As System.Collections.IComparer)
'Public Static Sub Sort(ByVal sarray As AArray, Optional ByVal index As Long, Optional ByVal Length As Long, Optional ByVal comparer As IComparer)
'braucht Icomparer 'keine Lust machs selber
'End Sub
'Public Static Sub SortK(ByVal Keys As AArray, ByVal items As AArray, Optional ByVal index As Long, Optional ByVal Length As Long, Optional ByVal comparer As IComparer)
'braucht Icomparer 'keine Lust machs selber
'End Sub

'Public Shared Function BinarySearch(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer, ByVal value As Object) As Integer
'Public Shared Function BinarySearch(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer, ByVal value As Object, ByVal comparer As System.Collections.IComparer) As Integer
'Public Shared Function BinarySearch(ByVal array As System.Array, ByVal value As Object) As Integer
'Public Shared Function BinarySearch(ByVal array As System.Array, ByVal value As Object, ByVal comparer As System.Collections.IComparer) As Integer
'Public Function BinarySearch(ByVal sarray As AArray, Optional ByVal index As Long, Optional ByVal Length As Long, Optional ByVal ObjValue As Object, Optional ByVal comparer As IComparer) As Long
  'Set BinarySearch = New cInteger
'End Function
Public Function BinarySearch(ByVal sarray As AArray, Value As Variant) As Long
    MArray.BinarySearch sarray.pArray, VarPtr(Value)
End Function

'Public Overridable Function Clone() As Object
Private Function ICloneable_Clone() As Object
'why oh dotnet why it's so stupid
  Set ICloneable_Clone = Clone
End Function
Public Function Clone() As Object
  'Set Clone = Me 'quatsch
End Function


'Public Shared Function CreateInstance(ByVal elementType As System.Type, ByVal lengths() As Integer, ByVal lowerBounds() As Integer) As System.Array
'Public Shared Function CreateInstance(ByVal elementType As System.Type, ByVal length As Integer) As System.Array
'Public Shared Function CreateInstance(ByVal elementType As System.Type, ByVal length1 As Integer, ByVal length2 As Integer) As System.Array
'Public Shared Function CreateInstance(ByVal elementType As System.Type, ByVal length1 As Integer, ByVal length2 As Integer, ByVal length3 As Integer) As System.Array
'Public Shared Function CreateInstance(ByVal elementType As System.Type, ByVal ParamArray lengths() As Integer) As System.Array
'Public Shared Function CreateInstance(ByVal elementType As System.Type, ByVal ParamArray lengths() As Long) As System.Array


'Public Shared Function CreateInstance(ByVal elementType As System.Type, ByVal length1 As Integer, ByVal length2 As Integer, ByVal length3 As Integer) As System.Array
'Public Function CreateInstance(ByVal elementType As SystemType, ByVal length1 As Long, Optional ByVal length2 As Long, Optional ByVal length3 As Long) As SystemArray
'  Set CreateInstance = New SystemArray
'End Function

'##############  Implements IEnumerable  #################
'Private Function IEnumerable_GetEnumerator() As IUnknown 'IEnumerator
'  Set IEnumerable_GetEnumerator = GetEnumerator
'End Function
''Public Overridable Function GetEnumerator() As System.Collections.IEnumerator
'Public Function GetEnumerator() As IUnknown 'IEnumerator
'  '
'End Function
'unter VB6 braucht ein Array keinen Enumerator, um mit for each jedes Element zu durchlaufen
'bei for each wirde automatisch ein Enumerator  quasi "on the fly" erzeugt

'Public Function GetLength(ByVal dimension As Integer) As Integer
Public Function GetLength(ByVal dimension As Long) As Long
  'Set GetLength = New cInteger
End Function

'Public Function GetLowerBound(ByVal dimension As Integer) As Integer
Public Function GetLowerBound(ByVal dimension As Long) As Long
  'Set GetLowerBound = New cInteger
  'GetLowerBound = LBound(mArray, dimension)
End Function

'Public Function GetUpperBound(ByVal dimension As Integer) As Integer
Public Function GetUpperBound(ByVal dimension As Long) As Long
  'Set GetUpperBound = New cInteger
  'GetUpperBound = UBound(mArray, dimension)
End Function

'Public Function GetValue(ByVal index1 As Integer, ByVal index2 As Integer, ByVal index3 As Integer) As Object
'Public Function GetValue(ByVal index1 As Long, Optional ByVal index2 As Long, Optional ByVal index3 As Long) As Variant 'Object
  'If IsObject(mArray(index1)) Then Set GetValue = mArray(index1) Else GetValue = mArray(index1)
  'If IsObject(mArray(index2)) Then Set GetValue = mArray(index1) Else GetValue = mArray(index2)
  'If IsObject(mArray(index3)) Then Set GetValue = mArray(index1) Else GetValue = mArray(index3)
'End Function

'Public Shared Function IndexOf(ByVal array As System.Array, ByVal value As Object, ByVal startIndex As Integer, ByVal count As Integer) As Integer
'Public Function IndexOf(ByVal sarray As AArray, ByVal ObjValue As Object, Optional ByVal startIndex As Long, Optional ByVal count As Long) As Long
'  '
'End Function

'Public Shared Function LastIndexOf(ByVal array As System.Array, ByVal value As Object, ByVal startIndex As Integer, ByVal count As Integer) As Integer
'Public Function LastIndexOf(ByVal sarray As AArray, ByVal ObjValue As Object, Optional ByVal startIndex As Integer, Optional ByVal count As Integer) As long
'  '
'End Function

'Public Overridable ReadOnly Property IsFixedSize() As Boolean
'Public Property Get IsFixedSize() As Boolean
'  '
'End Property

'Private Property Get IList_IsReadOnly() As Boolean
'  Set IList_IsReadOnly = IsReadOnly
'End Property
'Public Overridable ReadOnly Property IsReadOnly() As Boolean
'Public Property Get IsReadOnly() As Boolean
'  Set IsReadOnly = New cBoolean
'End Property

'Private Property Get ICollection_IsSynchronized() As Boolean
'  Set ICollection_IsSynchronized = IsSynchronized
'End Property
'Public Overridable ReadOnly Property IsSynchronized() As Boolean
'Public Property Get IsSynchronized() As Boolean
'  Set IsSynchronized = New cBoolean
'End Property

'Private Property Get ICollection_Count() As long
'  Set ICollection_Count = Length
'End Property
'Public ReadOnly Property Length() As Integer
'Public Property Get Length() As Long
'  Set Length = New cInteger
'  Length = UBound(mArray) - LBound(mArray)
'End Property

'Public ReadOnly Property Rank() As Integer
Public Property Get Rank() As Long
  'Wieviele Dimensionen hat das Array
  Rank = mcDims
End Property

Private Property Get ICollection_SyncRoot() As Object
  Set ICollection_SyncRoot = SyncRoot
End Property
'Public Overridable ReadOnly Property SyncRoot() As Object
Public Property Get SyncRoot() As Object
  'set syncroot = new object
End Property




'##############  Implements IList        #################
Private Function IList_Add(ByVal ObjValue As Object) As Long
  'Set IList_Add = add(
End Function

Private Sub IList_Clear()
  'Call Clear  'ByVal sarray As AArray, ByVal index As Long, ByVal Length As Long)
End Sub

Private Function IList_Contains(ByVal ObjValue As Object) As Boolean
    'Todo
End Function

Private Function IList_IndexOf(ByVal ObjValue As Object) As Long
    'Todo
End Function

Private Sub IList_Insert(ByVal index As Long, ByVal ObjValue As Object)
    'Todo
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    'Todo
End Property


Private Property Set IList_Item(ByVal index As Long, RHS As Object)
    'Todo
End Property

Private Property Let IList_Item(ByVal index As Long, RHS As Object)
    'Todo
End Property

Private Property Get IList_Item(ByVal index As Long) As Object
    'Todo
End Property

Private Sub IList_Remove(ByVal ObjValue As Object)
    'Todo
End Sub

Private Sub IList_RemoveAt(ByVal index As Long)
    'Todo
End Sub


'    mArr.BinarySearch()
'    mArr.Clear()
'    mArr.Clone()
'    mArr.Copy()
'    mArr.CopyTo()
'    mArr.CreateInstance()
'    mArr.GetEnumerator()
'    mArr.GetLength()
'    mArr.GetLongLength()
'    mArr.GetLowerBound()
'    mArr.GetType()
'    mArr.GetUpperBound()
'    mArr.GetValue()
'    mArr.IndexOf()
'    mArr.Initialize()
'    mArr.IsFixedSize()
'    mArr.IsReadOnly()
'    mArr.IsSynchronized()
'    mArr.LastIndexOf()
'    mArr.Length()
'    mArr.LongLength()
'    mArr.Rank()
'    mArr.Reverse()
'    mArr.SetValue()
'    mArr.Sort()
'    mArr.SyncRoot()
