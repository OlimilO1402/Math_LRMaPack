Attribute VB_Name = "MArray"
Option Explicit
'Arrays aushebeln (mit und ohne VBoost)
Private Const FADF_AUTO        As Long = &H1
Private Const FADF_STATIC      As Long = &H2
Private Const FADF_EMBEDDED    As Long = &H4
Private Const FADF_FIXEDSIZE   As Long = &H10
Private Const FADF_RECORD      As Long = &H20
Private Const FADF_HAVEIID     As Long = &H40
Private Const FADF_HAVEVARTYPE As Long = &H80
Private Const FADF_BSTR        As Long = &H100
Private Const FADF_UNKNOWN     As Long = &H200
Private Const FADF_DISPATCH    As Long = &H400
Private Const FADF_VARIANT     As Long = &H800

Private Type SafeArray
  cDims      As Integer     ' 2 Anzahl der Dimensionen
  fFeatures  As Integer     ' 2 different Flags
  cbElements As Long        ' 4 Anzahl an Bytes per Element
  cLocks     As Long        ' 4 Anzahl der Locks, wenn auf Array zugegriffen wird
  pvData     As Long        ' 4 Pointer zum Speicheranfang der Daten
  cElements  As Long        ' 4 Anzahl der Elemente die in der ersten Dimension erlaubt sind
  lLBound    As Long        ' 4 unterster Index der ersten Dimension
End Type

Private Type SafeArrayBound
  cElements  As Long        ' 4 Anzahl der Elemente die je Dimension erlaubt sind
  lLBound    As Long        ' 4 Unterster Index der Dimension
End Type

Private Type SafeArray2D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  Bounds(0 To 1) As SafeArrayBound
End Type

Public Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Arr() As Any) As Long
'^heißt manchmal auch VarPtrArray aber das ist einfach zu lang
Public Declare Function SafeArrayCreate Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SafeArrayBound) As Long
Public Declare Function SafeArrayCreateEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SafeArrayBound, ByRef pvExtra As Any) As Long
Public Declare Function SafeArrayCreateVector Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLBound As Long, ByVal cElements As Long) As Long
Public Declare Function SafeArrayCreateVectorEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLBound As Long, ByVal cElements As Long, ByRef pvExtra As Any) As Long
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef pSA As SafeArray) As Long
Public Declare Function SafeArrayGetElemsize Lib "oleaut32.dll" (ByRef pSA As SafeArray) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (pDst As Any, pSrc As Any, ByVal nBytes As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32.dll" (Destination As Any, ByVal Length As Long)
Public Declare Sub RtlFillMemory Lib "kernel32.dll" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

'Public Shared Function BinarySearch(ByVal array As System.Array, ByVal value As Object) As Integer
Public Function BinarySearch(pVarPtrArr As Long, ByVal aPtrValue As Long) As Long
'Zusammenfassung:
' Durchsucht ein ganzes sortiertes eindimensionales System.Array nach einem
' bestimmten Element. Dazu wird die von jedem Element des System.Array und
' vom angegebenen Objekt implementierte System.IComparable-Schnittstelle verwendet.
'Parameter:
' array: Das zu durchsuchende eindimensionale System.Array.
' value: Das Objekt, nach dem gesucht werden soll.
'Rückgabewerte:
' Der Index des angegebenen value im angegebenen array, sofern value gefunden wurde.
' – oder – Eine negative Zahl, die das bitweise Komplement des Index des ersten
' Elements darstellt, das größer als value ist, wenn value nicht gefunden wurde
' und value kleiner als mindestens ein Element von array ist.
' – oder – Eine negative Zahl, die das bitweise Komplement von 1 + Index des
' letzten Elements darstellt, wenn value nicht gefunden wurde und value größer
' als alle Elemente von array ist.
'Ok wir durchsuchen den Speicherbereich
TryE: On Error GoTo CatchE
  Dim SA As SafeArray: SA = GetArrayDescriptor(pVarPtrArr)
  Dim nBCount As Long: nBCount = SA.cbElements
  Dim pE1 As Long: pE1 = SA.pvData
  Dim aPtr As Long
  For BinarySearch = 0 To SA.cElements
    aPtr = UnsignedAdd(pE1, BinarySearch * nBCount)
    If IsEqualMemory(aPtr, aPtrValue, nBCount) Then Exit Function
  Next
  'ne kein bitweises Komplement
  'einfach nur'n simples -1 für nicht gefunden
  BinarySearch = -1
  Exit Function
CatchE:
  'BinarySearch = 0
End Function

'Public Shared Sub Clear(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer)
Public Sub Clear(pVarPtrArr As Long, index As Long, Length As Long)
'Zusammenfassung:
' Legt einen Bereich von Elementen des Array je nach Elementtyp auf
' 0, false oder null fest.
'
'Parameter:
' array:  Das System.Array, dessen Elemente gelöscht werden sollen.
' index:  Der Startindex für den Bereich der zu löschenden Elemente.
' length: Die Anzahl der zu löschenden Elemente.
TryE: On Error GoTo CatchE
  Dim SA As SafeArray: SA = GetArrayDescriptor(pVarPtrArr)
  'die Anzahl der zu löschenden Bytes ermitteln
  Dim nBCount As Long: nBCount = Length * SA.cbElements
  Dim pToM As Long: pToM = UnsignedAdd(SA.pvData, index * SA.cbElements)
  Call RtlZeroMemory(ByVal pToM, nBCount)
  Exit Sub
CatchE:
  'Clear = False
End Sub


'Public Shared Sub Copy(ByVal sourceArray As System.Array, ByVal destinationArray As System.Array, ByVal length As Integer)
'Public Shared Sub Copy(ByVal sourceArray As System.Array, ByVal destinationArray As System.Array, ByVal length As Long)
Public Function Copy(VarPtrArrSrc As Long, VarPtrArrDst As Long, Length As Long) As Boolean
'Zusammenfassung:
' Kopiert einen mit dem ersten Element beginnenden Elementbereich eines
' System.Array und fügt ihn ab dem ersten Element in ein anderes System.Array ein.
' Die Länge wird als 32-Bit-Ganzzahl angegeben.
'
'Parameter:
' sourceArray:      Das System.Array, das die zu kopierenden Daten enthält.
' destinationArray: Das System.Array, das die Daten empfängt.
' length:           Eine 32-Bit-Ganzzahl, die die Anzahl der zu kopierenden Elemente
'                   darstellt.
TryE: On Error GoTo CatchE
  Dim srcSA As SafeArray: srcSA = GetArrayDescriptor(VarPtrArrSrc)
  Dim dstSA As SafeArray: dstSA = GetArrayDescriptor(VarPtrArrDst)
  'die Anzahl der zu kopierenden Bytes ermitteln
  Dim nBCount As Long
  nBCount = Length * srcSA.cbElements 'bpe
  Dim pSrcToF As Long: pSrcToF = srcSA.pvData
  Dim pDstToF As Long: pDstToF = dstSA.pvData
  'der eigentliche Kopiervorgang des Speicherbereichs
  Call RtlMoveMemory(ByVal pDstToF, ByVal pSrcToF, nBCount)
  Copy = True
  Exit Function
CatchE:
  Copy = False
End Function
Public Function GetArrayDescriptor(pArr As Long) As SafeArray
Dim pSA As Long
  'pSA = DeRef(pArr)
  RtlMoveMemory pSA, ByVal pArr, 4
  'Make sure we have a descriptor
  If pSA Then
    'Copy the SafeArray descriptor
    RtlMoveMemory GetArrayDescriptor, ByVal pSA, LenB(GetArrayDescriptor)
  End If
End Function
Public Function DeRef(pArr As Long) As Long
  RtlMoveMemory DeRef, ByVal pArr, 4
End Function
Public Function UnsignedAdd(start As Long, Incr As Long) As Long
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

Public Function CopyB(VarPtrFirstElemSrc As Long, VarPtrFirstElemDst As Long, nBLength As Long) As Boolean
'Wenn Elementgröße, und VarPtrFirstElm bekannt sind,
'dann ist diese Funktion auch möglich
'Achtung in den Funktionen des System.Array von .NET
'wird immer zuerst Source und dann Destination angegeben
'bei der Function RtlMoveMemory ist es genau andersherum
'also zuerst Destination und dann Source angeben.
'das ist meines erachtens auch näher an der Zuweisung
'( A = B : wo A:Destination, B:Source)
TryE: On Error GoTo CatchE
  Call RtlMoveMemory(ByVal VarPtrFirstElemDst, ByVal VarPtrFirstElemSrc, nBLength)
  CopyB = True
  Exit Function
CatchE:
  CopyB = False
  MsgBox "Error in CopyB"
End Function
'Public Overridable Sub CopyTo(ByVal array As System.Array, ByVal index As Integer)
'Public Overridable Sub CopyTo(ByVal array As System.Array, ByVal index As Long)

'Public Shared Sub Copy(ByVal sourceArray As System.Array, ByVal sourceIndex As Integer, ByVal destinationArray As System.Array, ByVal destinationIndex As Integer, ByVal length As Integer)
'Public Shared Sub Copy(ByVal sourceArray As System.Array, ByVal sourceIndex As Long, ByVal destinationArray As System.Array, ByVal destinationIndex As Long, ByVal length As Long)
Public Function CopyI(VarPtrArrSrc As Long, srcIndex As Long, VarPtrArrDst As Long, dstIndex As Long, Length As Long) As Boolean
'Zusammenfassung:
' Kopiert einen beim angegebenen Quellindex beginnenden Elementbereich eines
' System.Array und fügt ihn ab dem angegebenen Zielindex in ein anderes System.Array
' ein. Die Länge und die Indizes werden als 32-Bit-Ganzzahlen angegeben.
'
'Parameter:
' sourceArray:      Das System.Array, das die zu kopierenden Daten enthält.
' sourceIndex:      Eine 32-Bit-Ganzzahl, die den Index im sourceArray angibt,
'                   ab dem kopiert werden soll.
' destinationArray: Das System.Array, das die Daten empfängt.
' destinationIndex: Eine 32-Bit-Ganzzahl, die den Index im destinationArray angibt,
'                   ab dem gespeichert werden soll.
' length:           Eine 32-Bit-Ganzzahl, die die Anzahl der zu kopierenden Elemente
'                   darstellt.
TryE: On Error GoTo CatchE
  Dim srcSA As SafeArray: srcSA = GetArrayDescriptor(VarPtrArrSrc)
  Dim dstSA As SafeArray: dstSA = GetArrayDescriptor(VarPtrArrDst)
  'die Anzahl der zu kopierenden Bytes ermitteln
  Dim nBCount As Long:  nBCount = Length * srcSA.cbElements 'bpe
  Dim bytPerElm As Long: bytPerElm = srcSA.cbElements
  'die Zeiger berechnen
  Dim pSrcToF As Long:  pSrcToF = UnsignedAdd(srcSA.pvData, bytPerElm * srcIndex)
  Dim pDstToF As Long:  pDstToF = UnsignedAdd(dstSA.pvData, bytPerElm * dstIndex)
  
  'der eigentliche Kopiervorgang auf dem Speicherbereich
  Call RtlMoveMemory(ByVal pDstToF, ByVal pSrcToF, nBCount)
  
  Exit Function
CatchE:
  CopyI = False
  MsgBox "Error in CopyI"
End Function

'Public Shared Sub Reverse(ByVal array As System.Array)
'Public Shared Sub Reverse(ByVal array As System.Array, ByVal index As Integer, ByVal length As Integer)
Public Sub Reverse(pVarPtrArr As Long, Optional index As Long = 0, Optional ByVal Length As Integer = -1)
'Zusammenfassung:
' Kehrt die Reihenfolge der Elemente im gesamten eindimensionalen System.Array um.
' Kehrt die Reihenfolge der Elemente in einem Abschnitt des eindimensionalen System.Array um.
'Parameter:
' array:  Das umzukehrende eindimensionale System.Array.
' index:  Der Startindex des umzukehrenden Abschnitts.
' length: Die Anzahl der Elemente im umzukehrenden Abschnitt.
TryE: On Error GoTo CatchE
  Dim SA As SafeArray: SA = GetArrayDescriptor(pVarPtrArr)
  Dim pI0 As Long: pI0 = SA.pvData
  Dim bytPerElm As Long: bytPerElm = SA.cbElements
  ReDim TempBA(0 To SA.cbElements - 1) As Byte
  Dim pTempBAE1 As Long: pTempBAE1 = VarPtr(TempBA(0))
  Dim i As Long 'i: Schleifenzähler
  If Length < 0 Then Length = SA.cElements
  Dim pI As Long 'der Pointer von Element i
  Dim pLI As Long 'der Pointer von Element length-i
  For i = index To (Length - 1) \ 2
    'Den Pointer von Element i ermitteln
    pI = UnsignedAdd(pI0, i * bytPerElm)
    'Den Pointer von Element length-i ermitteln
    pLI = UnsignedAdd(pI0, (Length - 1 - i) * bytPerElm)
    'das Element i in den TempBA-Speicher kopieren
    Call RtlMoveMemory(ByVal pTempBAE1, ByVal pI, bytPerElm)
    'das Element Length-i in den Speicher von Element i kopieren
    Call RtlMoveMemory(ByVal pI, ByVal pLI, bytPerElm)
    'das Element im TempBA-Speicher in den Speicher von Length-i kopieren
    Call RtlMoveMemory(ByVal pLI, ByVal pTempBAE1, bytPerElm)
  Next
  Exit Sub
CatchE:
  MsgBox "Error in Reverse"
End Sub

Public Sub CopyColPtr(pArrPtrDst As Long, pVarPtrRow As Long, cElements As Long)
TryE: On Error GoTo CatchE
  Dim pSA As Long: pSA = DeRef(pArrPtrDst)
  Dim p As Long: p = UnsignedAdd(pSA, 12)
  'den Datenpointer kopieren
  Call RtlMoveMemory(ByVal p, pVarPtrRow, 4)
  'die Anzahl kopieren
  p = UnsignedAdd(p, 4)
  Call RtlMoveMemory(ByVal p, cElements, 4)
  Exit Sub
CatchE:
  MsgBox "Error in CopyRowPtr"
End Sub

Public Sub AssignArray(pArrPtrDst As Long, pArrPtrSrc As Long)
  Call RtlMoveMemory(ByVal pArrPtrDst, ByVal pArrPtrSrc, 4)
End Sub

Public Sub ZeroSAPtr(pArrPtr As Long)
  Call RtlZeroMemory(ByVal pArrPtr, 4)
End Sub

Public Sub ZeroPvData(pArrPtr As Long)
  Dim pSA As Long: pSA = DeRef(pArrPtr)
  Dim ppvData As Long: ppvData = UnsignedAdd(pSA, 12)
  Call RtlZeroMemory(ppvData, 4)
End Sub
