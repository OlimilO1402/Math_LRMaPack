VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'diverse API-Ddeklarationen
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, ByRef lpNumberOfCharsRead As Long, lpReserved As Long) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, ByRef lpNumberOfCharsWritten As Long, lpReserved As Long) As Long

Private Const STD_INPUT_HANDLE As Long = -10&
Private Const STD_OUTPUT_HANDLE As Long = -11&
Private Const STD_ERROR_HANDLE As Long = -12&

Private Const FOREGROUND_BLUE As Long = &H1&
Private Const FOREGROUND_GREEN As Long = &H2&
Private Const FOREGROUND_RED As Long = &H4&
Private Const FOREGROUND_INTENSITY As Long = &H8&
Private Const BACKGROUND_BLUE As Long = &H10&
Private Const BACKGROUND_GREEN As Long = &H20&
Private Const BACKGROUND_RED As Long = &H40&
Private Const BACKGROUND_INTENSITY As Long = &H80&

Private Const ENABLE_LINE_INPUT As Long = &H2&
Private Const ENABLE_ECHO_INPUT As Long = &H4&
Private Const ENABLE_MOUSE_INPUT As Long = &H10&
Private Const ENABLE_PROCESSED_INPUT As Long = &H1&
Private Const ENABLE_WINDOW_INPUT As Long = &H8&

Private Const ENABLE_PROCESSED_OUTPUT As Long = &H1&
Private Const ENABLE_WRAP_AT_EOL_OUTPUT As Long = &H2&

Private hConsoleIn As Long
Private hConsoleOut As Long
Private hConsoleErr As Long
Private bInIDE As Boolean

Private Sub Class_Initialize()
  bInIDE = InIDE
  If Not bInIDE Then
    Call AllocConsole
    Call SetConsoleTitle("VB6Win32 Console .now-empowered by MBO-Ing.com")
    
    hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
    hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
    hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
  End If
End Sub
Private Function InIDE() As Boolean
TryE: On Error GoTo CatchE
  Debug.Print 1 / 0
  InIDE = False
  Exit Function
CatchE:
  InIDE = True
End Function
'Public Shared Sub SetError(ByVal newError As System.IO.TextWriter)
Public Sub SetError(ByVal newError) 'As System.IO.TextWriter)
  '
End Sub

'Public Shared Sub SetIn(ByVal newIn As System.IO.TextReader)
Public Sub SetIn(ByVal newIn) 'As System_IO.TextReader)
  '
End Sub

'Public Shared Sub SetOut(ByVal newOut As System.IO.TextWriter)
Public Sub SetOut(ByVal newOut) 'As System_IO.TextWriter)
  '
End Sub

'Public Shared Sub Write(ByVal value As Boolean)
Public Sub WWrite(ByVal varValue As Variant, Optional ByVal arg0 As Object, Optional ByVal arg1 As Object, Optional ByVal arg2 As Object)
Attribute WWrite.VB_Description = "Schreibt die Textdarstellung des angegebenen Wertes in den Standardausgabestream. Parameter:\r\nformat: Die Formatzeichenfolge.  \r\narg0: Das erste mit format zu schreibende Objekt.  \r\narg1: Das zweite mit format zu schreibende Objekt.  \r\narg2: Das dritte m"
Dim BytesWritten As Long
Dim szOut As String
  If IsObject(varValue) Then
    szOut = varValue.ToString
  Else
    szOut = CStr(varValue)
  End If
  If bInIDE Then
    Debug.Print szOut
  Else
    Call WriteConsole(hConsoleOut, szOut, Len(szOut), BytesWritten, 0&)
  End If
End Sub

'Public Shared Sub WriteLine()
Public Sub WriteLine(Optional varValue As Variant, Optional ByVal arg0 As Variant, Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant)
Dim BytesWritten As Long
Dim szOut As String
'Dim scOut As New cString
Dim szIn As String
  If IsObject(varValue) Then
    szIn = varValue.ToString
  Else
    szIn = CStr(varValue)
  End If
  If Len(szIn) > 0 Then
    'szOut = scOut.fFormat(Nothing, szIn, arg0, arg1, arg2) & vbCrLf
    If bInIDE Then
      szOut = szIn
    Else
      szOut = szIn & vbCrLf
    End If
  Else
    szOut = vbNullString & vbCrLf
  End If
  If bInIDE Then
    Debug.Print szOut
  Else
    Call WriteConsole(hConsoleOut, szOut, Len(szOut), BytesWritten, 0&)
  End If
End Sub

'Public Shared Function OpenStandardError() As System.IO.Stream
Public Function OpenStandardError(Optional ByVal bufferSize As Long) 'As System.IO.Stream
  hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
End Function

'Public Shared Function OpenStandardInput() As System.IO.Stream
Public Function OpenStandardInput(Optional ByVal bufferSize As Long) 'As System.IO.Stream
  hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
End Function

'Public Shared Function OpenStandardOutput(ByVal bufferSize As Integer) As System.IO.Stream
Public Function OpenStandardOutput(Optional ByVal bufferSize As Long) 'As System.IO.Stream
  hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
End Function

'Public Shared Function Read() As Integer
Public Function Read() As Long 'cInteger
Attribute Read.VB_Description = "Liest das nächste Zeichen aus dem Standardeingabestream.   \r\nRückgabewerte: Das nächste Zeichen aus dem Eingabestream, bzw. -1, wenn keine weiteren Zeichen verfügbar sind."
Dim sUserInput As String * 1 '256
Dim BytesRead As Long
  
  'Set Read = New cInteger
  If bInIDE Then
    'wie soll man vom Debugwindow etwas einlesen?
    Read = AscW("j")
  Else
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), BytesRead, 0&)
    ' - 2, um vbCrLf abzuschneiden
    'Read = AscW(Left$(sUserInput, BytesRead - 2))
    Read = AscW(sUserInput)
  End If
End Function

'Public Shared Function ReadLine() As String
Public Function ReadLine() As String 'cString
Attribute ReadLine.VB_Description = "Liest die nächste Zeile von Zeichen aus dem Standardeingabestream. Rückgabewerte: Die nächste Zeile aus dem Eingabestream oder null, wenn keine weiteren Zeichen verfügbar sind."
Dim sUserInput As String * 256
Dim BytesRead As Long
    
  'Set ReadLine = New cString
  If bInIDE Then
    ReadLine = "j"
  Else
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), BytesRead, 0&)
    If BytesRead > 2 Then
      ReadLine = left$(sUserInput, BytesRead - 2)
    End If
  End If
  ' - 2, um vbCrLf abzuschneiden
  
End Function

'Public Shared ReadOnly Property [Error]() As System.IO.TextWriter
Public Property Get EError() 'As System_IO.TextWriter

End Property

'Public Shared ReadOnly Property [In]() As System.IO.TextReader
Public Property Get IIn() 'As System_IO.TextReader

End Property

'Public Shared ReadOnly Property Out() As System.IO.TextWriter
Public Property Get Out() 'As System_IO.TextWriter

End Property

Private Sub Class_Terminate()
  If Not bInIDE Then FreeConsole
End Sub
