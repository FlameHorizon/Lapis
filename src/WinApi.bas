Attribute VB_Name = "WinApi"
Option Explicit

' Methods placed here are standard WinApi methods, for more information on how to use them, visit:
' https://docs.microsoft.com/en-us/previous-versions//aa383749(v=vs.85)
' Or:
' http://pinvoke.net/

Public Type GeneralInput
    dwType As Long
    xi(0 To 23) As Byte
End Type


Public Type KeyboardInput
    wVK As Long
    wScan As Long
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type


Private Const ModuleName As String = "WinApi"

' Methods for 64bit Windows and Office
#If VBA7 Then

    ' Methods to interact with operating system
    Public Declare PtrSafe Function SendInput Lib "user32.dll" _
    (ByVal nInputs As Long, pInputs As GeneralInput, ByVal cbSize As Long) As Long
    
    Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
    
    Public Declare PtrSafe Function GetLastError Lib "kernel32" _
    () As Long
    
    ' ANSI methods to acquire handles/names of windows
    ' Method returns handle to a window.
    Public Declare PtrSafe Function FindWindowA Lib "user32.dll" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    
    ' Method returns handle to a window child
    Public Declare PtrSafe Function FindWindowExA Lib "user32" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    
    Public Declare PtrSafe Function GetClassNameA Lib "user32" _
    (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
    Public Declare PtrSafe Function GetWindowTextA Lib "user32" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

    ' Unicode methods to acquire handles/names of windows
    ' Method returns handle to a window.
    Public Declare PtrSafe Function FindWindowW Lib "user32.dll" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    
    ' Method returns handle to a window child
    Public Declare PtrSafe Function FindWindowExW Lib "user32" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    
    Public Declare PtrSafe Function GetClassNameW Lib "user32" _
    (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
    Public Declare PtrSafe Function GetWindowTextW Lib "user32" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

    ' Same methods as above but for 32bit Windows and Office
#Else

    ' Methods to interact with operating system
    Public Declare Function SendInput Lib "user32.dll" _
                            (ByVal nInputs As Long, pInputs As GeneralInput, ByVal cbSize As Long) As Long
    
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                            (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                       (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
    
    Public Declare Function GetLastError Lib "kernel32" _
                            () As Long
    
    ' ANSI methods to acquire handles/names of windows
    ' Method returns handle to a window.
    Public Declare Function FindWindowA Lib "user32.dll" _
                            (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    ' Method returns handle to a window child
    Public Declare Function FindWindowExA Lib "user32" _
                            (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    
    Public Declare Function GetClassNameA Lib "user32" _
                            (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
    Public Declare Function GetWindowTextA Lib "user32" _
                            (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    
    ' Unicode methods to acquire handles/names of windows
    ' Method returns handle to a window.
    Public Declare Function FindWindowW Lib "user32.dll" _
                            (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    ' Method returns handle to a window child
    Public Declare Function FindWindowExW Lib "user32" _
                            (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    
    Public Declare Function GetClassNameW Lib "user32" _
                            (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
    Public Declare Function GetWindowTextW Lib "user32" _
                            (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

#End If









