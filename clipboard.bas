' vim: set ft=vb :
Option Explicit

Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
   As Long, ByVal hMem As Long) As Long
 
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Function set_clipboard(s As String) 
  Dim hGlobalMemory As Long, lpGlobalMemory As Long 
  Dim hClipMemory As Long, X As Long 

  ' Allocate moveable global memory. 
  '------------------------------------------- 
  hGlobalMemory = GlobalAlloc(GHND, Len(s) + 1) 

  ' Lock the block to get a far pointer 
  ' to this memory. 
  lpGlobalMemory = GlobalLock(hGlobalMemory) 

  ' Copy the string to this global memory. 
  lpGlobalMemory = lstrcpy(lpGlobalMemory, s) 

  ' Unlock the memory. 
  If GlobalUnlock(hGlobalMemory) <> 0 Then 
    MsgBox "Could not unlock memory location. Copy aborted." 
    GoTo OutOfHere2 
  End If 

  ' Open the Clipboard to copy data to. 
  If OpenClipboard(0&) = 0 Then 
    MsgBox "Could not open the Clipboard. Copy aborted." 
    Exit Function 
  End If 

  ' Clear the Clipboard. 
  X = EmptyClipboard() 

  ' Copy the data to the Clipboard. 
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory) 

OutOfHere2: 
  If CloseClipboard() = 0 Then 
    MsgBox "Could not close Clipboard." 
  End If 
End Function
