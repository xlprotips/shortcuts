' vim: set ft=vb :
Option Explicit

' The form in this module lets the user modify the default shortcuts that we
' set for them (e.g., change Ctrl+3 to Ctrl+n or whatever). To use it, you may
' want to add it to your "quick access toolbar" as described in the README.md.

Private Sub btn_assignsc_Click()
  Dim idx As Integer
  Dim old_key As String, old_proc As String
  Dim new_key As String, new_proc As String
  Dim tw As Worksheet
  Set tw = ThisWorkbook.Sheets(1)
  new_key  = Me.tbx_key.Value
  new_proc = Me.tbx_proc.Value
  With Me.lbox_shortcuts
    idx = .ListIndex
    old_key  = tw.Range("A" & idx + 2).Value ' .List(idx, 0)
    old_proc = tw.Range("B" & idx + 2).Value ' .List(idx, 1)
    ' Debug.Print old_proc ": " & old_key & " -> " & new_proc & ": " & new_key
    tw.Range("A" & idx + 2).Value = "'" & new_key
    tw.Range("B" & idx + 2).Value = "'" & new_proc
    .List(idx, 0) = new_key
    .List(idx, 1) = new_proc
  End With
  ThisWorkbook.Save
  Application.OnKey old_key            ' reset old key
  Application.OnKey new_key, new_proc  ' set shortcut w new key
  Me.tbx_key.Value  = ""
  Me.tbx_proc.Value = ""
End Sub

Private Sub btn_close_Click()
  Me.Hide
  Unload Me
End Sub

Private Sub lbox_shortcuts_Click()
  Dim idx As Integer
  idx = Me.lbox_shortcuts.ListIndex
  Me.tbx_key.Enabled = True
  Me.tbx_key.Value = Me.lbox_shortcuts.List(idx, 0)
  Me.tbx_proc.Enabled = True
  Me.tbx_proc.Value = Me.lbox_shortcuts.List(idx, 1)
  Me.tbx_key.SetFocus
End Sub

Private Sub UserForm_Activate()
  Dim tw As Worksheet, r As Long
  Dim keyseq As String, proc As String, desc As String
  Set tw = ThisWorkbook.Sheets(1)
  r = 2
  While tw.Range("A" & r).Value <> ""
    keyseq = tw.Range("A" & r).Value
    proc = tw.Range("B" & r).Value
    desc = tw.Range("C" & r).Value
    r = r + 1
    Me.lbox_shortcuts.AddItem keyseq
    lbox_shortcuts.List(lbox_shortcuts.ListCount - 1, 1) = proc
    lbox_shortcuts.List(lbox_shortcuts.ListCount - 1, 2) = desc
  Wend
End Sub

Private Sub tbx_key_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Dim msg As String
  Dim key As String
  Dim old As String

  old = Me.tbx_key.Value
  ' identify modifiers first
  msg = ""
  If 2 And Shift Then
    msg = "^"
  End If
  If 1 And Shift Then
    msg = msg & "+"
  End If
  If 4 And Shift Then
    msg = msg & "%"
  End If

  ' then the key code
  Select Case KeyCode
    Case 32
      key = "Space"
    Case 33
      key = "PgUp"
    Case 34
      key = "PgDown"
    Case 35
      key = "End"
    Case 36
      key = "Home"
    Case 37
      key = "Left"
    Case 38
      key = "Up"
    Case 39
      key = "Right"
    Case 40
      key = "Down"
    Case 45
      key = "Insert"
    Case 46
      key = "Delete"
    Case 48 To 57
      key = Chr(KeyCode)
    Case 65 To 90
      key = LCase(Chr(KeyCode))
    Case 112 To 123
      key = "F" & KeyCode - 111
    Case 186
      key = ";"
    Case 187
      key = "="
    Case 188
      key = ","
    Case 189
      key = "-"
    Case 190
      key = "."
    Case 191
      key = "/"
    Case 192
      key = "`"
    Case 219
      key = "["
    Case 220
      key = "\"
    Case 221
      key = "]"
    Case 222
      key = "'"
    Case 8
      key = "Backspace"
    Case 13
      key = old
    Case 9
      key = "Tab"
    Case Else
      key = ""
  End Select
  Me.tbx_key.Value = msg & key
End Sub
