' vim: set ft=vb :
' Last updated: 2021-Jul-08 @ 10:30:56 AM
Option Explicit

' An Excel add-in that allows you to use Alt-RightArrow to go to the precedents
' of a particular cell and Alt-LeftArrow to go back.  This small feat can
' already be accomplished in Excel with "Ctrl-[" followed by "Ctrl-G-Enter".
' The value of this add-in is that it keeps track of your last 100 hops so that
' you can use "Alt-LeftArrow" multiple times to travel back through multiple
' formulas until you get back to the original source.  This is useful when
' links are daisy chained and you want to go to the original source then back
' to the formula you are investigating.

' Known limitations:
'   - Does not work with links to files that are not open.
'   - Clobbers the stack when more than 1 instance of Excel is running
'   - Surely more!

' % = Alt
' ^ = Control
' + = Shift
Private Const FORWARD_SHORTCUT = "%{RIGHT}"
Private Const BACKWRD_SHORTCUT = "%{LEFT}"

Private Const STACK_DEPTH = 100

Private Function get_stack_idx() As Long
  get_stack_idx = ThisWorkbook.Sheets("Tracking").Range("A1").Value
End Function

Private Function get_top_of_stack() As String
  Dim idx As Long
  idx = get_stack_idx()
  If idx = 0 Then
    get_top_of_stack = ""
  Else
    get_top_of_stack = ThisWorkbook.Sheets("Hops").Range("A" & idx)
  End If
End Function

Private Sub inc_stack_idx()
  Dim curr_stack_idx As Long
  curr_stack_idx = get_stack_idx()
  If curr_stack_idx < STACK_DEPTH Then
    ThisWorkbook.Sheets("Tracking").Range("A1").Value = curr_stack_idx + 1
  End If
End Sub

Private Sub dec_stack_idx()
  Dim curr_stack_idx As Long
  curr_stack_idx = get_stack_idx()
  If curr_stack_idx > 0 Then
    ThisWorkbook.Sheets("Tracking").Range("A1").Value = curr_stack_idx - 1
  End If
End Sub

Public Sub reset_stack_idx()
  ThisWorkbook.Sheets("Tracking").Range("A1").Value = 0
End Sub

Private Sub add_to_stack(rng As Range)
  Dim stack_ptr As Range
  Dim next_idx As Integer
  next_idx = get_stack_idx()
  If next_idx < STACK_DEPTH Then
    next_idx = next_idx + 1
  Else
    ThisWorkbook.Sheets("Hops").Range("A1:A" & STACK_DEPTH - 1).Value = _
      ThisWorkbook.Sheets("Hops").Range("A2:A" & STACK_DEPTH).Value
  End If
  Set stack_ptr = ThisWorkbook.Sheets("Hops").Range("A" & next_idx)
  ' Need 2 single quotes bc one will be swallowed by excel never to be
  ' seen again in this macro ...
  stack_ptr.Value = "'" & get_relative_address(stack_ptr, rng)
  Call inc_stack_idx
End Sub

Private Sub go_forward()
  ' Go to the first precedent and store this cell on the stack so we can go
  ' back to it later if needed.
  Dim fp As String
  fp = find_first_precedent()
  If fp <> "" Then
    Call add_to_stack(ActiveCell)
    CenterOnCell Range(find_first_precedent())
  End If
End Sub

Private Sub go_back()
  Dim prev_cell As String
  prev_cell = get_top_of_stack()
  If prev_cell = "" Then Exit Sub
  Call dec_stack_idx
  CenterOnCell Range(prev_cell)
End Sub

Private Function get_relative_address(src As Range, dest As Range) As String
  ' Returns a range address (string) for 'dest' relative to 'src'.  This can
  ' be used in gotos, etc.
  get_relative_address = ""
  If src.Worksheet.Parent.Name = dest.Worksheet.Parent.Name Then
    If src.Worksheet.Name = dest.Parent.Name Then
      ' local
      get_relative_address = Selection.Address
    Else
      get_relative_address = "'" & Selection.Parent.Name & "'!" & Selection.Address
    End If
  Else
    ' external
    get_relative_address = Selection.Address(external:=True)
  End If
End Function

' http://www.ozgrid.com/forum/showthread.php?t=17028
Private Function find_first_precedent() As String
  find_first_precedent = ""
  ' written by Bill Manville
  ' With edits from PaulS
  ' this procedure finds the cells which are the direct precedents of the active cell
  Dim rLast As Range, iLinkNum As Integer, iArrowNum As Integer
  Dim stMsg As String
  Dim bNewArrow As Boolean
  Application.ScreenUpdating = False
  ActiveCell.ShowPrecedents
  Set rLast = ActiveCell
  iArrowNum = 1
  iLinkNum = 1
  bNewArrow = True
  Do
    Do
      Application.Goto rLast
      On Error Resume Next
      ActiveCell.NavigateArrow TowardPrecedent:=True, ArrowNumber:=iArrowNum, LinkNumber:=iLinkNum
      If Err.Number > 0 Then Exit Do
      On Error GoTo 0
      If rLast.Address(external:=True) = ActiveCell.Address(external:=True) Then Exit Do
      bNewArrow = False
      stMsg = get_relative_address(rLast, ActiveCell)
      ' Only need the 1st precedent (for now)
      GoTo finish_up
      iLinkNum = iLinkNum + 1  ' try another link
    Loop
    If bNewArrow Then Exit Do
    iLinkNum = 1
    bNewArrow = True
    iArrowNum = iArrowNum + 1  'try another arrow
  Loop
finish_up:
  rLast.Parent.ClearArrows
  Application.Goto rLast
  find_first_precedent = stMsg
End Function

Public Sub set_hyperlink_macro_keys()
  Application.OnKey FORWARD_SHORTCUT, "go_forward"
  Application.OnKey BACKWRD_SHORTCUT, "go_back"
End Sub

Public Sub unset_hyperlink_macro_keys()
  Application.OnKey FORWARD_SHORTCUT
  Application.OnKey BACKWRD_SHORTCUT
End Sub


' http://stackoverflow.com/a/11943260
Function CellIsInVisibleRange(cell As Range)
  CellIsInVisibleRange = Not Intersect(ActiveWindow.VisibleRange, cell) Is Nothing
End Function

' http://www.cpearson.com/excel/zoom.htm
Sub CenterOnCell(OnCell As Range)
  Dim VisRows As Integer
  Dim VisCols As Integer

  Application.ScreenUpdating = False
  ' Switch over to the OnCell's workbook and worksheet.
  OnCell.Parent.Parent.Activate
  OnCell.Parent.Activate

  ' Get the number of visible rows and columns for the active window.
  With ActiveWindow.VisibleRange
    VisRows = .Rows.Count
    VisCols = .Columns.Count
  End With

  ' Now, determine what cell we need to GOTO. The GOTO method will
  ' place that cell reference in the upper left corner of the screen,
  ' so that reference needs to be VisRows/2 above and VisCols/2 columns
  ' to the left of the cell we want to center on. Use the MAX function
  ' to ensure we're not trying to GOTO a cell in row <=0 or column <=0.
  If CellIsInVisibleRange(OnCell) = False Then
    With Application
      .Goto reference:=OnCell.Parent.Cells( _
      .WorksheetFunction.Max(1, OnCell.Row + _
      (OnCell.Rows.Count / 2) - (VisRows / 2)), _
      .WorksheetFunction.Max(1, OnCell.Column + _
      (OnCell.Columns.Count / 2) - _
      .WorksheetFunction.RoundDown((VisCols / 2), 0))), _
      Scroll:=True
    End With
  End If

  OnCell.Select
  Application.ScreenUpdating = True
End Sub



