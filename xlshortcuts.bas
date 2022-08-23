' vim: set ft=vb :
' Last updated: 2022-Jul-15 @ 8:39:40 AM
Option Explicit

' Shift key = "+" (plus sign)
' Ctrl key = "^" (caret)
' Alt key = "%" (percent sign)
' Enter key = "~" (numeric keypad enter = {ENTER})

Private Sub setup_shortcuts()
  Dim r As Long, tw As Worksheet
  Dim keyseq As String, proc As String
  Set tw = ThisWorkbook.sheets("Shortcuts")
  r = 2
  While tw.Range("A" & r).Value <> ""
    keyseq = tw.Range("A" & r).Value
    proc = tw.Range("B" & r).Value
    Application.OnKey keyseq, proc
    r = r + 1
  Wend
  ' Need to add these to customizations
  Application.OnKey "^%c", "toggle_center_across"           ' ctrl+alt+c
  Application.OnKey "^+r", "copy_active_range_to_clipboard" ' ctrl+shift+r
  Application.OnKey "^+%a", "set_font_gray"                 ' ctrl+shift+alt+a
  Application.OnKey "%~", "show_sheet_navigator"            ' alt+enter
End Sub

Sub Auto_Open()
  Call setup_shortcuts ' xlshortcuts
  Call set_hyperlink_macro_keys ' hyperlink
  Call reset_stack_idx
End Sub

Sub clear_status_bar(Optional clear_now As Boolean = False)
  If clear_now Then
    Application.StatusBar = False
  Else
    Application.OnTime Now + TimeValue("00:00:03"), "'clear_status_bar True'"
  End If
End Sub

Private Function style_exists(style_name As String) As Boolean
  On Error Resume Next
  Dim s As Style
  Set s = ActiveWorkbook.Styles(style_name)
  If s Is Nothing Then
    style_exists = False
  Else
    style_exists = True
  End If
End Function

Private Sub add_style(style_name As String, fmt As String)
  Dim s As Style

  If style_exists(style_name) = False Then
    ActiveWorkbook.Styles.Add style_name
  End If

  With ActiveWorkbook.Styles(style_name)
    .IncludeNumber = True
    .IncludeFont = False
    .IncludeAlignment = False ' originally TRUE
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlGeneral
    .ReadingOrder = xlContext
    .WrapText = False
    .ShrinkToFit = False
    .IncludeBorder = False
    .Locked = True
    .FormulaHidden = False
    .IncludeProtection = False
    .Interior.ColorIndex = xlNone
    .IncludePatterns = False
  End With
  ActiveWorkbook.Styles(style_name).NumberFormat = fmt
End Sub

Private Function style_needs_update(style_name As String, fmt As String) As Boolean
  style_needs_update = True
  If style_exists(style_name) Then
    If ActiveWorkbook.Styles(style_name).NumberFormat = fmt Then
      style_needs_update = False
    End If
  End If
End Function

Private Sub create_style_if_needed(ByVal style_name As String, ByVal fmt As String)
  If style_needs_update(style_name, fmt) Then
    Call add_style(style_name, fmt)
  End If
End Sub

Private Sub set_style(style_name As String, fmt As String)
  Call create_style_if_needed(style_name, fmt)
  Selection.Style = style_name
End Sub

Private Sub toggle_style(ByRef s(), default_style As String)
  On Error GoTo error_handler
  Dim n As String, fmt As String
  Dim sn As Integer
  Dim in_pt As Boolean
  Dim pf As PivotField
  ' create styles if needed
  For sn = 0 To UBound(s) - 1
    Call create_style_if_needed(s(sn)(0), s(sn)(1))
  Next sn
  in_pt = active_cell_is_in_pivot_table()
  ' now toggle
  For sn = 0 To UBound(s) - 1
    If in_pt Then
      ' if we're in a pivot table, we don't set the style/number format, but
      ' rather just set the field's number format to the one we want.
      Set pf = ActiveCell.PivotField
      If pf.NumberFormat = s(sn)(1) Or sn = (UBound(s) - 1) Then
        pf.NumberFormat = s((sn + 1) Mod UBound(s))(1)
        Exit For
      End If
    Else
      If Selection.Style.Name = s(sn)(0) Or sn = (UBound(s) - 1) Then
        Selection.Style = s((sn + 1) Mod UBound(s))(0)
        Exit For
      End If
    End If
  Next sn
  Exit Sub ' avoid error handler
error_handler:
  If Err.Number = 91 Then ' no consistent style
    Selection.Style = default_style
  Else
    MsgBox "Error: " & Err.Description, vbCritical, "Can't Continue"
  End If
End Sub

Private Sub comma_style()
  Dim s(0 To 7)
  s(0) = Array("Comma", "#,##0_);(#,##0);""–""_);@"" """)
  s(1) = Array("Zero", "0_);(0);0_);@"" """)
  s(2) = Array("Comma$", "$ #,##0_);$ (#,##0);$ ""–""_);@"" """)
  s(3) = Array("Accounting", "_($* #,##0_);_($* (#,##0);_($* ""–""_);@"" """)
  s(4) = Array("Number", "0_);(0);""–""_);@"" """)
  s(5) = Array("Number<", "$ #,##0_);$ (#,##0);$ ""–""_);@"" """)
  s(6) = Array("NegNoParen", "#,##0_);-#,##0_);""–""_);@"" """)
  Call toggle_style(s, "Comma")
  Selection.HorizontalAlignment = xlRight
End Sub

' TODO: consider taking this out. not used very often.
Private Sub comma_leftalign_style()
  ' Note: we don't include formats that use a "-" for zeros as these would not
  ' present well. We also do not include formats with a dollar sign (for now at
  ' least).

  Dim style_name As String
  style_name = Selection.Style.Name & " (Left Aligned)"

  ' Todo: can probably improve by making this dynamic instead of hard-coded.
  If style_name = "Zero (Left Aligned)" Then
    Call create_style_if_needed(style_name, "_)0;(0);_)0;"" ""@")
    Selection.Style = style_name
  Else
    GoTo nothing_to_do
  End If
  Selection.Style = style_name
  Selection.HorizontalAlignment = xlLeft

nothing_to_do:
End Sub

Private Sub thousands_style()
  Call set_style("Thousands", "#,##0,_);(#,##0,);""–""_);@"" """)
End Sub

Private Sub factor_style()
  Call set_style("Factor", "#,##0.0000_);(#,##0.0000);""–""_);@"" """)
End Sub

Private Sub percent_style()
  Call set_style("Percent", "0%_);-0%_);""–""_);@"" """)
End Sub

Private Sub toggle_date_style()
  Dim ds(0 To 6)
  ds(0) = Array("DateYearMonthDay", "yyyy-mm-dd_);;""–""_);@"" """)
  ds(1) = Array("DateMonthYear", "mmm-yyyy_);;""–""_);@"" """)
  ds(2) = Array("DateShort", "dd mmm yy_);;""–""_);@"" """)
  ds(3) = Array("DateLong", "dd mmm yyyy_);;""–""_);@"" """)
  ds(4) = Array("DateNorm", "mm-dd-yyyy_);;""–""_);@"" """)
  ds(5) = Array("DateVeryShort", "mmm'yy_);;""–""_);@"" """)
  Call toggle_style(ds, "DateYearMonthDay")
End Sub

Private Sub copy_across()
  On Error Resume Next
  Selection.Copy
  Range(Selection, Selection.End(xlToRight)).Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Selection.Calculate
End Sub

Private Sub quick_link()
  Dim formula As String, c As Range

  If ActiveWindow.SelectedSheets.Count > 1 Then
    MsgBox "More than one sheet selected can't continue.", vbInformation, "Quick Link Warning"
    GoTo quick_link_done
  End If

  Set c = ActiveCell
  If Application.CutCopyMode <> xlCopy Then
    MsgBox "Can't paste link - nothing copied", vbInformation, "Error - Can't Paste Link"
    Exit Sub
  End If
  ActiveSheet.Paste Link:=True
  For Each c In Selection
    formula = c.formula
    ' add space between equal sign and start of formula and set row anchor
    formula = "= " & Mid(formula, 2)
    formula = Application.ConvertFormula(formula, xlA1, xlA1, xlAbsRowRelColumn)
    c.formula = formula
    ' offsheet reference?
  Next c

quick_link_done:
  ' Clean up
  Application.CutCopyMode = False
End Sub

Private Sub sum_row()
  Dim actrow As Long, actcol As Long, endcol As Long
  Dim c As Range
  For Each c In Selection
    actrow = c.Row
    actcol = c.Column
    endcol = c.Offset(0, 2).End(xlToRight).Column
    Cells(actrow, actcol).FormulaR1C1 = "= SUM(RC[2]:RC[" & endcol - actcol & "])"
  Next c
End Sub

Private Sub set_font_blue()
  Selection.Font.Color = RGB(0, 0, 255)
End Sub

Private Sub toggle_font_color(colors() As Long)
  On Error GoTo error_handler
  Dim n As Integer
  For n = 0 To UBound(colors) - 1
    If Selection.Font.Color = colors(n) Or n = (UBound(colors) - 1) Then
      Selection.Font.Color = colors((n + 1) Mod UBound(colors))
      Exit For
    End If
  Next n
  Exit Sub ' avoid error handler
error_handler:
  MsgBox "Error: " & Err.Description, vbCritical, "Can't Set Background"
End Sub

Private Sub set_font_gray()
  Dim grays(0 To 5) As Long
  grays(5) = RGB(148, 163, 184)
  grays(0) = RGB(100, 116, 139)
  grays(1) = RGB(71, 85, 105)
  grays(2) = RGB(51, 65, 85)
  grays(3) = RGB(30, 41, 59)
  grays(4) = RGB(15, 23, 42)
  toggle_font_color grays
End Sub

Private Sub set_font_red()
  Selection.Font.Color = RGB(255, 0, 0)
End Sub

Private Sub set_font_auto()
  Selection.Font.ColorIndex = xlAutomatic
End Sub

Private Sub set_bg_light_yellow()
  Selection.Interior.Color = RGB(255, 255, 153)
End Sub

Private Sub set_font_white_bg_black()
  Selection.Interior.Color = RGB(0, 0, 0)
  Selection.Font.Color = RGB(255, 255, 255)
End Sub

Private Sub set_bg_light_gray()
  Dim grays(0 To 4) As Long
  ' Don't start with lightest color. Also, go from darker to lighter
  ' instead of lighter to darker. "Feels" more intuitive.
  grays(2) = RGB(248, 250, 252)
  grays(1) = RGB(241, 245, 249)
  grays(0) = RGB(226, 232, 240)
  grays(4) = RGB(203, 213, 225)
  grays(3) = RGB(148, 163, 184)
  ' too dark
  ' grays(5) = RGB(100, 116, 139)
  ' grays(6) = RGB(71, 85, 105)
  ' grays(7) = RGB(51, 65, 85)
  ' grays(8) = RGB(30, 41, 59)
  ' grays(9) = RGB(15, 23, 42)
  toggle_background_color grays
End Sub

Private Sub set_bg_yellow()
  Selection.Interior.Color = RGB(253, 245, 18)
End Sub

' Currently no shortcut
Private Sub set_bg_red()
  Selection.Interior.Color = RGB(255, 0, 0)
End Sub

' Currently no shortcut
Private Sub set_bg_green()
  Selection.Interior.Color = RGB(0, 255, 0)
End Sub

' Currently no shortcut
Private Sub set_bg_lime_yellow()
  Selection.Interior.Color = RGB(153, 204, 0)
End Sub

Private Sub set_bg_none()
  Selection.Interior.ColorIndex = xlNone
End Sub

Private Function add_zero_to_fmt(fmt As String) As String
  Dim new_fmt As String
  Dim dec_pos As Long
  Dim ins_pos As Long
  new_fmt = fmt ' default = don't do anything
  dec_pos = InStr(fmt, ".")
  If dec_pos > 0 Then
    ' note extra chars at end of std format strings need to be accounted
    ' for (i.e., not as simple as just adding a "0" on to the end).
    new_fmt = Left(fmt, dec_pos) & "0" & Right(fmt, Len(fmt) - dec_pos)
  Else
    ins_pos = InStrRev(fmt, "0")
    If ins_pos = 0 Then
      ins_pos = InStrRev(fmt, "#")
    End If
    If ins_pos > 0 Then
      new_fmt = Left(fmt, ins_pos) & ".0" & Right(fmt, Len(fmt) - ins_pos)
    End If
  End If
  add_zero_to_fmt = new_fmt
End Function

Private Function remove_zero_from_fmt(fmt As String) As String
  Dim dec_pos As Long
  Dim last_zero_pos As Long
  Dim last_pound_pos As Long
  Dim del_pos As Long
  Dim last_num_pos As Long
  Dim new_fmt As String

  new_fmt = fmt

  dec_pos = InStr(fmt, ".")
  last_zero_pos = InStrRev(fmt, "0")
  last_pound_pos = InStrRev(fmt, "#")

  del_pos = WorksheetFunction.Max(last_zero_pos, last_pound_pos)

  If dec_pos > 0 And del_pos > dec_pos Then
    new_fmt = Left(fmt, del_pos - 1) & Right(fmt, Len(fmt) - del_pos)
  End If

  dec_pos = InStr(new_fmt, ".")
  last_zero_pos = InStrRev(new_fmt, "0")
  last_pound_pos = InStrRev(new_fmt, "#")

  last_num_pos = WorksheetFunction.Max(last_zero_pos, last_pound_pos)

  If last_num_pos = dec_pos - 1 Then
    new_fmt = Left(new_fmt, dec_pos - 1) & Right(new_fmt, Len(new_fmt) - dec_pos)
  End If
  remove_zero_from_fmt = new_fmt
End Function

Private Sub fmt_change_zeros(directive As String)
  Dim break As Integer
  Dim ths As String, nxt As String, fmt As String
  Dim in_pt As Boolean
  Dim pf As PivotField
  in_pt = active_cell_is_in_pivot_table()
  fmt = ""
  ' TODO: This can throw an exception when Selection.NumberFormat is Null!
  ' (e.g., Ctrl-# on 3 full columns; Ctrl-# Ctrl-# on a full row; Try to add
  ' decimal place to full column)
  If in_pt Then
    Set pf = ActiveCell.PivotField
    ths = pf.NumberFormat
  Else
    ths = Selection.NumberFormat
  End If
  break = InStr(ths, ";")
  While break > 0
    nxt = Right(ths, Len(ths) - break)
    ths = Left(ths, break - 1)
    If fmt <> "" Then
      fmt = fmt & ";"
    End If
    If directive = "add" Then
      fmt = fmt & add_zero_to_fmt(ths)
    ElseIf directive = "remove" Then
      fmt = fmt & remove_zero_from_fmt(ths)
    End If
    break = InStr(nxt, ";")
    ths = nxt
  Wend
  If nxt <> "" Then
    If directive = "add" Then
      fmt = fmt & ";" & add_zero_to_fmt(nxt)
    ElseIf directive = "remove" Then
      fmt = fmt & ";" & remove_zero_from_fmt(nxt)
    End If
  End If
  If in_pt Then
    pf.NumberFormat = fmt
  Else
    Selection.NumberFormat = fmt
  End If
End Sub

Private Sub fmt_increase_zeros()
  ' Increase the number of decimal points for a given number
  Call fmt_change_zeros("add")
End Sub

Private Sub fmt_decrease_zeros()
  ' Decrease the number of decimal points for a given number
  Call fmt_change_zeros("remove")
End Sub

' Temp = [Temp]
Private Sub set_temporary_marker()
  Dim f As String
  Dim c As Range
  If Selection.Count > 10 Then
    Dim yn As Integer
    Dim msg As String
    msg = "You have more than 10 cells selected, this " & _
          Chr(13) & "may take a long time. Continue?"
    yn = MsgBox(msg, vbYesNo, "Continue?")
    If yn = vbNo Then Exit Sub
  End If
  For Each c In Selection
    f = c.formula
    If Left(f, 1) = "[" Or Right(f, 1) = "]" Then
      If Left(f, 1) = "[" Then
        f = Right(f, Len(f) - 1)
      End If

      If Right(f, 1) = "]" Then
        f = Left(f, Len(f) - 1)
      End If
      c.formula = f
    Else
      c.formula = "[" & f & "]"
    End If
  Next c
End Sub

Private Sub calc_selected()
  Dim r As Range
  Dim tmp
  On Error GoTo no_range
  Set r = Selection
  r.Calculate
  tmp = r.Worksheet.EnableFormatConditionsCalculation
  r.Worksheet.EnableFormatConditionsCalculation = False
  r.Worksheet.EnableFormatConditionsCalculation = True
  r.Worksheet.EnableFormatConditionsCalculation = tmp
  Exit Sub
no_range:
  MsgBox "Could not calculate range - nothing highlighted"
End Sub

Private Sub fix_pivot_defaults()
  On Error Resume Next
  Dim s As Worksheet
  Dim PT As PivotTable
  Dim pf As PivotField
  Set PT = ActiveSheet.PivotTables(1)
  PT.HasAutoFormat = False
  PT.PivotCache
  PT.PivotCache.MissingItemsLimit = xlMissingItemsNone
End Sub

Private Sub clear_values_and_contents()
  Selection.Clear
End Sub

Sub edit_shortcuts()
  Load form_help
  form_help.Show
End Sub

Sub toggle_background_color(colors() As Long)
  On Error GoTo error_handler
  Dim n As Integer
  For n = 0 To UBound(colors) - 1
    If Selection.Interior.Color = colors(n) Or n = (UBound(colors) - 1) Then
      Selection.Interior.Color = colors((n + 1) Mod UBound(colors))
      Exit For
    End If
  Next n
  Exit Sub ' avoid error handler
error_handler:
  MsgBox "Error: " & Err.Description, vbCritical, "Can't Set Background"
End Sub

Sub copy_filepath_to_clipboard()
  Call set_clipboard(ActiveWorkbook.FullName)
  Application.StatusBar = "Copied filename (full path) to clipboard"
  Call clear_status_bar
End Sub

Sub add_new_shortcut_keyseq(keyseq As String, proc As String, desc As String)
  Dim next_r As Long, tw As Worksheet
  ' Make it easy to type in from immediate window/etc.
  keyseq = Replace(keyseq, "C-", "^")
  keyseq = Replace(keyseq, "-C", "^")
  keyseq = Replace(keyseq, "S-", "+")
  keyseq = Replace(keyseq, "-S", "+")
  keyseq = Replace(keyseq, "A-", "%")
  keyseq = Replace(keyseq, "-A", "%")
  Set tw = ThisWorkbook.sheets("Shortcuts")
  next_r = tw.Range("A2").End(xlDown).Row + 1
  tw.Range("A" & next_r).Value = keyseq
  tw.Range("B" & next_r).Value = proc
  tw.Range("C" & next_r).Value = desc
  Application.OnKey keyseq, proc
End Sub

Private Function active_cell_is_in_pivot_table() As Boolean
  Dim PT As PivotTable
  On Error Resume Next
  Set PT = ActiveCell.PivotTable
  On Error GoTo 0
  active_cell_is_in_pivot_table = True
  If PT Is Nothing Then
    active_cell_is_in_pivot_table = False
  End If
End Function

Sub set_col_to_default_width()
  Selection.UseStandardWidth = True
End Sub

Sub remove_dups()
  Selection.RemoveDuplicates Columns:=1, Header:=xlNo
End Sub

Private Sub toggle_center_across()
  Dim r As Range
  Set r = Selection
  If r.HorizontalAlignment = xlHAlignCenterAcrossSelection Then
    r.HorizontalAlignment = xlHAlignGeneral
  Else
    r.HorizontalAlignment = xlHAlignCenterAcrossSelection
  End If
End Sub

Sub copy_active_range_to_clipboard()
  Call set_clipboard(Replace(Selection.Address, "$", ""))
  Application.StatusBar = "Copied range address to clipboard"
  Call clear_status_bar
End Sub

Sub reset_end_range()
  ActiveWorkbook.ActiveSheet.UsedRange.Calculate
End Sub

Sub toggle_auto_decimal()
  Dim curr As Boolean
  curr = Application.FixedDecimal
  Application.FixedDecimal = Not curr
End Sub

Private Sub show_sheet_navigator()
  Load sheet_navigator
  sheet_navigator.Show
End Sub
