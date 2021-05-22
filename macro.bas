Attribute VB_Name = "Module3"
Sub Gender()
    Dim Gender(1 To 5) As String
    Dim ttlcode As Integer
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    For i = 1 To 5
        Gender(i) = Application.InputBox(prompt:="var is", Default:="", Type:=2)
        If Gender(i) = "" Then
            ttlcode = i
            Exit For
        End If
    Next i
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1).Value = "gender"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(rc[1] = """ & Gender(1) & """, 1, if(rc[1] = """ & Gender(2) & """, 2, 3))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'    For i = 1 To 5
'        Gender(i) = Application.InputBox(Title:=i & " var", Default:="", Type:=2)
'        If Gender(i) = "" Then
'            Exit For
'        End If
'    Next i
    
End Sub

Sub Race()
    Dim Race(1 To 10) As String
    Dim ttlcode As Integer
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    For i = 1 To 15
        Race(i) = _
            Application.InputBox(prompt:="var is", Title:="ethnicity var", Default:="", Type:=2)
        If Race(i) = "" Then
            ttlcode = i
            Exit For
        End If
    Next i
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1) = "Race"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(and(" & ttlcode & " > 1, rc[1] = """ & Race(1) & """), 1, if(and(" & ttlcode & " > 2, rc[1] = """ & Race(2) & """), 2, if(and(" & ttlcode & " > 3, rc[1] = """ & Race(3) & """), 3, if(and(" & ttlcode & " > 4, rc[1] = """ & Race(4) & """), 4,if(and(" & ttlcode & " > 5, rc[1] = """ & Race(5) & """), 5,if(and(" & ttlcode & " > 6, rc[1] = """ & Race(6) & """), 6, if(and(" & ttlcode & " > 7, rc[1] = """ & Race(7) & """), 7, if(and(" & ttlcode & " > 8, rc[1] = """ & Race(8) & """), 8,if(and(" & ttlcode & " > 9, rc[1] = """ & Race(9) & """), 9,if(and(" & ttlcode & " > 10, rc[1] = """ & Race(10) & """), 10," & ttlcode & "))))))))))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Age()
    Dim Age(1 To 15) As String
    Dim ttlcode As Integer
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    For i = 1 To 15
        Age(i) = 0
    Next i
    For i = 1 To 15
        Age(i) = _
            Application.InputBox(prompt:="var is", Title:="age var", Default:=0, Type:=1)
        If Age(i) = 0 Then
            ttlcode = i
            Exit For
        End If
    Next i
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1) = "age"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(rc[1] < 18, " & ttlcode & ", if(and(" & ttlcode & " > 1, rc[1] <= " & Age(1) & "), 1, if(and(" & ttlcode & " > 2, rc[1] <= " & Age(2) & "), 2, if(and(" & ttlcode & " > 3, rc[1] <= " & Age(3) & "), 3, if(and(" & ttlcode & " > 4, rc[1] <= " & Age(4) & "), 4, if(and(" & ttlcode & " > 5, rc[1] <= " & Age(5) & "), 5, if(and(" & ttlcode & " > 6, rc[1] <= " & Age(6) & "), 6,if(and(" & ttlcode & " > 7, rc[1] <= " & Age(7) & "), 7,if(and(" & ttlcode & " > 8, rc[1] <= " & Age(8) & "), 8,if(and(" & ttlcode & " > 9, rc[1] <= " & Age(9) & "), 9,if(and(" & ttlcode & " > 10, rc[1] <= " & Age(10) & "), 10,if(and(" & ttlcode & " > 11, rc[1] <= " & Age(11) & "), 11,if(and(" & ttlcode & " > 12, rc[1] <= " & Age(12) & "), 12,if(and(" & ttlcode & " > 13, rc[1] <= " & Age(13) & "), 13,if(and(" & ttlcode & " > 14, rc[1] <= " & Age(14) & "), 14,if(and(" & ttlcode & " > 15, rc[1] <= " & Age(15) & "), 15," & ttlcode & "))))))))))))))))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Party()
    Dim Party(1 To 10) As String
    Dim ttlcode As Integer
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range

    For i = 1 To 10
    Party(i) = _
        Application.InputBox(prompt:="var is", Title:="party var", Default:="", Type:=2)
        If Party(i) = "" Then
            ttlcode = i
            Exit For
        End If
    Next i
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1).Value = "party"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(and(" & ttlcode & " > 1, rc[1] =""" & Party(1) & """), 1, if(and(" & ttlcode & " > 2, rc[1] =""" & Party(2) & """), 2, if(and(" & ttlcode & " > 3, rc[1] =""" & Party(3) & """), 3, if(and(" & ttlcode & " > 4, rc[1] =""" & Party(4) & """), 4,if(and(" & ttlcode & " > 5, rc[1] =""" & Party(5) & """), 5,if(and(" & ttlcode & " > 6, rc[1] =""" & Party(6) & """), 6,if(and(" & ttlcode & " > 7, rc[1] =""" & Party(7) & """), 7,if(and(" & ttlcode & " > 8, rc[1] =""" & Party(8) & """), 8,if(and(" & ttlcode & " > 9, rc[1] =""" & Party(9) & """), 9,if(and(" & ttlcode & " > 10, rc[1] =""" & Party(10) & """), 10, " & ttlcode & "))))))))))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Soslang()
    Dim lang(1 To 10) As String
    Dim ttlcode As Integer
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range

    For i = 1 To 10
        lang(i) = _
            Application.InputBox(prompt:="var is", Title:="lang var", Default:="", Type:=2)
        If lang(i) = "" Then
            ttlcode = i
            Exit For
        End If
    Next i
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1).Value = "lang"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(and(" & ttlcode & " > 1, rc[1] = """ & lang(1) & """), 1, if(and(" & ttlcode & " > 2, rc[1] = """ & lang(2) & """), 2, if(and(" & ttlcode & " > 3, rc[1] = """ & lang(3) & """), 3, if(and(" & ttlcode & " > 4, rc[1] = """ & lang(4) & """), 4, if(and(" & ttlcode & " > 5, rc[1] = """ & lang(5) & """), 5, if(and(" & ttlcode & " > 6, rc[1] = """ & lang(6) & """), 6, if(and(" & ttlcode & " > 7, rc[1] = """ & lang(7) & """), 7, if(and(" & ttlcode & " > 8, rc[1] = """ & lang(8) & """), 8, if(and(" & ttlcode & " > 9, rc[1] = """ & lang(9) & """), 9, if(and(" & ttlcode & " > 10, rc[1] = """ & lang(10) & """), 10, " & ttlcode & "))))))))))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Sample_Coded()
    Dim splcoded(1 To 10) As String
    Dim ttlcode As Integer
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    For i = 1 To 10
        splcoded(i) = _
            Application.InputBox(prompt:="var is", Title:="sample coded var", Default:="", Type:=2)
        If splcoded(i) = "" Then
            ttlcode = i
            Exit For
        End If
    Next i
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1).Value = "sample coded"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcrange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(and(" & ttlcode & " > 1, rc[1] = splcoded(1)), 1, if(and(" & ttlcode & " > 2, rc[1] = splcoded(2)), 2, if(and(" & ttlcode & " > 3, rc[1] = splcoded(3)), 3, if(and(" & ttlcode & " > 4, rc[1] = splcoded(4)), 4, if(and(" & ttlcode & " > 5, rc[1] = splcoded(5)), 5, if(and(" & ttlcode & " > 6, rc[1] = splcoded(6)), 6, if(and(" & ttlcode & " > 7, rc[1] = splcoded(7)), 7, if(and(" & ttlcode & " > 8, rc[1] = splcoded(8)), 8, if(and(" & ttlcode & " > 9, rc[1] = splcoded(9)), 9, if(and(" & ttlcode & " > 10, rc[1] = splcoded(10)), 10, " & ttlcode & "))))))))))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Eth_Afam()
    Dim eth(1 To 10) As String
    Dim ttlcode As Integer
    Dim blackcode As Integer
    Dim asiancode As Integer
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    blackcode = 0
    asiancode = 0
    
    For i = 1 To 10
        eth(i) = _
            Application.InputBox(prompt:="var is", Title:="eth var", Default:="", Type:=2)
        If eth(i) = "" Then
            ttlcode = i
            If blackcode = 0 Then
                blackcode = i
            End If
            If asiancode = 0 Then
                asiancode = i
            End If
            Exit For
        ElseIf eth(i) = "black" Then
            blackcode = i
        ElseIf eth(i) = "asian" Then
            asiancode = i
        End If
    Next i
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1).Value = "ethnicity"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(and(" & ttlcode & " > 1, rc[1] = """ & eth(1) & """, rc[2] = ""N""), 1, if(and(" & ttlcode & " > 2, rc[1] = """ & eth(2) & """, rc[2] = ""N""), 2, if(and(" & ttlcode & " > 3, rc[1] = """ & eth(3) & """, rc[2] = ""N""), 3, if(and(" & ttlcode & " > 4, rc[1] = """ & eth(4) & """, rc[2] = ""N""), 4, if(and(" & ttlcode & " > 5, rc[1] = """ & eth(5) & """, rc[2] = ""N""), 5, if(and(" & ttlcode & " > 6, rc[1] = """ & eth(6) & """, rc[2] = ""N""), 6, if(and(" & ttlcode & " > 7, rc[1] = """ & eth(7) & """, rc[2] = ""N""), 7, if(and(" & ttlcode & " > 8, rc[1] = """ & eth(8) & """, rc[2] = ""N""), 8, if(and(" & ttlcode & " > 9, rc[1] = """ & eth(9) & """, rc[2] = ""N""), 9, if(and(" & ttlcode & " > 10, rc[1] = """ & eth(10) & """, rc[2] = ""N""), 10, " & ttlcode & "))))))))))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Columns(Selection.Columns.Address).Insert
    Selection(1) = "ethnicity"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(and(rc[3] = ""Y"", " & blackcode & " < " & ttlcode & "), " & blackcode & ", if(and(or(rc[2] = ""c"", rc[2] = ""d"", rc[2] = ""e"", rc[2] = ""f"", rc[2] = ""k"", rc[2] = ""l"", rc[2] = ""m"", rc[2] = ""n"", rc[2] = ""u"", rc[2] = ""v"", rc[2] = ""w"", rc[2] = ""z""), rc[3] = ""n"", " & asiancode & " < " & ttlcode & "), " & asiancode & ", rc[1]))"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Selection(1).Offset(0, 1).Activate
    Selection.EntireColumn.Delete
    Selection.Offset(0, -1).Activate
End Sub

Sub Perman()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "permav"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = "=if(rc[1] = ""Y"", 1, 2)"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub ALG_Voted()
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Selection.Replace what:="absentee", replacement:=1, _
    lookat:=xlWhole, searchorder:=xlByColumns, _
    MatchCase:=False, searchformat:=False, ReplaceFormat:=False
 
Selection.Replace what:="earlyvote", replacement:=1, _
    lookat:=xlWhole, searchorder:=xlByColumns, _
    MatchCase:=False, searchformat:=False, ReplaceFormat:=False
 
Selection.Replace what:="mail", replacement:=1, _
    lookat:=xlWhole, searchorder:=xlByColumns, _
    MatchCase:=False, searchformat:=False, ReplaceFormat:=False
 
Selection.Replace what:="polling", replacement:=1, _
    lookat:=xlWhole, searchorder:=xlByColumns, _
    MatchCase:=False, searchformat:=False, ReplaceFormat:=False
 
Selection.Replace what:="unknown", replacement:=1, _
    lookat:=xlWhole, searchorder:=xlByColumns, _
    MatchCase:=False, searchformat:=False, ReplaceFormat:=False
 
Range(Selection.Rows(2), Selection.Rows(lastrow)).Replace what:="", replacement:=2, _
    lookat:=xlWhole, searchorder:=xlByColumns, _
    MatchCase:=False, searchformat:=False, ReplaceFormat:=False
 
End Sub

Sub ALG_VoteHistory()
Dim nvote As Integer
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

nvote = Application.InputBox(prompt:="nbr of vote history", _
    Title:="ttl nbr of vhistory", Type:=1)
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "vote history"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=countif(rc[-1]:rc[-" & nvote & "], 1)"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub ALG_Neighborhood()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range
Dim ncode As Integer

ncode = Application.InputBox(prompt:="ncode", _
    Title:="ttlnbr of codes", Type:=1)
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "neighborhood"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=if(" & ncode & " = 12, if(or(rc[1] = """", rc[1] = ""z""), 12, rc[1]), if(rc[1] = ""z"", 12, if(rc[1] = """", 13, rc[1])))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub ALG_Partisan()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "partisan coded"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=if(rc[1] = """", 11, if(rc[1] < 11, 1, if(rc[1] < 21, 2, if(rc[1] < 31, 3, if(rc[1] < 41, 4, if(rc[1] < 51, 5, if(rc[1] < 61, 6, if(rc[1] < 71, 7, if(rc[1]< 81, 8, if(rc[1] < 90, 9, 10))))))))))"
'sourcerange.FormulaR1C1 = _
'    "=if(rc[1] <11, 1, if(rc[1] < 21, 2, if(rc[1] < 31, 3, if(rc[1] < 41, 4, if(rc[1] < 51, 5 if(rc[1] < 61, 6, if(rc[1] < 71, 7 if(rc[1] < 81, 8, if(rc[1] < 90, 9, 10)))))))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub ALG_MediaMarket()
Dim lastrow As Long
Dim ttlcode As Integer
Dim sourcerange As Range
Dim fillrange As Range
Dim mval(1 To 10) As Variant

For i = 1 To 10
    mval(i) = _
        Application.InputBox(prompt:="Media Market Val", Title:="Media Market cell select", Default:=Selection(1).Address, Type:=8)
    If mval(i) = Selection(1).Value Then
        ttlcode = i
        Exit For
    End If
Next i
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "media market"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=if(and(" & ttlcode & " > 1, rc[1] = """ & mval(1) & """), 1, if(and(" & ttlcode & "  > 2, rc[1] = """ & mval(2) & """), 2, if(and(" & ttlcode & "  > 3, rc[1] = """ & mval(3) & """), 3, if(and(" & ttlcode & "  > 4, rc[1] = """ & mval(4) & """), 4, if(and(" & ttlcode & " > 5, rc[1] = """ & mval(5) & """), 5, if(and(" & ttlcode & " > 6, rc[1] = """ & mval(6) & """), 6, if(and(" & ttlcode & " > 7, rc[1] = """ & mval(7) & """), 7, if(and(" & ttlcode & " > 8, rc[1] = """ & mval(8) & """), 8, if(and(" & ttlcode & " > 9, rc[1] = """ & mval(9) & """), 9, if(and(" & ttlcode & " > 10, rc[1] = """ & mval(10) & """), 10, " & ttlcode & "))))))))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub ALG_VoteHistoryLogic()
Dim ttlcode As Integer
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

ttlcode = Application.InputBox(prompt:="ttl nbr of code", Title:="ttlcode", Type:=1)
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "vote history logic"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=" & ttlcode & " - countif(rc[-1]:rc[-" & ttlcode & "], """")"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub ALG_RegSince()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range
Dim regdate As String

regdate = Application.InputBox(prompt:="reg date", Title:="Registration Date", Type:=2)
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=text(rc[1], ""yyyymmdd"")"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Columns(Selection.Columns.Address).Insert
Selection(1) = "regdate since"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=if(rc[1] >= """ & regdate & """, 1, 2)"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Selection(1).Offset(0, 1).Activate
Selection.EntireColumn.Delete
Selection.Offset(0, -1).Activate
End Sub

Sub split_AB()
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1) = "Split AB"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(left(rc[1], 1) = ""A"", 1, 2)"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub split_CD()
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1) = "Split CD"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(right(left(rc[1], 2), 1) = ""C"", 1, 2)"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub split_EF()
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Selection(1) = "Split EF"
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=if(right(left(rc[1], 3), 1) = ""E"", 1, 2)"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Reg_date()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Selection.Sort Key1:=Selection(1), order1:=xlDescending
Columns(Selection.Columns.Address).Insert
Selection(1).Value = "Regdate"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)
sourcerange.FormulaR1C1 = _
    "=if(rc[1] > (r2c[1]-10000), 1, if(rc[1] > (r2c[1]-50000), 2, if(rc[1] > (r2c[1]-100000), 3, if(rc[1] >= (r2c[1]-200000), 4, 5))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Hpt_t()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert shift:=xlToLeft
Cells(1, Selection.Column) = "Hpt_t"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)

sourcerange.FormulaR1C1 = _
    "=if(or(rc[1]=0, rc[1]=1, rc[1]=5, rc[1]=6, rc[1]=8, rc[1]=9), 4, if(or(rc[1]=2, rc[1]=3, rc[1]=4, rc[1]=7, rc[1]=""C"", rc[1]=""D"", rc[1]=""E"", rc[1]=""F"", rc[1]=""G"", rc[1]=""H"", rc[1]=""I"", rc[1]=""J"", rc[1]=""K"", rc[1]=""L"", rc[1]=""M"", rc[1]=""N"", rc[1]=""O"", rc[1]=""P"",rc[1]=""Q"",rc[1]=""T"",rc[1]=""U"",rc[1]=""V"",rc[1]=""W"",rc[1]=""X"",rc[1]=""Y"",rc[1]=""Z"",), 3, if(or(rc[1]=""R"", rc[1]=""S""), 1,2)))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Ebplace()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1).Value = "ebplace"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)
sourcerange.FormulaR1C1 = _
    "=if(left(rc[1], 1)=""1"", 1, if(or(left(rc[1], 1)=""2"", left(rc[1], 1)=""3""), 2, 3))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Home()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1).Value = "home"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)

sourcerange.FormulaR1C1 = _
    "=if(rc[1] = ""H"", 1, 2)"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Gender_Age()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1).Value = "gender_age"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)

MsgBox _
    "make sure your right first column is gender coded and second is age coded"
sourcerange.FormulaR1C1 = _
    "=if(and(rc[1] = 1, or(rc[2] = 1, rc[2] = 2)), 1, if(and(rc[1] = 1, or(rc[2] = 3, rc[2] = 4)), 2, if(and(rc[1] = 2, or(rc[2] = 1, rc[2] = 2)), 3, if(and(rc[1] = 2, or(rc[2] = 3, rc[2] = 4)), 4, 5))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Party_Gender()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1).Value = "party_gender"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)

MsgBox _
    "make sure your right first column is party coded and the second is gender coded"
sourcerange.FormulaR1C1 = _
    "=if(and(rc[1] = 1, rc[2] = 1, 1, if(and(rc[1] = 1, rc[2] = 2, 2, if(and(rc[1] = 3, rc[2] = 1, 3, if(and(rc[1] = 3, rc[2] = 2, 4, if(and(rc[1] = 2, rc[2] = 1, 5, if(and(rc[1] = 2, rc[2] = 2, 6, 7))))))))))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Party_Age()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection, Columns.Address).Insert
Selection(1).Value = "party_age"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)

MsgBox _
    "make sure your right first column is party coded and the second is age coded"
sourcerange.FormulaR1C1 = _
    "=if(and(rc[1] = 1, or(rc[2] = 1, rc[2] = 2)), 1, if(and(rc[1] = 1, or(rc[2] = 3, rc[2] = 4)), 2, if(and(rc[1] = 3, or(rc[2] = 1, rc[2] = 2)), 3, if(and(rc[1] = 3, or(rc[2] = 3, rc[2] = 4)), 4, if(and(rc[1] = 2, or(rc[2] = 1, rc[2] = 2)), 5, if(and(rc[1] = 2, or(rc[2] = 3, rc[2] = 4)), 6, 7))))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Vote_Prop()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox _
    "make sure your right first is regdate and then the 5 F columns"
Selection.Sort Key1:=Selection(1), order1:=xlDescending

Columns(Selection.Columns.Address).Insert
Selection(1).Value = "vote_prop"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)
sourcerange.FormulaR1C1 = _
    "=if(countif(rc[2]:rc[6], ""N"") < 5, countif(rc[2]:rc[6], ""N"") + 1, if(rc[1] > (r2c[1] - 10000), 7, 6))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub CrossTab_Voter_Type()

Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1).Value = "voter_type"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange.Address, Selection(lastrow).Address)
MsgBox _
    "make sure your right columns in order are count A, count P, count N, permav, and all F columns"
sourcerange.FormulaR1C1 = _
    "=if(rc[4]=""Y"", 1, if(rc[1]>0, 2, if(rc[2]>0, 3, 4)))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

'Sub FM3_VotebyMail()
'Dim flags As Range
'Dim lastrow As Long
'Dim sourcerange As Range
'Dim fillrange As Range
'
'Set flags = Application.InputBox(prompt:="select the flags range", Title:="flags range selection", Type:=8)
'lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'Columns(Selection.Columns.Address).Insert
'Selection(1) = "vote by mail"
'Set sourcerange = Selection(2)
'Set fillrange = Range(sourcerange, Selection(lastrow))
'sourcerange.FormulaR1C1 = _
'    "=countif(" & flags.Address & ", ""a"")"
'End Sub

Sub ALG_PhoneType()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "phonetype"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = _
    "=if(rc[1] = ""y"", 2, 1)"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Lake_DLCCsupport()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "DLCC support"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = "=if(rc[1] >=90, 1, if(rc[1] >=80,2, if(rc[1] >=70, 3, if(rc[1] >=60, 4, if(rc[1] >=50, 5, if(rc[1] >=40, 6, if(rc[1] >=30, 7, if(rc[1] >=20, 8, if(rc[1] >=10,9, if(rc[1] >=0, 10, 11))))))))))"

'    "=if(rc[1] >= 90, 1, if(rc[1] >= 80, 2, if(rc[1] >=70, 3, if(rc[1] >= 60, 4, if(rc[1] >= 50, 5, if(rc[1] >= 40, 6, if(rc[1] >= 30, 7 if(rc[1] >= 20, 8 if(rc[1] >= 10, 9, if(rc[1] >=0, 10, 11))))))))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub Lake_DLCCturnout()
Dim lastrow As Long
Dim sourcerange As Range
Dim fillrange As Range

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Columns(Selection.Columns.Address).Insert
Selection(1) = "DLCC turnout"
Set sourcerange = Selection(2)
Set fillrange = Range(sourcerange, Selection(lastrow))
sourcerange.FormulaR1C1 = "=if(rc[1] >=90, 1, if(rc[1] >=80,2, if(rc[1] >=70, 3, if(rc[1] >=60, 4, if(rc[1] >=50, 5, if(rc[1] >=40, 6, if(rc[1] >=30, 7, if(rc[1] >=20, 8, if(rc[1] >=10,9, if(rc[1] >=0, 10, 11))))))))))"

'    "=if(rc[1] >= 90, 1, if(rc[1] >= 80, 2, if(rc[1] >=70, 3, if(rc[1] >= 60, 4, if(rc[1] >= 50, 5, if(rc[1] >= 40, 6, if(rc[1] >= 30, 7 if(rc[1] >= 20, 8 if(rc[1] >= 10, 9, if(rc[1] >=0, 10, 11))))))))))"
fillrange.FillDown
Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub MasterScrub()
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(Selection.Columns.Address).Insert
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=rc[-2] & rc[-1]"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Selection(1).Offset(, 1).Activate
    Selection.EntireColumn.Select
    Columns(Selection.Columns.Address).Insert
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'[NEW Master Scrub List.xlsx]Names'!C3,1,0)"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'fillrange.Sort Key1:=sourcerange, order1:=xlAscending
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Add Key:=Selection, Order:=xlAscending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub TextScrub()
    Dim lastrow As Long
    Dim sourcerange As Range
    Dim fillrange As Range
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Selection(1).Offset(, 1).Activate
    Selection.EntireColumn.Select
    Columns(Selection.Columns.Address).Insert
    Set sourcerange = Selection(2)
    Set fillrange = Range(sourcerange, Selection(lastrow))
    sourcerange.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'Text Scrub list.xlsx'!C2,1,0)"
    fillrange.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'fillrange.Sort Key1:=sourcerange, order1:=xlAscending
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Add Key:=Selection, Order:=xlAscending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

