Attribute VB_Name = "Module1"
Sub ex()
    Range("B9:E10").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 30
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Borders.Weight = 2
    End With
    
    Selection.Merge
    ActiveCell.FormulaR1C1 = "รายงานยอดขายเดือนมกราคม 256X"
    Range("B11").Value = "  พนักงานขาย  "
    Range("C11").Value = "  ยอดขาย  "
    Range("D11").Value = "  ค่านายหน้า(%)  "
    Range("E11").Value = "  ค่านายหน้า  "
    
    With Range("B11:E11")
        .Columns.AutoFit
        .Interior.ColorIndex = 1
        .Font.ColorIndex = 2
        .Font.Bold = True
    End With
    
End Sub


Sub remove()
    Dim Name As String
    
    i = 11
    Name = InputBox("Input Name or Type Exit:", "Input Name")
    If Name <> "" And Name <> "Exit" Then
        Do
        i = i + 1
        If Range("B" & i).Value = Name Then
            Rows(i).Select
            Selection.Delete Shift:=xlUp
            Exit Do
        End If
        Loop While 1 = 1
    ElseIf Name = "Exit" Then
        MsgBox ("Exit Input Name, Bye")
    End If
 
End Sub

Sub add()
    'Range("B5:E100").Clear
    Dim Name As String
    Dim sale As String
    
    Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Select
    sol = MsgBox("Do you want to input?", vbYesNo)
    
    Do While ActiveCell.Value = ""
        If sol = vbNo Then
            MsgBox ("Click No, Ending Program")
            Exit Do
        ElseIf sol = vbYes Then
            Name = InputBox("Input Name or Type Exit:", "Input Name")
            For i = 12 To 100
                If Name <> "" And Name <> "Exit" Then
                    If Range("B" & i).Value = Name Then
                        MsgBox ("name equal, Input Again!!")
                        Name = InputBox("Input Name or Type Exit:", "Input Name")
                    End If
                ElseIf Name = "Exit" Then
                    MsgBox ("Exit Input Name, Bye")
                    Exit Do
                End If
            Next
            sale = InputBox("Input sale :", "Input sale")
            If sale <> "" Then
                ActiveCell.Value = Name
                ActiveCell.Offset(0, 1).Value = sale
                ActiveCell.Offset(0, -1).Select
                MyCommission
            End If
        End If
        ActiveCell.Offset(0, -1).Select
        Loop
 
End Sub

Sub updatename()
    Dim oldName, newName As String
    Dim newSale As String
    
    i = 11
    oldName = InputBox("Input old Name or Type Exit:", "Input old Name")
    If oldName <> "" And oldName <> "Exit" Then
        Do
        i = i + 1
        If Range("B" & i).Value = oldName Then
            newName = InputBox("Input new Name or Type Exit:", "Input new Name")
            Range("B" & i).Value = newName
            Exit Do
        End If
        Loop While 1 = 1
    ElseIf Name = "Exit" Then
        MsgBox ("Exit Input Name, Bye")
    End If
 
End Sub

Sub updatesale()
    Dim Name As String
    Dim newSale As String

    
    i = 11
    Name = InputBox("Input Name or Type Exit:", "Input Name")
    If Name <> "" And Name <> "Exit" Then
        Do
        i = i + 1
        If Range("B" & i).Value = Name Then
            newSale = InputBox("Input  new sale or Type Exit:", "Input new Sale")
            Range("C" & i).Value = newSale
            MyCommission
            Exit Do
        End If
        Loop While 1 = 1
    ElseIf Name = "Exit" Then
        MsgBox ("Exit Input Name, Bye")
    End If
 
End Sub


Sub MyCommission()
    
    Range("C12").Select
    Do While ActiveCell.Value <> ""      'ถ้าช่อง C5 ไม่เป็นช่องว่างให้ทำการวนซ้ำ
        If ActiveCell.Value <= 10000 Then     'ถ้าในช่องมีค่ามากกว่าหรือเท่ากับ 10000
            ActiveCell.Offset(0, 1).Value = 0.02     'ให้เลื่อนไปทางขวา 1 ช่องและใส่ค่า 0.02
            ActiveCell.Offset(0, 2).Value = ActiveCell.Value * 0.02      'ให้เลื่อนไปทางขวา 2 ช่องและใส่ค่าในช่องที่ active * 0.02
        ElseIf ActiveCell.Value <= 20000 Then
            ActiveCell.Offset(0, 1).Value = 0.03
            ActiveCell.Offset(0, 2).Value = ActiveCell.Value * 0.03
        ElseIf ActiveCell.Value <= 40000 Then
            ActiveCell.Offset(0, 1).Value = 0.05
            ActiveCell.Offset(0, 2).Value = ActiveCell.Value * 0.05
        Else
            ActiveCell.Offset(0, 1).Value = 0.05
            ActiveCell.Offset(0, 2).Value = ActiveCell.Value * 0.05
        End If
        ActiveCell.Offset(0, 1).Select
        Selection.NumberFormat = "0.00%"
        ActiveCell.Offset(0, 1).Select
        Selection.NumberFormat = "#,##0.00"
        ActiveCell.Offset(1, -2).Select
    Loop
    
End Sub



