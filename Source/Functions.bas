Attribute VB_Name = "Functions"
Option Explicit


'move the table to left
Function MoveObjectLeft()

    If ObjectName = "Table" Then
        
        'left move restrictions
        If Editor.TableSelector(ObjectIndex).Left > 0.07 Then
            If ObjectName = "Table" Then
                Editor.Table(ObjectIndex).Left = Editor.Table(ObjectIndex).Left - 0.05
                Editor.TableShadow(ObjectIndex).Left = Editor.TableShadow(ObjectIndex).Left - 0.05
                Editor.TableSelector(ObjectIndex).Left = Editor.TableSelector(ObjectIndex).Left - 0.05
                Editor.Shape.Left = Editor.Shape.Left - 0.05
                Editor.TableName(ObjectIndex).Left = Editor.TableName(ObjectIndex).Left - 0.05
            End If
            
            ' move fields to left
            For i = 0 To Editor.FieldName.Count - 1
                If Editor.FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then

                    ' move field
                    Editor.FieldName(i).Left = Editor.FieldName(i).Left - 0.05
                    Editor.FieldType(i).Left = Editor.FieldType(i).Left - 0.05
                End If
            Next i
            
        End If
        MoveMovers
        MoveRelations
        ProjectModified
    End If

End Function

'move the table to right
Function MoveObjectRight()
        
    If ObjectName = "Table" Then
        'right move restrictions
        If Editor.TableSelector(ObjectIndex).Left < ((Editor.EditField.Width - Editor.Table(ObjectIndex).Width) - 0.15) Then
            If ObjectName = "Table" Then
                Editor.Table(ObjectIndex).Left = Editor.Table(ObjectIndex).Left + 0.05
                Editor.TableShadow(ObjectIndex).Left = Editor.TableShadow(ObjectIndex).Left + 0.05
                Editor.TableSelector(ObjectIndex).Left = Editor.TableSelector(ObjectIndex).Left + 0.05
                Editor.Shape.Left = Editor.Shape.Left + 0.05
                Editor.TableName(ObjectIndex).Left = Editor.TableName(ObjectIndex).Left + 0.05
            End If
            
            ' move fields to right
            For i = 0 To Editor.FieldName.Count - 1
                If Editor.FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                    ' move field
                    Editor.FieldName(i).Left = Editor.FieldName(i).Left + 0.05
                    Editor.FieldType(i).Left = Editor.FieldType(i).Left + 0.05
                End If
            Next i
            
        End If
        MoveMovers
        MoveRelations
        ProjectModified
    End If
        
End Function

'move the table up
Function MoveObjectUp()

    If ObjectName = "Table" Then
        'up move restrictions
        If Editor.TableSelector(ObjectIndex).Top > 0.07 Then
            
            'move table
            Editor.Table(ObjectIndex).Top = Editor.Table(ObjectIndex).Top - 0.05
            Editor.TableShadow(ObjectIndex).Top = Editor.TableShadow(ObjectIndex).Top - 0.05
            Editor.TableSelector(ObjectIndex).Top = Editor.TableSelector(ObjectIndex).Top - 0.05
            Editor.Shape.Top = Editor.Shape.Top - 0.05
            Editor.TableName(ObjectIndex).Top = Editor.TableName(ObjectIndex).Top - 0.05
            
            ' move fields up
            For i = 0 To Editor.FieldName.Count - 1
                If Editor.FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                    ' move field
                    Editor.FieldName(i).Top = Editor.FieldName(i).Top - 0.05
                    Editor.FieldType(i).Top = Editor.FieldType(i).Top - 0.05
                End If
            Next i
        End If
        MoveMovers
        MoveRelations
        ProjectModified
    ElseIf ObjectName = "FieldName" Then
        aSaveOrder = Split(Editor.FieldName(ObjectIndex).ToolTipText, ".")
        If Editor.FieldName(ObjectIndex).Top > (Editor.TableName(aSaveOrder(1)).Top + Editor.TableName(aSaveOrder(1)).Height) + 0.05 Then
            'move FieldName and FieldType down in Table
            Editor.FieldName(ObjectIndex).Top = (Editor.FieldName(ObjectIndex).Top - Editor.FieldName(ObjectIndex).Height) - 0.05
            Editor.FieldType(ObjectIndex).Top = (Editor.FieldType(ObjectIndex).Top - Editor.FieldType(ObjectIndex).Height) - 0.05
        End If
        MoveRelations
        ProjectModified
    End If

End Function

'move the table down
Function MoveObjectDown()

    If ObjectName = "Table" Then
        
        'down move restrictions
        If Editor.TableSelector(ObjectIndex).Top < ((Editor.EditField.Height - Editor.Table(ObjectIndex).Height) - 0.15) Then
            
            'move table
            Editor.Table(ObjectIndex).Top = Editor.Table(ObjectIndex).Top + 0.05
            Editor.TableShadow(ObjectIndex).Top = Editor.TableShadow(ObjectIndex).Top + 0.05
            Editor.TableSelector(ObjectIndex).Top = Editor.TableSelector(ObjectIndex).Top + 0.05
            Editor.Shape.Top = Editor.Shape.Top + 0.05
            Editor.TableName(ObjectIndex).Top = Editor.TableName(ObjectIndex).Top + 0.05
            
            ' move fields down
            For i = 0 To Editor.FieldName.Count - 1
                If Editor.FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                    ' move field
                    Editor.FieldName(i).Top = Editor.FieldName(i).Top + 0.05
                    Editor.FieldType(i).Top = Editor.FieldType(i).Top + 0.05
                End If
            Next i
        End If
        MoveMovers
        MoveRelations
        ProjectModified
    ElseIf ObjectName = "FieldName" Then
    
        'move FieldName and FieldType down in Table
        Editor.FieldName(ObjectIndex).Top = (Editor.FieldName(ObjectIndex).Top + Editor.FieldName(ObjectIndex).Height) + 0.05
        Editor.FieldType(ObjectIndex).Top = (Editor.FieldType(ObjectIndex).Top + Editor.FieldType(ObjectIndex).Height) + 0.05
        MoveRelations
        ProjectModified
    End If

End Function

'open a external project file
Function OpenFile()

    On Error GoTo Err:
    With Editor.CommonDialog1
        .CancelError = True
        .Filter = "Simple Database Designer (*.sdd)|*.sdd|"
        .ShowOpen
        FileName = .FileName
        Editor.Enabled = False
        Editor.MousePointer = vbHourglass
        Editor.StatusBar1.Panels(1).Text = "Loading file..."
        Open .FileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, FileContent
                
                'project properties
                If Left(FileContent, 7) = "PROJECT" Then
                    ArrayLineContent = Split(FileContent, "|")
                    'background color
                    Editor.EditField.Width = ArrayLineContent(1)
                    Editor.EditField.Height = ArrayLineContent(2)
                    Editor.EditField.BackColor = ArrayLineContent(3)
                    If ArrayLineContent(4) <> "" Then
                        Editor.Relation_1(0).BorderColor = ArrayLineContent(4)
                        Editor.Relation_1(0).BorderStyle = ArrayLineContent(5)
                        Editor.Relation_Caption(0).BackColor = ArrayLineContent(6)
                        Editor.Relation_Caption(0).ForeColor = ArrayLineContent(7)
                        Editor.Relation_Caption(0).BorderStyle = ArrayLineContent(8)
                    End If
                
                'table
                ElseIf Left(FileContent, 5) = "TABLE" Then
                    ArrayLineContent = Split(FileContent, "|")
                    Load Editor.Table(Editor.Table.Count)
                    Load Editor.TableShadow(Editor.TableShadow.Count)
                    Load Editor.TableName(Editor.TableName.Count)
                    Load Editor.TableSelector(Editor.TableSelector.Count)
                    'shadow
                    With Editor.TableShadow(Editor.TableShadow.ubound)
                        .Left = ArrayLineContent(2) + 0.03
                        .Top = ArrayLineContent(3) + 0.03
                        .Width = ArrayLineContent(4)
                        .Height = ArrayLineContent(5)
                        .FillColor = ArrayLineContent(6)
                        .ZOrder (0)
                        .Visible = True
                    End With
                    'table
                    With Editor.Table(Editor.Table.ubound)
                        .Left = ArrayLineContent(2)
                        .Top = ArrayLineContent(3)
                        .Width = ArrayLineContent(4)
                        .Height = ArrayLineContent(5)
                        .FillColor = ArrayLineContent(6)
                        .ZOrder (0)
                        .Visible = True
                    End With
                    'selector
                    With Editor.TableSelector(Editor.TableSelector.ubound)
                        .Left = ArrayLineContent(2)
                        .Top = ArrayLineContent(3)
                        .Width = ArrayLineContent(4)
                        .Height = ArrayLineContent(5)
                        .ZOrder (0)
                        .Visible = True
                    End With
                    'name
                    With Editor.TableName(Editor.TableName.ubound)
                        .Left = Editor.Table(Editor.Table.ubound).Left
                        .Top = Editor.Table(Editor.Table.ubound).Top
                        .Caption = ArrayLineContent(7)
                        .BackColor = ArrayLineContent(8)
                        .Width = ArrayLineContent(4)
                        .ZOrder (0)
                        .Visible = True
                    End With
                    'list
                    Editor.List1.AddItem "TABLE." & Editor.TableName.ubound
                'field
                ElseIf Left(FileContent, 5) = "FIELD" Then
                    ArrayLineContent = Split(FileContent, "|")
                    If ArrayLineContent(1) = 1 Then
                        'create name
                        Load Editor.FieldName(Editor.FieldName.Count)
                        With Editor.FieldName(Editor.FieldName.ubound)
                            .Left = Editor.Table(Editor.Table.ubound).Left + 0.1
                            .Top = ArrayLineContent(2)
                            .Width = (Editor.Table(Editor.Table.ubound).Width - Editor.FieldType(Editor.FieldType.ubound).Width) - 0.2
                            .Caption = ArrayLineContent(3)
                            .ZOrder (0)
                            .ToolTipText = "TABLE." & Editor.Table.ubound
                            .Visible = True
                        End With
                    ElseIf ArrayLineContent(1) = 2 Then
                        'create type
                        Load Editor.FieldType(Editor.FieldType.Count)
                        With Editor.FieldType(Editor.FieldType.ubound)
                            .Left = ((Editor.Table(Editor.Table.ubound).Left + Editor.Table(Editor.Table.ubound).Width) - .Width) - 0.05
                            .Top = ArrayLineContent(2)
                            .Caption = ArrayLineContent(3)
                            .ZOrder (0)
                            .ToolTipText = "TABLE." & Editor.Table.ubound
                            .Visible = True
                        End With
                    End If
                'relation
                ElseIf Left(FileContent, 8) = "RELATION" Then
                    ArrayLineContent = Split(FileContent, "|")
                    
                    Load Editor.Relation_1(Editor.Relation_1.Count)
                    Load Editor.Relation_Caption(Editor.Relation_Caption.Count)
                    
                    With Editor.Relation_1(Editor.Relation_1.ubound)
                        .X1 = ArrayLineContent(1)
                        .X2 = ArrayLineContent(2)
                        .Y1 = ArrayLineContent(3)
                        .Y2 = ArrayLineContent(4)
                        .Tag = ArrayLineContent(5)
                        .ZOrder (0)
                        .Visible = True
                    End With
                    With Editor.Relation_Caption(Editor.Relation_Caption.ubound)
                        .Caption = ArrayLineContent(5)
                        .Visible = True
                    End With

                    Editor.List2.AddItem ArrayLineContent(5)
                    Editor.Relation_Properties.AddItem ArrayLineContent(6)

                End If
            Loop
        Close #1
        
        Editor.mnuFileSave.Enabled = True
        Editor.Toolbar1.Buttons(4).Enabled = True

        DeSelect
        ResizeElements

        Editor.StatusBar1.Panels(3).Text = "Tables: " & Editor.TableName.ubound
        Editor.StatusBar1.Panels(4).Text = "Fields: " & Editor.FieldName.ubound

        Editor.Enabled = True
        Editor.MousePointer = vbDefault
        Editor.StatusBar1.Panels(1).Text = ""

        MoveRelations

    End With
Err:
    Exit Function

End Function

'deselect fields inside the table
Function DeselectFields()

    'deselect all other fieldnames
    For i = Editor.FieldName.LBound To Editor.FieldName.ubound
        Editor.FieldName(i).BorderStyle = 0
    Next i
    
    'deselect all other fieldtypes
    For i = 0 To Editor.FieldType.Count - 1
        Editor.FieldType(i).BorderStyle = 0
    Next i

End Function

'deselect all selected objects (tables and relations)
Function DeSelect()

    With Editor.Shape
        .Visible = False
    End With
    
    Editor.StatusBar1.Panels(1).Text = ""
    Editor.SetApply.Enabled = False
    
    Editor.mnuEditDeselect.Enabled = False
    Editor.mnuProjectTableFields.Enabled = False
    Editor.mnuEditFront.Enabled = False
    Editor.mnuEditBack.Enabled = False
    
    DeselectFields

    Editor.MoveTable.Visible = False
    Editor.MoveTable2.Visible = False
    
    'reset selected relation
    For i = 1 To Editor.Relation_1.Count - 1
        Editor.Relation_1(i).BorderWidth = 1
        Editor.Relation_Caption(i).FontBold = False
    Next i
    
    Editor.List1.Text = ""
    Editor.List2.Text = ""
    
    Editor.mnuEditDelete.Enabled = False
    
    ObjectName = ""

End Function

'create a new project
Function NewFile()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Create new project") = vbYes Then
        
        Editor.StatusBar1.Panels(1).Text = "Creating new project..."
        Editor.MousePointer = vbHourglass
        Editor.Enabled = False
        
        DeSelect
        
        'unload all table elements
        For i = 1 To Editor.Table.Count - 1
            Unload Editor.Table(i)
        Next i
        For i = 1 To Editor.TableShadow.Count - 1
            Unload Editor.TableShadow(i)
        Next i
        For i = 1 To Editor.TableName.Count - 1
            Unload Editor.TableName(i)
        Next i
        For i = 1 To Editor.TableSelector.Count - 1
            Unload Editor.TableSelector(i)
        Next i
        
        'unload all field elements
        For i = 1 To Editor.FieldName.Count - 1
            Unload Editor.FieldName(i)
        Next i
        For i = 1 To Editor.FieldType.Count - 1
            Unload Editor.FieldType(i)
        Next i
        
        'unload all relation elements
        For i = 1 To Editor.Relation_1.Count - 1
            Unload Editor.Relation_1(i)
        Next i
        For i = 1 To Editor.Relation_Caption.Count - 1
            Unload Editor.Relation_Caption(i)
        Next i
        Editor.List2.Clear
        Editor.Relation_Properties.Clear
        
        Editor.EditField.Width = 150 / 10
        Editor.EditField.Height = 100 / 10
        
        Editor.List1.Clear
        Editor.EditField.BackColor = 16777215
        Editor.mnuFileSave.Enabled = False
        Editor.Toolbar1.Buttons(4).Enabled = False
        FileName = ""
        
        Editor.Enabled = True
        Editor.MousePointer = vbDefault
        Editor.StatusBar1.Panels(1).Text = ""
        Editor.StatusBar1.Panels(3).Text = "Tables: 0"
        Editor.StatusBar1.Panels(4).Text = "Fields: 0"
        
        ResizeElements
        Editor.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
        
    End If

End Function

'save the current project
Function SaveFile(FileName As String)

    Editor.StatusBar1.Panels(1).Text = "Saving file..."
    Editor.MousePointer = vbHourglass
    Editor.Enabled = False
    Open FileName For Output As #1
        
        'save project properties
        Print #1, "PROJECT|" & Editor.EditField.Width & "|" & Editor.EditField.Height & "|" & Editor.EditField.BackColor & "|" & Editor.Relation_1(0).BorderColor & "|" & Editor.Relation_1(0).BorderStyle & "|" & Editor.Relation_Caption(0).BackColor & "|" & Editor.Relation_Caption(0).ForeColor & "|" & Editor.Relation_Caption(0).BorderStyle
        
        'save table data
        For i = 0 To Editor.List1.ListCount - 1
            aSaveOrder = Split(Editor.List1.List(i), ".")
            If Editor.Table(aSaveOrder(1)).Visible = True Then
                Print #1, "TABLE|" & aSaveOrder(1) & "|" & Editor.Table(aSaveOrder(1)).Left & "|" & Editor.Table(aSaveOrder(1)).Top & "|" & Editor.Table(aSaveOrder(1)).Width & "|" & Editor.Table(aSaveOrder(1)).Height & "|" & Editor.Table(aSaveOrder(1)).FillColor & "|" & Editor.TableName(aSaveOrder(1)).Caption & "|" & Editor.TableName(aSaveOrder(1)).BackColor
            End If
            
            'save fields from current table
            For j = 0 To Editor.FieldName.Count - 1
                If Editor.FieldName(j).Visible = True Then
                    If Editor.FieldName(j).ToolTipText = "TABLE." & aSaveOrder(1) Then
                        Print #1, "FIELD|1|" & Editor.FieldName(j).Top & "|" & Editor.FieldName(j).Caption
                        Print #1, "FIELD|2|" & Editor.FieldType(j).Top & "|" & Editor.FieldType(j).Caption
                    End If
                End If
            Next j
        Next i
        
        'save relations
        For i = 1 To Editor.Relation_1.Count - 1
            If Editor.Relation_1(i).Visible = True Then
                Print #1, "RELATION|" & Editor.Relation_1(i).X1 & "|" & Editor.Relation_1(i).X2 & "|" & Editor.Relation_1(i).Y1 & "|" & Editor.Relation_1(i).Y2 & "|" & Editor.Relation_1(i).Tag & "|" & Editor.Relation_Properties.List(i - 1)
            End If
        Next i
        
    Close #1
    Editor.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Editor.MousePointer = vbDefault
    Editor.Enabled = True
    Editor.StatusBar1.Panels(1).Text = "Done"

End Function

'insert new field to the selected table
Function InsertField()

    Load Editor.FieldName(Editor.FieldName.Count)
    Load Editor.FieldType(Editor.FieldType.Count)
    
    'fieldname
    With Editor.FieldName(Editor.FieldName.ubound)
        .Left = Editor.Table(ObjectIndex).Left + 0.05
        .Top = Editor.TableName(ObjectIndex).Top + Editor.TableName(ObjectIndex).Height + 0.05
        .Width = (Editor.Table(ObjectIndex).Width - Editor.FieldType(Editor.FieldType.ubound).Width) - 0.15
        .ZOrder (0)
        .ToolTipText = "TABLE." & ObjectIndex
        .Visible = True
    End With
    
    'fieldtype
    With Editor.FieldType(Editor.FieldType.ubound)
        .Left = ((Editor.Table(ObjectIndex).Left + Editor.Table(ObjectIndex).Width) - .Width) - 0.05
        .Top = Editor.TableName(ObjectIndex).Top + Editor.TableName(ObjectIndex).Height + 0.05
        .ZOrder (0)
        .ToolTipText = "TABLE." & ObjectIndex
        .Visible = True
    End With
    
    Editor.StatusBar1.Panels(4).Text = "Fields: " & Editor.FieldName.ubound

End Function

' move relations to the right position
Function MoveRelations()

    'move relations
    For i = 1 To Editor.Relation_1.Count - 1
        aTable = Split(Editor.Relation_Properties.List(i - 1), "-")
        aTable1 = Split(aTable(0), ".")
        aTable2 = Split(aTable(1), ".")
        
        With Editor.Relation_1(i)
            .X1 = Editor.FieldType(aTable1(0)).Left + Editor.FieldType(aTable1(0)).Width + 0.05 'left position
            .Y1 = Editor.FieldType(aTable1(0)).Top + (Editor.FieldType(aTable1(0)).Height / 2) 'top position
            .X2 = Editor.FieldType(aTable2(0)).Left + Editor.FieldType(aTable1(0)).Width + 0.05 'right width
            .Y2 = Editor.FieldType(aTable2(0)).Top + (Editor.FieldType(aTable2(0)).Height / 2) 'left bottom
            .ZOrder (0)
        End With
        
        With Editor.Relation_Caption(i)
            .Left = Editor.Relation_1(i).X1 + ((Editor.Relation_1(i).X2 - Editor.Relation_1(i).X1) / 2) - (Len(.Caption) / 12) / 2
            .Top = Editor.Relation_1(i).Y1 + ((Editor.Relation_1(i).Y2 - Editor.Relation_1(i).Y1) / 2)
            .ZOrder (0)
        End With
        
    Next i
End Function

' move the movers to the right position around the selecte table
Function MoveMovers()
    If Editor.mnuEditMovers.Checked = True Then

        'movetable 1
        Editor.MoveTable.Left = (Editor.Table(ObjectIndex).Left + Editor.Table(ObjectIndex).Width) + 0.05
        Editor.MoveTable.Top = (Editor.Table(ObjectIndex).Top + Editor.Table(ObjectIndex).Height) + 0.05
        Editor.MoveTable.Visible = True
        Editor.MoveTable.ZOrder (0)

        'movetable 2
        Editor.MoveTable2.Left = Editor.Table(ObjectIndex).Left - Editor.MoveTable.Width
        Editor.MoveTable2.Top = Editor.Table(ObjectIndex).Top - Editor.MoveTable.Width
        Editor.MoveTable2.Visible = True
        Editor.MoveTable2.ZOrder (0)
    End If
End Function

'resize form elements
Function ResizeElements()

    On Error Resume Next

    Editor.WorkArea.Width = Editor.ScaleWidth - Editor.WorkBar.ScaleWidth
    Editor.Shadow.Left = Editor.EditField.Left + 0.05
    Editor.Shadow.Top = Editor.EditField.Top + 0.05
    Editor.Shadow.Width = Editor.EditField.Width + 0.02
    Editor.Shadow.Height = Editor.EditField.Height + 0.02
    
    Editor.HScroll1.Left = 0
    Editor.HScroll1.Width = Editor.WorkArea.ScaleWidth - Editor.VScroll1.Width
    Editor.HScroll1.Top = Editor.WorkArea.ScaleHeight - Editor.HScroll1.Height
    
    Editor.VScroll1.Left = Editor.WorkArea.ScaleWidth - Editor.VScroll1.Width
    Editor.VScroll1.Top = 0
    Editor.VScroll1.Height = Editor.WorkArea.ScaleHeight - Editor.HScroll1.Height
    
    Editor.EmptySpace.Width = Editor.VScroll1.Width
    Editor.EmptySpace.Height = Editor.HScroll1.Height
    Editor.EmptySpace.Left = Editor.HScroll1.Width
    Editor.EmptySpace.Top = Editor.VScroll1.Height
    
    'enabled of disable the scrollbars
    'horizontal scrollbar
    If Editor.WorkArea.ScaleWidth > (Editor.EditField.ScaleWidth + Editor.HScroll1.Height) + 0.5 Then
        Editor.HScroll1.Enabled = False
    Else
       Editor.HScroll1.Enabled = True
       Editor.HScroll1.LargeChange = 1
       Editor.HScroll1.SmallChange = 1
       Editor.HScroll1.Max = Round(Editor.EditField.Width)
    End If

    'vertical scrollbar
    If Editor.WorkArea.ScaleHeight > (Editor.EditField.ScaleHeight + Editor.VScroll1.Width) + 0.5 Then
        Editor.VScroll1.Enabled = False
    Else
       Editor.VScroll1.Enabled = True
       Editor.VScroll1.LargeChange = 1
       Editor.VScroll1.SmallChange = 1
       Editor.VScroll1.Max = Round(Editor.EditField.Height)
    End If
    
    Editor.StatusBar1.Panels(2).Text = Round(Editor.EditField.Width, 2) * 10 & " x " & Round(Editor.EditField.Height, 2) * 10 & " cm"

End Function

'check if project was changes
Function ProjectModified()

    Editor.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & " *"

End Function
