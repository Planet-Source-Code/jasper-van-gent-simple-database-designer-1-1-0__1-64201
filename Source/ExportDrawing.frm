VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E5CEE37F-8CF8-489E-BFA0-8201CBD6AEE8}#1.0#0"; "PicFormat32.ocx"
Begin VB.Form ExportDrawing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export to Drawing"
   ClientHeight    =   3816
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3816
   ScaleWidth      =   3480
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   1920
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin PicFormat32a.PicFormat32 PicFormat321 
      Height          =   252
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   445
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drawing layout"
      Height          =   2772
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3012
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   288
         Left            =   600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2160
         Width           =   1692
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Add legend"
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1812
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Add current date/time"
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1812
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Display table shadow"
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1932
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Display labels"
         Height          =   252
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   1332
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Display relations"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1572
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Create"
      Default         =   -1  'True
      Height          =   372
      Left            =   960
      TabIndex        =   1
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   3240
      Width           =   1092
   End
End
Attribute VB_Name = "ExportDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StartLeft As String
Dim StartTop As String
Dim NewStartLeft As String
Dim NewStarttop As String

Dim aFieldTypeExtra As Variant
Dim aRelation As Variant
Dim aRelationExtra As Variant

Private Sub Check1_Click()

    If Check1.Value = 1 Then
        Check2.Enabled = True
    Else
        Check2.Enabled = False
    End If

End Sub

Private Sub Check5_Click()

    If Check5.Value = 1 Then
        Combo1.Enabled = True
    Else
        Combo1.Enabled = False
    End If

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Command3_Click()

    On Error GoTo Err:

    With CommonDialog1
        .CancelError = True
        .Filter = "Bitmap File (*.bmp)|*.bmp|GIF File (*.gif)|*.gif|JPG File (*.jpg)|*.jpg|"
        .Flags = cdlOFNOverwritePrompt
        .CancelError = True
        .ShowSave
        
        Me.MousePointer = vbHourglass
        Me.Enabled = False
        
        'remove old picture
        Editor.BitmapFile.Picture = LoadPicture("")
        
        'set drawing sizes
        Editor.BitmapFile.Width = Editor.EditField.Width
        Editor.BitmapFile.Height = Editor.EditField.Height
        Editor.BitmapFile.BackColor = Editor.EditField.BackColor
        
        'draw all tables (use list order)
        For i = 0 To Editor.List1.ListCount - 1
            aSaveOrder = Split(Editor.List1.List(i), ".")
            
            If Check3.Value = 1 Then
                'draw shadow
                Editor.BitmapFile.Line (Editor.Table(aSaveOrder(1)).Left + 0.06, Editor.Table(aSaveOrder(1)).Top + 0.06)-Step(Editor.TableShadow(aSaveOrder(1)).Width, Editor.TableShadow(aSaveOrder(1)).Height), Editor.TableShadow(aSaveOrder(1)).BackColor, BF
            End If
            
            'draw table shadow
            Editor.BitmapFile.Line (Editor.Table(aSaveOrder(1)).Left - 0.02, Editor.Table(aSaveOrder(1)).Top - 0.02)-Step(Editor.Table(aSaveOrder(1)).Width + 0.04, Editor.Table(aSaveOrder(1)).Height + 0.04), vbBlack, BF
            
            'draw table
            Editor.BitmapFile.Line (Editor.Table(aSaveOrder(1)).Left, Editor.Table(aSaveOrder(1)).Top)-Step(Editor.Table(aSaveOrder(1)).Width, Editor.Table(aSaveOrder(1)).Height), Editor.Table(aSaveOrder(1)).FillColor, BF
            
            'draw tablename
            Editor.BitmapFile.Line (Editor.TableName(aSaveOrder(1)).Left, Editor.TableName(aSaveOrder(1)).Top)-Step(Editor.TableName(aSaveOrder(1)).Width, Editor.TableName(aSaveOrder(1)).Height), Editor.TableName(aSaveOrder(1)).BackColor, BF
            
            'draw text on tablename
            Editor.BitmapFile.CurrentX = Editor.TableName(aSaveOrder(1)).Left + (Editor.TableName(aSaveOrder(1)).Width / 2) - ((Len(Editor.TableName(aSaveOrder(1)).Caption) / 6) / 2)
            Editor.BitmapFile.CurrentY = Editor.TableName(aSaveOrder(1)).Top + 0.05
            Editor.BitmapFile.FontName = Editor.TableName(aSaveOrder(1)).FontName
            Editor.BitmapFile.ForeColor = Editor.TableName(aSaveOrder(1)).ForeColor
            Editor.BitmapFile.FontSize = Editor.TableName(aSaveOrder(1)).FontSize
            Editor.BitmapFile.Print Editor.TableName(aSaveOrder(1)).Caption
        
            'draw fields from table
            For j = 0 To Editor.FieldName.Count - 1
                If Editor.FieldName(j).ToolTipText = "TABLE." & aSaveOrder(1) Then
                    
                    'draw fields
                    Editor.BitmapFile.Line (Editor.FieldName(j).Left, Editor.FieldName(j).Top)-Step(Editor.FieldName(j).Width, Editor.FieldName(j).Height), Editor.FieldName(j).BackColor, BF
                    Editor.BitmapFile.Line (Editor.FieldType(j).Left, Editor.FieldType(j).Top)-Step(Editor.FieldType(j).Width, Editor.FieldType(j).Height), Editor.FieldType(j).BackColor, BF
                    
                    'draw fieldname text inside fields
                    Editor.BitmapFile.CurrentX = Editor.FieldName(j).Left + 0.02
                    Editor.BitmapFile.CurrentY = Editor.FieldName(j).Top + 0.02
                    Editor.BitmapFile.FontName = Editor.FieldName(j).FontName
                    Editor.BitmapFile.ForeColor = Editor.FieldName(j).ForeColor
                    Editor.BitmapFile.FontSize = Editor.FieldName(j).FontSize
                    Editor.BitmapFile.Print Editor.FieldName(j).Caption
                    
                    aFieldTypeExtra = Split(Editor.FieldType(j).Caption, " ")
                    
                    'draw fieldtype text inside fields
                    Editor.BitmapFile.CurrentX = Editor.FieldType(j).Left + 0.02
                    Editor.BitmapFile.CurrentY = Editor.FieldType(j).Top + 0.02
                    Editor.BitmapFile.FontName = Editor.FieldType(j).FontName
                    Editor.BitmapFile.ForeColor = Editor.FieldType(j).ForeColor
                    Editor.BitmapFile.FontSize = Editor.FieldType(j).FontSize
                    Editor.BitmapFile.Print aFieldTypeExtra(0)
                    
                    If aFieldTypeExtra(1) = "P" Then
                        Editor.BitmapFile.PaintPicture Editor.Bitmaps.ListImages(1).Picture, (Editor.FieldType(j).Left + Editor.FieldType(j).Width) - 0.6, Editor.FieldType(j).Top, 0.3, 0.3
                    ElseIf aFieldTypeExtra(1) = "A" Then
                        Editor.BitmapFile.PaintPicture Editor.Bitmaps.ListImages(2).Picture, (Editor.FieldType(j).Left + Editor.FieldType(j).Width) - 0.3, Editor.FieldType(j).Top, 0.3, 0.3
                    ElseIf aFieldTypeExtra(1) = "PA" Then
                        Editor.BitmapFile.PaintPicture Editor.Bitmaps.ListImages(1).Picture, (Editor.FieldType(j).Left + Editor.FieldType(j).Width) - 0.6, Editor.FieldType(j).Top, 0.3, 0.3
                        Editor.BitmapFile.PaintPicture Editor.Bitmaps.ListImages(2).Picture, (Editor.FieldType(j).Left + Editor.FieldType(j).Width) - 0.3, Editor.FieldType(j).Top, 0.3, 0.3
                    End If

                End If
            Next j
            
            'draw relations
            If Check1.Value = 1 Then
                For j = 1 To Editor.Relation_1.Count - 1
                    Editor.BitmapFile.Line (Editor.Relation_1(j).X1, Editor.Relation_1(j).Y1)-(Editor.Relation_1(j).X2, Editor.Relation_1(j).Y2), Editor.Relation_1(0).BorderColor
                    If Check2.Value = 1 Then
                        'draw text
                        Editor.BitmapFile.CurrentX = Editor.Relation_1(j).X1 + ((Editor.Relation_1(j).X2 - Editor.Relation_1(j).X1) / 2) - (Len(Editor.Relation_1(j).Tag) / 12) / 2
                        Editor.BitmapFile.CurrentY = Editor.Relation_1(j).Y1 + ((Editor.Relation_1(j).Y2 - Editor.Relation_1(j).Y1) / 2)
                        Editor.BitmapFile.Print Editor.Relation_1(j).Tag
                    End If
                Next j
            End If
        
        Next i
        
        'add current date and time
        If Check4.Value = 1 Then
            'draw current date and time
            Editor.BitmapFile.CurrentX = Editor.EditField.ScaleWidth - 2.3
            Editor.BitmapFile.CurrentY = Editor.EditField.ScaleHeight - 0.4
            Editor.BitmapFile.FontName = "Arial"
            Editor.BitmapFile.ForeColor = vbBlack
            Editor.BitmapFile.FontSize = 6
            Editor.BitmapFile.FontBold = True
            Editor.BitmapFile.Print Format(Now(), "yyyy-mm-dd hh:nn:ss")
        End If
        
        ' add legend
        If Check5.Value = 1 Then
            
            ' get start position
            If Combo1.Text = "Left top" Then
                StartLeft = 0.5
                StartTop = 0.5
            ElseIf Combo1.Text = "Right top" Then
                StartLeft = Editor.EditField.ScaleWidth - 10.5
                StartTop = 0.5
            ElseIf Combo1.Text = "Left bottom" Then
                StartLeft = 0.5
                StartTop = Editor.EditField.ScaleHeight - (Editor.List2.ListCount / 3.5)
            ElseIf Combo1.Text = "Right bottom" Then
                StartLeft = Editor.EditField.ScaleWidth - 10.5
                StartTop = Editor.EditField.ScaleHeight - (Editor.List2.ListCount / 3.5)
            End If

            'draw box behind legend
            Editor.BitmapFile.Line (StartLeft + 0.02, StartTop + 0.02)-Step(10, StartTop + (Editor.List2.ListCount / 3.5) + 0.05), vbBlack, BF
            Editor.BitmapFile.Line (StartLeft, StartTop)-Step(10, StartTop + (Editor.List2.ListCount / 3.5) + 0.05), vbYellow, BF
            Editor.BitmapFile.Line (StartLeft, StartTop)-Step(10, StartTop + (Editor.List2.ListCount / 3.5) + 0.05), vbBlack, B
            
            Editor.BitmapFile.CurrentX = StartLeft + 0.1
            Editor.BitmapFile.CurrentY = StartTop + 0.1
            Editor.BitmapFile.FontName = "Arial"
            Editor.BitmapFile.ForeColor = vbBlack
            Editor.BitmapFile.FontSize = 7
            Editor.BitmapFile.FontBold = True
            Editor.BitmapFile.Print "Legend"
            Editor.BitmapFile.FontBold = False
            NewStarttop = StartTop + 0.45

            For i = 0 To Editor.List2.ListCount - 1
                aRelation = Split(Editor.Relation_Properties.List(i), ".")
                aRelationExtra = Split(aRelation(2), "-")
                Editor.BitmapFile.CurrentY = NewStarttop + (i / 3.5)
                Editor.BitmapFile.CurrentX = StartLeft + 0.1
                Editor.BitmapFile.Print Editor.List2.List(i) & " = " & aRelation(1) & "." & aRelationExtra(0) & " with " & aRelation(3) & "." & aRelation(4)
            Next i
            
        End If
        
        Set Editor.BitmapFile.Picture = Editor.BitmapFile.Image
        SavePicture Editor.BitmapFile.Picture, App.Path & "\~tmp.bmp"

        'bmp
        If .FilterIndex = 1 Then
            SavePicture Editor.BitmapFile.Picture, .FileName
        
        'gif
        ElseIf .FilterIndex = 2 Then
            PicFormat321.SaveBmpToGif App.Path & "\~tmp.bmp", .FileName
        
        'jpg
        ElseIf .FilterIndex = 3 Then
            PicFormat321.SaveBmpToJpeg App.Path & "\~tmp.bmp", .FileName, 80
        
        End If
        
        'remove temp picture
        Kill App.Path & "\~tmp.bmp"
        
        Me.MousePointer = vbDefault
        Me.Enabled = True
        
        Unload Me

    End With
    
Err:
    Exit Sub

End Sub

Private Sub Form_Load()

    Combo1.AddItem "Left top"
    Combo1.AddItem "Right top"
    Combo1.Text = "Right top"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Editor.Enabled = True
    Editor.Show

End Sub
