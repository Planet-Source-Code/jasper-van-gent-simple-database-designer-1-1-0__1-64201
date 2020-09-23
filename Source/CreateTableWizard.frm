VERSION 5.00
Begin VB.Form CreateTableWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create new table"
   ClientHeight    =   3780
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3612
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3612
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   "Near table"
      Height          =   2772
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3132
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   1080
         TabIndex        =   8
         Text            =   "0,5"
         Top             =   2280
         Width           =   732
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1920
         Width           =   1812
      End
      Begin VB.ListBox List1 
         Height          =   1392
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2652
      End
      Begin VB.Label Label2 
         Caption         =   "Spacing:"
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Position:"
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   732
      End
   End
   Begin VB.Label SelectedTable 
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   492
   End
End
Attribute VB_Name = "CreateTableWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PositionLeft As String
Dim PositionTop As String

Dim aTable As Variant

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    'check on other size
    If List1.Text <> "" And Combo1.Text <> "" Then
        aTable = Split(List1.Text, ".")
        If Combo1.Text = "Right" Then
            PositionLeft = (Editor.Table(aTable(0)).Left + Editor.Table(aTable(0)).Width) + Text1.Text
            PositionTop = Editor.Table(aTable(0)).Top
        ElseIf Combo1.Text = "Left" Then
            PositionLeft = (Editor.Table(aTable(0)).Left - Editor.Table(aTable(0)).Width) - Text1.Text
            PositionTop = Editor.Table(aTable(0)).Top
        ElseIf Combo1.Text = "Below" Then
            PositionLeft = Editor.Table(aTable(0)).Left
            PositionTop = (Editor.Table(aTable(0)).Top + Editor.Table(aTable(0)).Height) + Text1.Text
        ElseIf Combo1.Text = "Above" Then
            PositionTop = (Editor.Table(aTable(0)).Top - Editor.Table(aTable(0)).Height) - Text1.Text
            PositionLeft = Editor.Table(aTable(0)).Left
        End If
    End If

    With Editor.Shape
        .Visible = False
    End With
    
    Load Editor.Table(Editor.Table.Count)
    Load Editor.TableShadow(Editor.TableShadow.Count)
    Load Editor.TableSelector(Editor.TableSelector.Count)
    Load Editor.TableName(Editor.TableName.Count)
    
    'shadow
    With Editor.TableShadow(Editor.TableShadow.UBound)
        .Left = PositionLeft + 0.03
        .Top = PositionTop + 0.03
        .FillColor = Editor.SetColor.BackColor
        .Width = Editor.SetWidth.Text
        .Height = Editor.Setheight.Text
        .ZOrder (0)
        .Visible = True
    End With
    'shape
    With Editor.Table(Editor.Table.UBound)
        .Left = PositionLeft
        .Top = PositionTop
        .FillColor = Editor.SetColor.BackColor
        .Width = Editor.SetWidth.Text
        .Height = Editor.Setheight.Text
        .ZOrder (0)
        .Visible = True
    End With
    'selector
    With Editor.TableSelector(Editor.TableSelector.UBound)
        .Left = Editor.Table(Editor.Table.UBound).Left
        .Top = Editor.Table(Editor.Table.UBound).Top
        .Width = Editor.Table(Editor.Table.UBound).Width
        .Height = Editor.Table(Editor.Table.UBound).Height
        .ZOrder (0)
        .Visible = True
    End With
    'tablename
    With Editor.TableName(Editor.TableName.UBound)
        .Left = Editor.Table(Editor.Table.UBound).Left
        .Top = Editor.Table(Editor.Table.UBound).Top
        .Width = Editor.Table(Editor.Table.UBound).Width
        .Caption = "Untitled_" & Editor.TableName.UBound
        .ZOrder (0)
        .Visible = True
    End With
    
    'resize EditField if required
    'height
    If Editor.EditField.ScaleHeight <= (Editor.Table(Editor.Table.UBound).Top + Editor.Table(Editor.Table.UBound).Height) Then
        Editor.EditField.Height = (Editor.Table(Editor.Table.UBound).Top + Editor.Table(Editor.Table.UBound).Height) + Text1.Text
    End If
    'width
    If Editor.EditField.ScaleWidth <= (Editor.Table(Editor.Table.UBound).Left + Editor.Table(Editor.Table.UBound).Width) Then
        Editor.EditField.Width = (Editor.Table(Editor.Table.UBound).Left + Editor.Table(Editor.Table.UBound).Width) + Text1.Text
    End If
    
    Editor.List1.AddItem "TABLE." & Editor.Table.UBound
    Editor.StatusBar1.Panels(3).Text = "Tables: " & Editor.Table.UBound
    
    ProjectModified
    ResizeElements
    
    'select the created table
    With Editor.Shape
        .Left = Editor.Table(Editor.Table.UBound).Left - 0.05
        .Top = Editor.Table(Editor.Table.UBound).Top - 0.05
        .Width = Editor.Table(Editor.Table.UBound).Width + 0.15
        .Height = Editor.Table(Editor.Table.UBound).Height + 0.15
        .Visible = True
    End With
    
    Editor.SetWidth = Round(Editor.Table(Editor.Table.UBound).Width, 2)
    Editor.Setheight = Round(Editor.Table(Editor.Table.UBound).Height, 2)
    
    ObjectName = Editor.Table(Editor.Table.UBound).Name
    ObjectIndex = Editor.Table(Editor.Table.UBound).Index
    Editor.StatusBar1.Panels(1).Text = ObjectName & " " & ObjectIndex & " (" & Editor.TableName(ObjectIndex).Caption & ")"
    
    Editor.Shape.ZOrder (0)
    
    Editor.OrgColor.BackColor = Editor.Table(Editor.Table.UBound).FillColor
    Editor.List1.Text = "TABLE." & Editor.Table.UBound

    Editor.SetApply.Enabled = True
    Editor.mnuEditDeselect.Enabled = True
    Editor.mnuProjectTableFields.Enabled = True
    Editor.mnuEditFront.Enabled = True
    Editor.mnuEditBack.Enabled = True
    
    Editor.mnuEditDelete.Caption = "Delete table"
    Editor.mnuEditDelete.Enabled = True

    MoveMovers
    
    DeselectFields
    
    Unload Me

End Sub

Private Sub Form_Load()

    PositionLeft = 0.5
    PositionTop = 0.5

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Editor.Enabled = True
    Editor.Show

End Sub
