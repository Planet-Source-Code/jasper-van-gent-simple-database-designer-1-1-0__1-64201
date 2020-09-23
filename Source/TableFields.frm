VERSION 5.00
Begin VB.Form TableFields 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table fields"
   ClientHeight    =   4284
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   7824
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4284
   ScaleWidth      =   7824
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   3252
      Left            =   3960
      TabIndex        =   13
      Top             =   240
      Width           =   3612
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   1200
         TabIndex        =   2
         Text            =   "45"
         Top             =   1200
         Width           =   612
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "VARCHAR"
         Top             =   840
         Width           =   1692
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   1692
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Insert as new"
         Height          =   372
         Left            =   1440
         TabIndex        =   6
         Top             =   2640
         Width           =   1572
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   372
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   1092
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto increment"
         Height          =   252
         Left            =   1200
         TabIndex        =   4
         Top             =   2040
         Width           =   1452
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Primary key"
         Height          =   252
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "Length:"
         Height          =   252
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label Label2 
         Caption         =   "DataType:"
         Height          =   252
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   852
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fields"
      Height          =   3252
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   3492
      Begin VB.CommandButton Command7 
         Caption         =   "Remove"
         Height          =   372
         Left            =   2040
         TabIndex        =   17
         Top             =   2640
         Width           =   1212
      End
      Begin VB.CommandButton Command6 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   10.2
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   720
         TabIndex        =   9
         Top             =   2640
         Width           =   372
      End
      Begin VB.CommandButton Command5 
         Caption         =   "6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   10.2
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   372
      End
      Begin VB.ListBox List1 
         Height          =   2160
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3012
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   372
      Left            =   5040
      TabIndex        =   10
      Top             =   3720
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   6360
      TabIndex        =   11
      Top             =   3720
      Width           =   1212
   End
End
Attribute VB_Name = "TableFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NewFieldName As String
Dim ExtraValue As String

Dim aFieldName As Variant
Dim aFieldNameLength As Variant

Function CreateField()

    If Check1.Value = 1 Then
        ExtraValue = "P"
    ElseIf Check2.Value = 1 Then
        ExtraValue = "A"
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
        ExtraValue = "PA"
    End If
    If Check1.Value = 0 And Check2.Value = 0 Then
        ExtraValue = ""
    End If

    NewFieldName = Replace(Text1.Text, " ", "_") & " -> " & Replace(UCase(Combo1.Text), " ", "") & "(" & Text2.Text & ") " & ExtraValue

End Function

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    'remove all old fields
    For i = Editor.FieldName.LBound To Editor.FieldName.ubound
        If Editor.FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
            ' remove field
            Editor.FieldName(i).Visible = False
            Editor.FieldType(i).Visible = False
            Editor.FieldName(i).ToolTipText = ""
            Editor.FieldType(i).ToolTipText = ""
        End If
    Next i

    'create new fields
    For i = 0 To List1.ListCount - 1

        Load Editor.FieldName(Editor.FieldName.ubound + 1)
        Load Editor.FieldType(Editor.FieldType.ubound + 1)

        aFieldName = Split(List1.List(i), " ")

        ' set fieldname
        With Editor.FieldName(Editor.FieldName.ubound)
            .Caption = aFieldName(0)
            .Left = Editor.Table(ObjectIndex).Left + 0.05
            If i = 0 Then
                .Top = (Editor.TableName(ObjectIndex).Top + Editor.TableName(ObjectIndex).Height) + 0.05
            Else
                .Top = (Editor.FieldName(Editor.FieldName.ubound - 1).Top + Editor.FieldName(Editor.FieldName.ubound - 1).Height) + 0.05
            End If
            .Width = (Editor.Table(ObjectIndex).Width - Editor.FieldType(Editor.FieldType.ubound).Width) - 0.15
            .ZOrder (0)
            .ToolTipText = "TABLE." & ObjectIndex
            .Visible = True
        End With
        
        ' set fieldtype
        With Editor.FieldType(Editor.FieldType.ubound)
            .Left = ((Editor.Table(ObjectIndex).Left + Editor.Table(ObjectIndex).Width) - .Width) - 0.05
            .Top = Editor.FieldName(Editor.FieldName.ubound).Top
            .Caption = aFieldName(2) & " " & aFieldName(3)
            .ZOrder (0)
            .ToolTipText = "TABLE." & ObjectIndex
            .Visible = True
        End With

    Next i
    
    Unload Me

End Sub

Private Sub Command3_Click()

    CreateField

    i = List1.ListIndex
    List1.RemoveItem i
    List1.AddItem NewFieldName, i
    
    List1.Text = NewFieldName
    
    Command2.Enabled = True

End Sub

Private Sub Command4_Click()

    CreateField

    List1.Text = NewFieldName
    If List1.Text = "" Then
        List1.AddItem NewFieldName
        List1.Text = NewFieldName
        Text1.SetFocus
        SendKeys "{END}+{HOME}"
        Command2.Enabled = True
    Else
        MsgBox "Field already exists!", vbCritical + vbOKOnly, "Insert field"
    End If

End Sub

Private Sub Command5_Click()

    If List1.ListIndex < (List1.ListCount - 1) Then
        i = List1.ListIndex
        NewFieldName = List1.Text
        List1.RemoveItem i
        List1.AddItem NewFieldName, (i + 1)
        List1.Text = NewFieldName
        Command2.Enabled = True
    End If

End Sub

Private Sub Command6_Click()

    If List1.ListIndex <> 0 Then
        i = List1.ListIndex
        NewFieldName = List1.Text
        List1.RemoveItem i
        List1.AddItem NewFieldName, (i - 1)
        List1.Text = NewFieldName
        Command2.Enabled = True
    End If

End Sub

Private Sub Command7_Click()

    List1.RemoveItem List1.ListIndex
    Command2.Enabled = True

End Sub

Private Sub Form_Load()


    Combo1.AddItem "TINYINT"
    Combo1.AddItem "SMALLINT"
    Combo1.AddItem "MEDIUMINT"
    Combo1.AddItem "INT"
    Combo1.AddItem "BIGINT"
    Combo1.AddItem "FLOAT"
    Combo1.AddItem "DOUBLE"
    Combo1.AddItem "PRECISION"
    Combo1.AddItem "DECIMAL"
    Combo1.AddItem "NUMBERIC"
    Combo1.AddItem "DATE"
    Combo1.AddItem "DATETIME"
    Combo1.AddItem "TIMESTAMP"
    Combo1.AddItem "TIME"
    Combo1.AddItem "YEAR"
    Combo1.AddItem "CHAR"
    Combo1.AddItem "VARCHAR"
    Combo1.AddItem "TINYBLOB"
    Combo1.AddItem "TINYTEXT"
    Combo1.AddItem "BLOB"
    Combo1.AddItem "TEXT"
    Combo1.AddItem "REAL"
    Combo1.AddItem "MEDIUMBLOB"
    Combo1.AddItem "MEDIUMTEXT"
    Combo1.AddItem "LONGBLOB"
    Combo1.AddItem "LONGTEXT"
    Combo1.AddItem "ENUM"
    Combo1.AddItem "SET"
    Combo1.AddItem "VARCHAR"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Editor.Enabled = True
    Editor.Show

End Sub

Private Sub List1_Click()

    If List1.Text <> "" Then
    
        'display values inside properties field
        aFieldName = Split(List1.Text, " ")
        Text1.Text = aFieldName(0)
        aFieldNameLength = Split(aFieldName(2), "(")
        Combo1.Text = aFieldNameLength(0)
        Text2.Text = Replace(aFieldNameLength(1), ")", "")
        
        If aFieldName(3) = "P" Then
            Check1.Value = 1
            Check2.Value = 0
        ElseIf aFieldName(3) = "A" Then
            Check1.Value = 0
            Check2.Value = 1
        ElseIf aFieldName(3) = "PA" Then
            Check1.Value = 1
            Check2.Value = 1
        Else
            Check1.Value = 0
            Check2.Value = 0
        End If
    
        Command3.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
    Else
        Command3.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
    End If

End Sub

Private Sub Text1_GotFocus()

    Command4.Default = True

End Sub

Private Sub Text2_GotFocus()

    SendKeys "{END}+{HOME}"

End Sub
