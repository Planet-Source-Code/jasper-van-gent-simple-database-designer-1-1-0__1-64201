VERSION 5.00
Begin VB.Form FieldTypeWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Fieldtype"
   ClientHeight    =   2892
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3528
   ControlBox      =   0   'False
   Icon            =   "FieldType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   3528
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   972
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Auto_increment"
      Height          =   252
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2160
      TabIndex        =   5
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Primary Key"
      Height          =   252
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   852
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1932
   End
   Begin VB.Label Label4 
      Height          =   252
      Left            =   1320
      TabIndex        =   9
      Top             =   120
      Width           =   1932
   End
   Begin VB.Label Label3 
      Caption         =   "Object:"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Size:"
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Datatype:"
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "FieldTypeWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ExtraValue As String

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

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
    
    Editor.FieldType(Label4.Caption).Caption = UCase(Combo1.Text) & "(" & Text1.Text & ")" & " " & ExtraValue
    
    Unload Me

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
