VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Properties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   4764
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6264
   ControlBox      =   0   'False
   Icon            =   "Properties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4764
   ScaleWidth      =   6264
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4200
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   372
      Left            =   3840
      TabIndex        =   0
      Top             =   4200
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   5040
      TabIndex        =   1
      Top             =   4200
      Width           =   1092
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3852
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6012
      _ExtentX        =   10605
      _ExtentY        =   6795
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Properties.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Relations"
      TabPicture(1)   =   "Properties.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Label"
         Height          =   1572
         Left            =   -74760
         TabIndex        =   23
         Top             =   1920
         Width           =   5412
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00000000&
            Height          =   252
            Left            =   1920
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   28
            Top             =   720
            Width           =   372
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00000000&
            Height          =   252
            Left            =   1920
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   27
            Top             =   360
            Width           =   372
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Label border"
            Height          =   252
            Left            =   1920
            TabIndex        =   26
            Top             =   1080
            Width           =   1332
         End
         Begin VB.Label Label10 
            Caption         =   "Text color:"
            Height          =   252
            Left            =   360
            TabIndex        =   25
            Top             =   720
            Width           =   852
         End
         Begin VB.Label Label7 
            Caption         =   "Background color:"
            Height          =   252
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   1452
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Line layout"
         Height          =   1332
         Left            =   -74760
         TabIndex        =   18
         Top             =   480
         Width           =   5412
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00000000&
            Height          =   252
            Left            =   1320
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   22
            Top             =   720
            Width           =   372
         End
         Begin VB.ComboBox Combo2 
            Height          =   288
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   1812
         End
         Begin VB.Label Label9 
            Caption         =   "Line color:"
            Height          =   252
            Left            =   360
            TabIndex        =   21
            Top             =   720
            Width           =   852
         End
         Begin VB.Label Label8 
            Caption         =   "Line style:"
            Height          =   252
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   852
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Colors"
         Height          =   1212
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   3972
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00000000&
            Height          =   252
            Left            =   3000
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   16
            Top             =   360
            Width           =   372
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFC0C0&
            Height          =   252
            Left            =   1440
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   14
            Top             =   720
            Width           =   372
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   252
            Left            =   1440
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   13
            Top             =   360
            Width           =   372
         End
         Begin VB.Label Label5 
            Caption         =   "Tablename:"
            Height          =   252
            Left            =   2040
            TabIndex        =   15
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label4 
            Caption         =   "Table:"
            Height          =   252
            Left            =   360
            TabIndex        =   12
            Top             =   720
            Width           =   1092
         End
         Begin VB.Label Label3 
            Caption         =   "Background:"
            Height          =   252
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Width           =   1092
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sizes"
         Height          =   1572
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   3972
         Begin VB.CommandButton Command3 
            Caption         =   "Flip"
            Height          =   372
            Left            =   1080
            TabIndex        =   5
            Top             =   1080
            Width           =   972
         End
         Begin VB.ComboBox Combo1 
            Height          =   288
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   852
         End
         Begin VB.TextBox Text2 
            Height          =   288
            Left            =   1080
            TabIndex        =   3
            Text            =   "0"
            Top             =   720
            Width           =   972
         End
         Begin VB.TextBox Text1 
            Height          =   288
            Left            =   1080
            TabIndex        =   2
            Text            =   "0"
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label6 
            Caption         =   "Preset:"
            Height          =   252
            Left            =   2280
            TabIndex        =   17
            Top             =   600
            Width           =   612
         End
         Begin VB.Label Label2 
            Caption         =   "Height:"
            Height          =   252
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   612
         End
         Begin VB.Label Label1 
            Caption         =   "Width:"
            Height          =   252
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   492
         End
      End
   End
End
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OldText1 As Integer
Dim OldText2 As Integer

Private Sub Combo1_Click()

    If Combo1.Text = "A2" Then
        Text1.Text = 420
        Text2.Text = 594
    ElseIf Combo1.Text = "A3" Then
        Text1.Text = 297
        Text2.Text = 420
    ElseIf Combo1.Text = "A4" Then
        Text1.Text = 210
        Text2.Text = 297
    ElseIf Combo1.Text = "A5" Then
        Text1.Text = 148
        Text2.Text = 210
    ElseIf Combo1.Text = "A6" Then
        Text1.Text = 105
        Text2.Text = 148
    End If

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    Editor.EditField.BackColor = Picture1.BackColor
    Editor.SetColor.BackColor = Picture2.BackColor
    If Text1.Text > 5 Then
        Editor.EditField.Width = Text1.Text / 10
    Else
        Editor.EditField.Width = 5
    End If
    If Text2.Text > 5 Then
        Editor.EditField.Height = Text2.Text / 10
    Else
        Editor.EditField.Height = 5
    End If
    Editor.StatusBar1.Panels(2).Text = "Size: " & Round(Editor.EditField.Width, 2) & " x " & Round(Editor.EditField.Height, 2)
    
    'change tablename backcolors
    For i = 0 To Editor.TableName.Count - 1
        Editor.TableName(i).BackColor = Picture3.BackColor
    Next i
    
    'change relation colors
    For i = 0 To Editor.Relation_1.Count - 1
        If Combo2.Text = "Solid" Then
            Editor.Relation_1(i).BorderStyle = 1
        ElseIf Combo2.Text = "Dashed" Then
            Editor.Relation_1(i).BorderStyle = 2
        ElseIf Combo2.Text = "Dotted" Then
            Editor.Relation_1(i).BorderStyle = 3
        End If
        Editor.Relation_1(i).BorderColor = Picture4.BackColor
    Next i
    For i = 0 To Editor.Relation_Caption.Count - 1
        Editor.Relation_Caption(i).BackColor = Picture5.BackColor
        Editor.Relation_Caption(i).ForeColor = Picture6.BackColor
        Editor.Relation_Caption(i).BorderStyle = Check1.Value
    Next i
    
    ResizeElements
    Unload Me

End Sub

Private Sub Command3_Click()

    OldText1 = Text1.Text
    OldText2 = Text2.Text
    
    Text1.Text = OldText2
    Text2.Text = OldText1

End Sub

Private Sub Form_Load()

    Combo1.AddItem "A2"
    Combo1.AddItem "A3"
    Combo1.AddItem "A4"
    Combo1.AddItem "A5"
    Combo1.AddItem "A6"
    
    Combo2.AddItem "Solid"
    Combo2.AddItem "Dashed"
    Combo2.AddItem "Dotted"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Editor.Enabled = True
    Editor.Show

End Sub

Private Sub Picture1_Click()
    
    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture1.BackColor
        .ShowColor
        Picture1.BackColor = .Color
    End With
Err:
    Exit Sub
    
End Sub

Private Sub Picture2_Click()
    
    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture2.BackColor
        .ShowColor
        Picture2.BackColor = .Color
    End With
Err:
    Exit Sub
    
End Sub

Private Sub Picture3_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture3.BackColor
        .ShowColor
        Picture3.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Picture4_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture4.BackColor
        .ShowColor
        Picture4.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Picture5_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture5.BackColor
        .ShowColor
        Picture5.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Picture6_Click()


    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Flags = 1
        .Color = Picture6.BackColor
        .ShowColor
        Picture6.BackColor = .Color
    End With
Err:
    Exit Sub


End Sub
