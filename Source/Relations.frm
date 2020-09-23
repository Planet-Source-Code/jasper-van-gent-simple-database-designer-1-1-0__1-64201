VERSION 5.00
Begin VB.Form Relations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Build relation"
   ClientHeight    =   3216
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3216
   ScaleWidth      =   6144
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   840
      TabIndex        =   6
      Text            =   "Relation 1"
      Top             =   2640
      Width           =   1932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   372
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Frame Frame2 
      Caption         =   "Create relation between"
      Height          =   2172
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5652
      Begin VB.ListBox List2 
         Height          =   1584
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   2532
      End
      Begin VB.ListBox List1 
         Height          =   1584
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2532
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4800
      TabIndex        =   0
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Title:"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   492
   End
End
Attribute VB_Name = "Relations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    aTable1 = Split(List1.Text, ".")
    aTable2 = Split(List2.Text, ".")

    If aTable1(1) <> aTable2(1) Then
        Editor.List2.Text = Text1.Text
          If Editor.List2.Text = Text1.Text Then
            MsgBox "There is already a relation with this name!", vbCritical + vbOKOnly, "Create relation"
        Else
            
            ' display relation inside relations list
            Editor.List2.AddItem Text1.Text
            
            ' add relation properties into relation properties list
            Editor.Relation_Properties.AddItem List1.Text & "-" & List2.Text
            
            ' create relation in editor
            Load Editor.Relation_1(Editor.Relation_1.Count)
            Load Editor.Relation_Caption(Editor.Relation_Caption.Count)
            
            'set relation line from first table
            With Editor.Relation_1(Editor.Relation_1.UBound)
                .X1 = Editor.FieldType(aTable1(0)).Left + Editor.FieldType(aTable1(0)).Width + 0.05 'left position
                .Y1 = Editor.FieldType(aTable1(0)).Top + (Editor.FieldType(aTable1(0)).Height / 2) 'top position
                .X2 = Editor.FieldType(aTable2(0)).Left + Editor.FieldType(aTable1(0)).Width + 0.05 'right width
                .Y2 = Editor.FieldType(aTable2(0)).Top + (Editor.FieldType(aTable2(0)).Height / 2) 'left bottom
                .Tag = Text1.Text
                .Visible = True
                .ZOrder (0)
            End With
            
            With Editor.Relation_Caption(Editor.Relation_Caption.UBound)
                .Caption = Text1.Text
            End With
            With Editor.Relation_Caption(Editor.Relation_Caption.UBound)
                .Left = Editor.Relation_1(Editor.Relation_1.UBound).X1 + ((Editor.Relation_1(Editor.Relation_1.UBound).X2 - Editor.Relation_1(Editor.Relation_1.UBound).X1) / 2) - (Len(.Caption) / 12) / 2
                .Top = Editor.Relation_1(Editor.Relation_1.UBound).Y1 + ((Editor.Relation_1(Editor.Relation_1.UBound).Y2 - Editor.Relation_1(Editor.Relation_1.UBound).Y1) / 2)
                .ZOrder (0)
                .Visible = True
            End With
            
            Unload Me
        End If
    Else
        MsgBox "Cannot create a relation with the same table!", vbCritical + vbOKOnly, "Create relation"
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Editor.Enabled = True
    Editor.Show

End Sub

Private Sub List1_Click()

    
    If List1.Text <> List2.Text Then
        Command2.Enabled = True
    End If

End Sub

Private Sub List2_Click()

    If List1.Text <> List2.Text Then
        Command2.Enabled = True
    End If

End Sub
