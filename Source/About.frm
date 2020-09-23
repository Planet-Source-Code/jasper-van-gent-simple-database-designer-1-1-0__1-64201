VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   3912
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   5352
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3912
   ScaleWidth      =   5352
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   852
      Left            =   0
      ScaleHeight     =   852
      ScaleWidth      =   5352
      TabIndex        =   3
      Top             =   0
      Width           =   5352
      Begin VB.Image Image1 
         Height          =   348
         Left            =   360
         Picture         =   "About.frx":000C
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   1332
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1332
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   4212
   End
   Begin VB.Label Label2 
      Caption         =   "Created by Jasper van Gent"
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   3252
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Label3.Caption = "Simple Database Designer was developed with Visual Basic 6. Use this software at your one risk. This is still just a beta release. Check our website for more information and updates." & vbCrLf & vbCrLf & "http://www.webrazor.nl/"
    Label3.Caption = "Simple Database Designer is a easy to use application for creating simple database models. This is a final release. Check our website for more information and updates." & vbCrLf & vbCrLf & "http://www.webrazor.nl/"
    Command1.Left = (Me.Width - Command1.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Editor.Enabled = True
    Editor.Show

End Sub
