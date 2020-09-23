VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Export_SQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export project to SQL Script"
   ClientHeight    =   3756
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3504
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3756
   ScaleWidth      =   3504
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1200
      TabIndex        =   1
      Text            =   "test"
      Top             =   2640
      Width           =   2052
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1440
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2160
      TabIndex        =   3
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      Default         =   -1  'True
      Height          =   372
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export tables"
      Height          =   2172
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3012
      Begin VB.ListBox List1 
         Height          =   1560
         Left            =   240
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   360
         Width           =   2532
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Database:"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   852
   End
End
Attribute VB_Name = "Export_SQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'define strings
Dim Auto_Inc As String
Dim Primary_Key As String
Dim SetEnd As String

'define integers
Dim iTables As Integer
Dim iFields As Integer
Dim iNewFields As Integer

'define arrays
Dim aTable As Variant
Dim aFieldTypeExtra As Variant

Private Sub Command1_Click()

    On Error GoTo Err:
    
    If List1.SelCount = 0 Then
        MsgBox "You must have selected at  least one table!", vbCritical + vbOKOnly, "Export error"
    Else
        With CommonDialog1
            .CancelError = True
            .Flags = cdlOFNOverwritePrompt
            .Filter = "Structured Query Language (*.sql)|*.sql|"
            .ShowSave
            
            Open .FileName For Output As #1
                'print default comment
                Print #1, "-- SQL Script created with Simple Database Designer v" & App.Major & "." & App.Minor & "." & App.Revision
                Print #1, "-- Date/time: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
                Print #1, ""
                'get all selected tables
                For i = 0 To List1.ListCount - 1
                    'only continue if listitem is selected
                    If List1.Selected(i) = True Then
                        aTable = Split(List1.List(i), ".")
                        'get selected table properties
                        Print #1, "CREATE TABLE " & LCase(Text1.Text) & "." & LCase(aTable(1)) & " ("
                        
                        'reset primary_key value
                        Primary_Key = ""
                        
                        'count fields from current table
                        iFields = 0
                        For j = 0 To Editor.FieldName.Count - 1
                            If Editor.FieldName(j).ToolTipText = "TABLE." & aTable(0) Then
                                iFields = iFields + 1
                            End If
                        Next j
                        
                        'get fields
                        iNewFields = 0
                        For j = 0 To Editor.FieldName.Count - 1
                            If Editor.FieldName(j).ToolTipText = "TABLE." & aTable(0) Then
                                iNewFields = iNewFields + 1
                                Auto_Inc = ""
                                
                                'save fields
                                aFieldTypeExtra = Split(Editor.FieldType(j).Caption, " ")
                                'only auto_increment
                                If aFieldTypeExtra(1) = "A" Then
                                    Auto_Inc = " AUTO_INCREMENT"
                                'primary key and auto_increment
                                ElseIf aFieldTypeExtra(1) = "PA" Then
                                    Auto_Inc = " AUTO_INCREMENT"
                                    If Primary_Key = "" Then
                                        Primary_Key = Editor.FieldName(j).Caption
                                    Else
                                        Primary_Key = Primary_Key & ", " & Editor.FieldName(j).Caption
                                    End If
                                'only primary key
                                ElseIf aFieldTypeExtra(1) = "P" Then
                                    If Primary_Key = "" Then
                                        Primary_Key = Editor.FieldName(j).Caption
                                    Else
                                        Primary_Key = Primary_Key & ", " & Editor.FieldName(j).Caption
                                    End If
                                End If
                                
                                'check if there is a comma required at the end or not
                                If iNewFields < iFields Then
                                    SetEnd = ","
                                Else
                                    SetEnd = ""
                                End If
                                
                                'save field properties
                                Print #1, "  " & Editor.FieldName(j).Caption & " " & UCase(aFieldTypeExtra(0)) & " NOT NULL" & Auto_Inc & SetEnd
                                
                            End If

                        Next j
                        
                        'save primary key
                        If Primary_Key <> "" Then
                            Print #1, "  PRIMARY KEY(" & Primary_Key & ")"
                        End If
                        
                        Print #1, ")" & vbCrLf
                    End If
                Next i

            Close #1
            
            Unload Me
        
        End With
    End If

Err:
    Exit Sub

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Editor.Enabled = True
    Editor.Show

End Sub
