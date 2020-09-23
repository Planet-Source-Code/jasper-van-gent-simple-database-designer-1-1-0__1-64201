VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Editor 
   Caption         =   "Simple Database Designer"
   ClientHeight    =   7908
   ClientLeft      =   48
   ClientTop       =   732
   ClientWidth     =   12888
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7908
   ScaleWidth      =   12888
   Begin MSComctlLib.ImageList Bitmaps 
      Left            =   8040
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   0
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":1708A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":173DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Cursors 
      Left            =   8040
      Top             =   6120
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":1772E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":17A48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   6240
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12888
      _ExtentX        =   22733
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   7656
      Width           =   12888
      _ExtentX        =   22733
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11980
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Size: 0 x 0"
            TextSave        =   "Size: 0 x 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Tables: 0"
            TextSave        =   "Tables: 0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Fields: 0"
            TextSave        =   "Fields: 0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox WorkArea 
      Align           =   3  'Align Left
      BackColor       =   &H8000000C&
      Height          =   7368
      Left            =   2052
      ScaleHeight     =   12.912
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   16.849
      TabIndex        =   1
      Top             =   288
      Width           =   9600
      Begin VB.HScrollBar HScroll1 
         Height          =   252
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1452
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1812
         Left            =   9240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   0
         Width           =   252
      End
      Begin VB.PictureBox EmptySpace 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6000
         ScaleHeight     =   252
         ScaleWidth      =   252
         TabIndex        =   22
         Top             =   5160
         Width           =   252
      End
      Begin VB.PictureBox EditField 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5669
         Left            =   120
         ScaleHeight     =   9.948
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   14.965
         TabIndex        =   3
         Top             =   120
         Width           =   8504
         Begin VB.PictureBox MoveTable2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   227
            Left            =   720
            Picture         =   "Editor.frx":17D62
            ScaleHeight     =   0.402
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   0.402
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   227
         End
         Begin VB.PictureBox MoveTable 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000008&
            Height          =   227
            Left            =   2160
            Picture         =   "Editor.frx":18051
            ScaleHeight     =   0.402
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   0.402
            TabIndex        =   24
            Top             =   2040
            Visible         =   0   'False
            Width           =   227
         End
         Begin VB.PictureBox BitmapFile 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   3720
            ScaleHeight     =   0.614
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   0.614
            TabIndex        =   23
            Top             =   1680
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.Label Relation_Caption 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   168
            Index           =   0
            Left            =   4080
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   1200
            Visible         =   0   'False
            Width           =   468
         End
         Begin VB.Line Relation_1 
            Index           =   0
            Visible         =   0   'False
            X1              =   5.927
            X2              =   7.62
            Y1              =   1.27
            Y2              =   1.27
         End
         Begin VB.Label FieldName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   168
            Index           =   0
            Left            =   960
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   1440
            Visible         =   0   'False
            Width           =   336
         End
         Begin VB.Label FieldType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "VARCHAR(45) "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   168
            Index           =   0
            Left            =   960
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   1200
            Visible         =   0   'False
            Width           =   1332
         End
         Begin VB.Label TableName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Untitled"
            BeginProperty Font 
               Name            =   "Lucida Sans"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   221
            Index           =   0
            Left            =   840
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   720
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.Label TableSelector 
            BackStyle       =   0  'Transparent
            Height          =   1332
            Index           =   0
            Left            =   840
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   720
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.Shape Table 
            BorderColor     =   &H00000000&
            FillColor       =   &H00FFD9D9&
            FillStyle       =   0  'Solid
            Height          =   1701
            Index           =   0
            Left            =   720
            Top             =   600
            Visible         =   0   'False
            Width           =   1701
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00000000&
            BorderStyle     =   3  'Dot
            Height          =   1932
            Left            =   600
            Top             =   480
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.Shape TableShadow 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            Height          =   612
            Index           =   0
            Left            =   720
            Top             =   2520
            Visible         =   0   'False
            Width           =   732
         End
      End
      Begin VB.PictureBox Shadow 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3012
         Left            =   240
         ScaleHeight     =   5.313
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   8.7
         TabIndex        =   5
         Top             =   240
         Width           =   4932
      End
   End
   Begin VB.PictureBox WorkBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7368
      Left            =   0
      ScaleHeight     =   7368
      ScaleWidth      =   2052
      TabIndex        =   0
      Top             =   288
      Width           =   2052
      Begin VB.ListBox List1 
         Height          =   1776
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Table list"
         Top             =   2280
         Width           =   1812
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         Height          =   1812
         Left            =   120
         ScaleHeight     =   1764
         ScaleWidth      =   1764
         TabIndex        =   6
         Top             =   120
         Width           =   1812
         Begin VB.PictureBox OrgColor 
            BackColor       =   &H00FFD9D9&
            Height          =   252
            Left            =   1200
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   16
            ToolTipText     =   "Original color"
            Top             =   840
            Width           =   372
         End
         Begin VB.PictureBox SetColor 
            BackColor       =   &H00FFD9D9&
            Height          =   252
            Left            =   720
            ScaleHeight     =   204
            ScaleWidth      =   324
            TabIndex        =   13
            ToolTipText     =   "New color"
            Top             =   840
            Width           =   372
         End
         Begin VB.CommandButton SetApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   372
            Left            =   720
            TabIndex        =   11
            Top             =   1200
            Width           =   732
         End
         Begin VB.TextBox Setheight 
            Height          =   288
            Left            =   720
            TabIndex        =   10
            Text            =   "3"
            Top             =   480
            Width           =   852
         End
         Begin VB.TextBox SetWidth 
            Height          =   288
            Left            =   720
            TabIndex        =   9
            Text            =   "3"
            Top             =   120
            Width           =   852
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   252
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   612
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            Height          =   252
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   612
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            Height          =   252
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   612
         End
      End
      Begin VB.ListBox List2 
         Height          =   1200
         Left            =   120
         TabIndex        =   26
         Top             =   4440
         Width           =   1812
      End
      Begin VB.ListBox Relation_Properties 
         Height          =   432
         Left            =   120
         TabIndex        =   29
         Top             =   4440
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000004&
         Caption         =   "Tables"
         Height          =   252
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   1812
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         Caption         =   "Relations"
         Height          =   252
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   1812
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4536
      Top             =   3000
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":18344
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":18456
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":18568
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.frx":1867A
            Key             =   "Properties"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileEmpty4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export project to"
         Begin VB.Menu mnuFileExportSQL 
            Caption         =   "SQL Script"
         End
         Begin VB.Menu mnuFileExportEmpty1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileExportDrawing 
            Caption         =   "Drawing..."
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu mnuFileEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditMovers 
         Caption         =   "&Show movers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFront 
         Caption         =   "Bring to Front"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "Send to Back"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuEditEmpty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDeselect 
         Caption         =   "&Deselect"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditEmpty3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectCreateTable 
         Caption         =   "&Create table"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuProjectTableFields 
         Caption         =   "&Table fields"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuProjectEmpty3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectRelation 
         Caption         =   "Build relation..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuProjectEmpty4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectProperties 
         Caption         =   "&Properties..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'define string variables
Dim TableNameTitle As String
Dim FieldNameTitle As String
Dim FieldTypeTitle As String
Dim RelationTitle As String
Dim OldRelationTitle As String
Dim CurrentTableLeft
Dim CurrentTableTop
Dim MoveObject As String
Dim RemovedRelation As String

'define arrays
Dim aFieldType As Variant
Dim aFieldTypeExtra As Variant
Dim aRemovedRelations As Variant

Private Sub EditField_Click()

    DeSelect

End Sub

Private Sub EditField_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

    If MoveObject = "Bottom" Then
        MoveTable.ZOrder (0)
        
        MoveTable.Left = X - CurrentTableLeft
        MoveTable.Top = Y - CurrentTableTop
        
        ' table and shape positions
        Table(ObjectIndex).Left = (MoveTable.Left - Table(ObjectIndex).Width)
        Table(ObjectIndex).Top = (MoveTable.Top - Table(ObjectIndex).Height)
        TableShadow(ObjectIndex).Left = (MoveTable.Left - TableShadow(ObjectIndex).Width) + 0.01
        TableShadow(ObjectIndex).Top = (MoveTable.Top - TableShadow(ObjectIndex).Height) + 0.01
        TableName(ObjectIndex).Left = (MoveTable.Left - Table(ObjectIndex).Width)
        TableName(ObjectIndex).Top = (MoveTable.Top - Table(ObjectIndex).Height)
        TableSelector(ObjectIndex).Left = (MoveTable.Left - Table(ObjectIndex).Width)
        TableSelector(ObjectIndex).Top = (MoveTable.Top - Table(ObjectIndex).Height)
        Shape.Left = ((MoveTable.Left - Table(ObjectIndex).Width) - 0.05)
        Shape.Top = ((MoveTable.Top - Table(ObjectIndex).Height) - 0.05)
        MoveTable2.Left = Table(ObjectIndex).Left - MoveTable2.Width
        MoveTable2.Top = Table(ObjectIndex).Top - MoveTable2.Height
        
        j = 0
        ' move fields
        For i = 0 To FieldName.Count - 1
            If FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                j = j + 1
                
                ' move field
                If FieldName(i).Visible = True Then
                    FieldName(i).Left = Table(ObjectIndex).Left + 0.05
                    FieldName(i).Top = (Table(ObjectIndex).Top + TableName(ObjectIndex).Height + ((FieldName(i).Height + 0.05) * (j - 1))) + 0.05
                    FieldType(i).Left = (Table(ObjectIndex).Left + FieldName(i).Width) + 0.1
                    FieldType(i).Top = (Table(ObjectIndex).Top + TableName(ObjectIndex).Height + ((FieldName(i).Height + 0.05) * (j - 1))) + 0.05
                End If
            End If
        Next i
        MoveRelations

    ElseIf MoveObject = "Top" Then
        MoveTable2.ZOrder (0)
        
        ' table and shape positions
        MoveTable2.Left = X - CurrentTableLeft
        MoveTable2.Top = Y - CurrentTableTop
        Table(ObjectIndex).Left = MoveTable2.Left + MoveTable2.Width
        Table(ObjectIndex).Top = MoveTable2.Top + MoveTable2.Height
        TableShadow(ObjectIndex).Left = (MoveTable2.Left + MoveTable2.Width) + 0.03
        TableShadow(ObjectIndex).Top = (MoveTable2.Top + MoveTable2.Height) + 0.03
        TableName(ObjectIndex).Left = MoveTable2.Left + MoveTable2.Width
        TableName(ObjectIndex).Top = MoveTable2.Top + MoveTable2.Height
        TableSelector(ObjectIndex).Left = MoveTable2.Left + MoveTable2.Width
        TableSelector(ObjectIndex).Top = MoveTable2.Top + MoveTable2.Height
        Shape.Left = (MoveTable2.Left + MoveTable2.Width) - 0.05
        Shape.Top = (MoveTable2.Top + MoveTable2.Height) - 0.05
        MoveTable.Left = Table(ObjectIndex).Left + Table(ObjectIndex).Width
        MoveTable.Top = Table(ObjectIndex).Top + Table(ObjectIndex).Height
        
        j = 0
        ' move fields
        For i = 0 To FieldName.Count - 1
            If FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                j = j + 1
                
                ' move field
                If FieldName(i).Visible = True Then
                    FieldName(i).Left = Table(ObjectIndex).Left + 0.05
                    FieldName(i).Top = (Table(ObjectIndex).Top + TableName(ObjectIndex).Height + ((FieldName(i).Height + 0.05) * (j - 1))) + 0.05
                    FieldType(i).Left = (Table(ObjectIndex).Left + FieldName(i).Width) + 0.1
                    FieldType(i).Top = (Table(ObjectIndex).Top + TableName(ObjectIndex).Height + ((FieldName(i).Height + 0.05) * (j - 1))) + 0.05
                End If
            End If
        Next i
        MoveRelations
    End If

End Sub

Private Sub EditField_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Then
        MoveObjectLeft
    ElseIf KeyCode = vbKeyRight Then
        MoveObjectRight
    ElseIf KeyCode = vbKeyUp Then
        MoveObjectUp
    ElseIf KeyCode = vbKeyDown Then
        MoveObjectDown
    End If

End Sub

Private Sub EditField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(5).Text = Round(X, 1) * 10 & " x " & Round(Y, 1) * 10
End Sub

Private Sub EditField_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuProject
    End If

End Sub

Private Sub FieldName_Click(Index As Integer)

    DeselectFields
    
    FieldName(Index).BorderStyle = 1
    
    ObjectName = "FieldName"
    ObjectIndex = Index

End Sub

Private Sub FieldName_DblClick(Index As Integer)

    FieldNameTitle = InputBox("Enter the new FieldName:", "FieldName", FieldName(Index).Caption)
    
    If FieldNameTitle <> "" Then
        FieldName(Index).Caption = FieldNameTitle
    End If

End Sub

Private Sub FieldType_Click(Index As Integer)
    
    DeselectFields
    FieldType(Index).BorderStyle = 1
    
    ObjectName = "FieldName"
    ObjectIndex = Index

End Sub

Private Sub FieldType_DblClick(Index As Integer)

    FieldTypeWizard.Icon = Me.Icon
    Me.Enabled = False
    FieldTypeWizard.Label4.Caption = ObjectIndex
    
    aFieldType = Split(FieldType(ObjectIndex).Caption, "(")
    FieldTypeWizard.Combo1.Text = UCase(aFieldType(0))
    
    aFieldTypeExtra = Split(Replace(aFieldType(1), ")", ""), " ")
    FieldTypeWizard.Text1.Text = Replace(aFieldTypeExtra(0), ")", "")
    
    If aFieldTypeExtra(1) = "P" Then
        FieldTypeWizard.Check1.Value = 1
    ElseIf aFieldTypeExtra(1) = "A" Then
        FieldTypeWizard.Check2.Value = 1
    ElseIf aFieldTypeExtra(1) = "PA" Then
        FieldTypeWizard.Check1.Value = 1
        FieldTypeWizard.Check2.Value = 1
    End If
    
    FieldTypeWizard.Show

End Sub

Private Sub Form_Resize()

    ResizeElements

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub HScroll1_Change()

    EditField.Left = 0.2 - HScroll1.Value
    ResizeElements

End Sub

Private Sub HScroll1_Scroll()

    EditField.Left = 0.2 - HScroll1.Value
    ResizeElements

End Sub

Private Sub List2_Click()


    If Relation_Properties.List(List2.ListIndex) <> "" Then
        
        'select lines
        For i = 1 To Relation_1.Count - 1
            Relation_1(i).BorderWidth = 1
            Relation_Caption(i).FontBold = False
            If Relation_1(i).Tag = List2.Text Then
                Relation_1(i).BorderWidth = 2
                Relation_Caption(i).FontBold = True
            End If
        Next i
    
    End If

End Sub

Private Sub mnuEditBack_Click()

    If ObjectName = "Table" Then
        
        'reorder fields
        For i = 0 To FieldName.Count - 1
            If FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                ' resize
                FieldName(i).ZOrder (1)
                FieldType(i).ZOrder (1)
            End If
        Next i
    
        TableName(ObjectIndex).ZOrder (1)
        Table(ObjectIndex).ZOrder (1)
        TableShadow(ObjectIndex).ZOrder (1)
        List1.RemoveItem List1.ListIndex
        List1.AddItem "TABLE." & ObjectIndex, 0
    End If
    
    DeSelect

End Sub

Private Sub mnuEditDelete_Click()

    If mnuEditDelete.Caption = "Delete table" Then

        If MsgBox("Are you sure?" & vbCrLf & "(all relations with this table will also be removed)", vbQuestion + vbYesNo, "Delete table") = vbYes Then
            
            Table(ObjectIndex).Visible = False
            TableName(ObjectIndex).Visible = False
            TableSelector(ObjectIndex).Visible = False
            TableShadow(ObjectIndex).Visible = False
            List1.RemoveItem List1.ListIndex
            
            For i = 0 To FieldName.Count - 1
                If FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
    
                    ' delete field
                    FieldName(i).Visible = False
                    FieldType(i).Visible = False
                End If
            Next i
            
            ' delete all relations
            For i = 0 To List2.ListCount - 1
            
                aTable = Split(Relation_Properties.List(i), ".")
    
                If aTable(1) = TableName(ObjectIndex).Caption Then
                    For j = 1 To Relation_1.Count - 1
                        If Relation_1(j).Tag = List2.List(i) Then
                            Relation_1(j).Visible = False
                            Relation_Caption(j).Visible = False
                            RemovedRelation = RemovedRelation & "|" & i
                        End If
                    Next j
                ElseIf aTable(3) = TableName(ObjectIndex).Caption Then
                    For j = 1 To Relation_1.Count - 1
                        If Relation_1(j).Tag = List2.List(i) Then
                            Relation_1(j).Visible = False
                            Relation_Caption(j).Visible = False
                            RemovedRelation = RemovedRelation & "|" & i
                        End If
                    Next j
                End If
            
            Next i
            
            'remove list items
            aRemovedRelations = Split(RemovedRelation, "|")
    
            'reverse order
            For i = UBound(aRemovedRelations) To 1 Step -1
                List2.RemoveItem aRemovedRelations(i)
            Next i
            
            RemovedRelation = ""
            
        End If
        
        DeSelect

    ElseIf mnuEditDelete.Caption = "Delete relation" Then

        If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Delete relation") = vbYes Then
            'select lines
            For i = 1 To Relation_1.Count - 1
                If Relation_1(i).Tag = List2.Text Then
                    Relation_1(i).Visible = False
                    Relation_Caption(i).Visible = False
                End If
            Next i
            List2.RemoveItem List2.ListIndex
        End If

    End If

End Sub

Private Sub mnuEditDeselect_Click()

    DeSelect

End Sub

Private Sub mnuEditFront_Click()

    If ObjectName = "Table" Then
        TableShadow(ObjectIndex).ZOrder (0)
        Table(ObjectIndex).ZOrder (0)
        TableName(ObjectIndex).ZOrder (0)
        List1.RemoveItem List1.ListIndex
        List1.AddItem "TABLE." & ObjectIndex, List1.ListCount
        
        'reorder fields
        For i = 0 To FieldName.Count - 1
            If FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                ' resize
                FieldName(i).ZOrder (0)
                FieldType(i).ZOrder (0)
            End If
        Next i
        
    End If
    
    DeSelect

End Sub

Private Sub mnuEditMovers_Click()

    If mnuEditMovers.Checked = True Then
        mnuEditMovers.Checked = False
    Else
        mnuEditMovers.Checked = True
    End If

End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileExportDrawing_Click()

    ExportDrawing.Icon = Me.Icon
    
    ExportDrawing.Show
    Me.Enabled = False

End Sub

Private Sub mnuFileExportSQL_Click()

    Export_SQL.Icon = Me.Icon
    Me.Enabled = False
    
    ' count number of tables
    For i = 1 To Table.Count - 1
        Export_SQL.List1.AddItem TableName(i).Index & "." & TableName(i).Caption
    Next i
    
    ' select all tables
    For i = 0 To Export_SQL.List1.ListCount - 1
        Export_SQL.List1.Selected(i) = True
    Next i
    
    Export_SQL.Show

End Sub

Private Sub mnuFileNew_Click()

    NewFile

End Sub

Private Sub mnuFileOpen_Click()
    
    OpenFile
    
End Sub

Private Sub mnuFileSave_Click()

    SaveFile (FileName)

End Sub

Private Sub mnuFileSaveAs_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Filter = "Simple Database Designer (*.sdd)|*.sdd|"
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        FileName = .FileName
        SaveFile (FileName)
        mnuFileSave.Enabled = True
        Toolbar1.Buttons(4).Enabled = True
    End With
Err:
    Exit Sub

End Sub

Private Sub mnuHelpAbout_Click()

    About.Icon = Me.Icon
    Me.Enabled = False
    About.Show

End Sub

Private Sub mnuProjectCreateTable_Click()

    CreateTableWizard.Icon = Me.Icon
    
    'create list with added tables
    For i = 1 To Editor.Table.Count - 1
        CreateTableWizard.List1.AddItem i & "." & TableName(i).Caption
    Next i
    
    If ObjectName = "Table" Then
        'select current selected table
        CreateTableWizard.List1.Text = ObjectIndex & "." & TableName(ObjectIndex).Caption
    End If
    

    'create positions list
    CreateTableWizard.Combo1.AddItem "Below"
    CreateTableWizard.Combo1.AddItem "Above"
    CreateTableWizard.Combo1.AddItem "Left"
    CreateTableWizard.Combo1.AddItem "Right"
    
    'select default position
    CreateTableWizard.Combo1.Text = "Right"

    'enable elements
    If Editor.Table.ubound = 0 Then
        CreateTableWizard.Combo1.Enabled = False
        CreateTableWizard.List1.Enabled = False
        CreateTableWizard.Frame1.Enabled = False
        CreateTableWizard.Label1.Enabled = False
        CreateTableWizard.Label2.Enabled = False
        CreateTableWizard.Text1.Enabled = False
    End If
        
    
    CreateTableWizard.Show
    Me.Enabled = False

End Sub

Private Sub mnuProjectProperties_Click()

    Me.Enabled = False
    
    Properties.Icon = Me.Icon
    Properties.Text1.Text = Round(EditField.Width, 2) * 10
    Properties.Text2.Text = Round(EditField.Height, 2) * 10
    Properties.Picture1.BackColor = EditField.BackColor
    Properties.Picture2.BackColor = SetColor.BackColor
    Properties.Picture3.BackColor = TableName(TableName.ubound).BackColor
    
    'relation properties
    If Relation_1(Relation_1.ubound).BorderStyle = 1 Then
        Properties.Combo2.Text = "Solid"
    ElseIf Relation_1(Relation_1.ubound).BorderStyle = 2 Then
        Properties.Combo2.Text = "Dashed"
    ElseIf Relation_1(Relation_1.ubound).BorderStyle = 3 Then
        Properties.Combo2.Text = "Dotted"
    End If
    Properties.Picture4.BackColor = Relation_1(Relation_1.ubound).BorderColor
    Properties.Picture5.BackColor = Relation_Caption(Relation_Caption.ubound).BackColor
    Properties.Picture6.BackColor = Relation_Caption(Relation_Caption.ubound).ForeColor
    Properties.Check1.Value = Relation_Caption(Relation_1.ubound).BorderStyle
    
    Properties.Show

End Sub

Private Sub mnuProjectRelation_Click()

    Relations.Icon = Me.Icon
    
    For i = 1 To TableName.Count - 1
        'check for field
        For j = 0 To FieldName.Count - 1
            If FieldName(j).ToolTipText = "TABLE." & i Then
                Relations.List1.AddItem j & "." & TableName(i).Caption & "." & FieldName(j).Caption
                Relations.List2.AddItem j & "." & TableName(i).Caption & "." & FieldName(j).Caption
            End If
        Next j
    Next i
    
    Relations.Text1.Text = "Relation " & List2.ListCount + 1
    
    Relations.Show
    Me.Enabled = False

End Sub

Private Sub mnuProjectTableFields_Click()

    TableFields.Icon = Me.Icon
    
    ' create list with fields
    For i = 0 To FieldName.Count - 1
        If FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
            ' add field
            TableFields.List1.AddItem FieldName(i).Caption & " -> " & FieldType(i).Caption
        End If
    Next i

    TableFields.Show
    Me.Enabled = False

End Sub

Private Sub MoveTable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CurrentTableLeft = X
    CurrentTableTop = Y
    MoveObject = "Bottom"

End Sub

Private Sub MoveTable2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CurrentTableLeft = X
    CurrentTableTop = Y
    MoveObject = "Top"

End Sub

Private Sub OrgColor_Click()

    SetColor.BackColor = OrgColor.BackColor

End Sub

Private Sub Relation_Caption_Click(Index As Integer)

    List2.Text = Relation_Caption(Index).Caption
    mnuEditDelete.Caption = "Delete relation"
    mnuEditDelete.Enabled = True

End Sub

Private Sub Relation_Caption_DblClick(Index As Integer)

    OldRelationTitle = List2.Text
    RelationTitle = InputBox("Enter the new name for this relation:", "Relation name", Relation_Caption(Index).Caption)
    If RelationTitle <> "" Then
        List2.Text = RelationTitle
        If List2.Text = RelationTitle Then
            MsgBox "Relation already exists!", vbCritical + vbOKOnly, "Relation name"
        Else
            Relation_Caption(Index).Caption = RelationTitle
            Relation_1(Index).Tag = RelationTitle
            List2.Text = OldRelationTitle
            i = List2.ListIndex
            List2.RemoveItem i
            List2.AddItem RelationTitle, i
            List2.Text = RelationTitle
        End If
    End If

End Sub

Private Sub SetApply_Click()

    If ObjectName = "Table" Then
    
        Table(ObjectIndex).Width = SetWidth.Text
        Table(ObjectIndex).Height = Setheight.Text
        TableShadow(ObjectIndex).Width = SetWidth.Text
        TableShadow(ObjectIndex).Height = Setheight.Text
        Table(ObjectIndex).FillColor = SetColor.BackColor
        TableName(ObjectIndex).Width = SetWidth.Text
        TableSelector(ObjectIndex).Width = SetWidth.Text
        TableSelector(ObjectIndex).Height = Setheight.Text
        Shape.Width = SetWidth.Text + 0.12
        Shape.Height = Setheight.Text + 0.12
    
        'resize FieldType
        For i = 0 To FieldType.Count - 1
            If FieldType(i).ToolTipText = "TABLE." & ObjectIndex Then
                ' resize
                FieldType(i).Left = ((Table(ObjectIndex).Left + Table(ObjectIndex).Width) - FieldType(i).Width) - 0.05
            End If
        Next i
        
        'resize FieldName
        For i = 0 To FieldName.Count - 1
            If FieldName(i).ToolTipText = "TABLE." & ObjectIndex Then
                ' resize
                FieldName(i).Width = (Table(ObjectIndex).Width - FieldType(i).Width) - 0.15
            End If
        Next i
    
    End If
    
    MoveMovers

End Sub

Private Sub SetColor_Click()

    On Error GoTo Err:
    With CommonDialog1
        .CancelError = True
        .Color = SetColor.BackColor
        .Flags = 1
        .ShowColor
        SetColor.BackColor = .Color
    End With
Err:
    Exit Sub

End Sub

Private Sub Setheight_gotfocus()
    SetApply.Default = True
End Sub

Private Sub SetWidth_gotfocus()
    SetApply.Default = True
End Sub

Private Sub TableName_Click(Index As Integer)

    TableNameTitle = InputBox("Enter the new tablename:", "Tablename", TableName(Index).Caption)
    If TableNameTitle <> "" Then
        TableName(Index).Caption = TableNameTitle
    End If

End Sub

Private Sub TableSelector_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    With Shape
        .Left = Table(Index).Left - 0.05
        .Top = Table(Index).Top - 0.05
        .Width = Table(Index).Width + 0.15
        .Height = Table(Index).Height + 0.15
        .Visible = True
    End With
    
    SetWidth = Round(Table(Index).Width, 2)
    Setheight = Round(Table(Index).Height, 2)
    
    ObjectName = Table(Index).Name
    ObjectIndex = Table(Index).Index
    StatusBar1.Panels(1).Text = ObjectName & " " & ObjectIndex & " (" & TableName(ObjectIndex).Caption & ")"
    
    Shape.ZOrder (0)
    
    OrgColor.BackColor = Table(Index).FillColor
    List1.Text = "TABLE." & Index

    SetApply.Enabled = True
    mnuEditDeselect.Enabled = True
    mnuProjectTableFields.Enabled = True
    mnuEditFront.Enabled = True
    mnuEditBack.Enabled = True

    mnuEditDelete.Caption = "Delete table"
    mnuEditDelete.Enabled = True

    MoveMovers
    
    DeselectFields

End Sub

Private Sub TableSelector_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ObjectName <> "" Then
        If Button = 2 Then
            PopupMenu mnuProject
        End If
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Properties"
            
            Me.Enabled = False
            
            Properties.Text1.Text = Round(EditField.Width, 2) * 10
            Properties.Text2.Text = Round(EditField.Height, 2) * 10
            Properties.Picture1.BackColor = EditField.BackColor
            Properties.Picture2.BackColor = SetColor.BackColor
            Properties.Picture3.BackColor = TableName(TableName.ubound).BackColor
            
            Properties.Show

        Case "Save"
            SaveFile (CommonDialog1.FileName)
        Case "New"
            NewFile
        Case "Open"
            OpenFile
    End Select
End Sub

Private Sub Form_Load()
    
    Me.Caption = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision

    App.HelpFile = App.Path & "\Help.chm"
    
    ' set cursors
    TableSelector(0).MouseIcon = Cursors.ListImages(2).Picture
    TableName(0).MouseIcon = Cursors.ListImages(1).Picture
    Relation_Caption(0).MouseIcon = Cursors.ListImages(1).Picture
    FieldName(0).MouseIcon = Cursors.ListImages(1).Picture
    FieldType(0).MouseIcon = Cursors.ListImages(1).Picture
    
    ' set form sizes
    Editor.Width = Screen.Width / 1.4
    Editor.Height = Screen.Height / 1.4

    ' center form
    Editor.Left = (Screen.Width - Me.ScaleWidth) / 2
    Editor.Top = (Screen.Height - Me.ScaleHeight) / 2

    StatusBar1.Panels(2).Text = "Size: " & Round(EditField.Width, 2) & " x " & Round(EditField.Height, 2)

End Sub

Private Sub VScroll1_Change()

    EditField.Top = 0.2 - VScroll1.Value
    ResizeElements

End Sub

Private Sub VScroll1_Scroll()

    EditField.Top = 0.2 - VScroll1.Value
    ResizeElements

End Sub

Private Sub WorkArea_Click()

    DeSelect

End Sub

Private Sub WorkArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mnuFile
    End If

End Sub
