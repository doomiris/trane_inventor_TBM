VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   8220
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "&Back"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox hv 
      Height          =   5295
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9340
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmHelp.frx":06EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView hl 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   9340
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      ForeColor       =   4210688
      VariousPropertyBits=   8388627
      Caption         =   "Help and About"
      Size            =   "7646;661"
      FontName        =   "Arial Black"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LibPath As String

Private Sub Form_Load()
LibPath = LS("ServerPath", "\\shahns07\NPI_Program\Chiller_Projects\Global_Scroll_AC\Member_backup\Dominic\TBM") & "\Doc"
Call HelpListINI
End Sub

Private Sub cmdBack_Click()
Me.Hide
Unload Me
End Sub

Private Sub HelpListINI()
On Error Resume Next
With hl.Nodes
.Add , 1, "TBM", "TBM v" & c_ver
.Add "TBM", 4, "a", "About TBM"
.Add "TBM", 4, "u", "Update History"
.Add "TBM", 4, "f", "Functions"
.Add "TBM", 4, "q", "Q & A"
.Add "TBM", 4, "e", "For Epinal"
End With
hl.Nodes("a").Selected = True
hv.FileName = LibPath & "\about.rtf"
End Sub


Private Sub Form_Resize()
On Error Resume Next
If Not Me.WindowState = 1 Then
hl.Height = Me.Height - 1700
hv.Height = Me.Height - 1700
hv.Width = Me.Width - hv.Left - 300
cmdBack.Top = Me.Height - 600 - cmdBack.Height
cmdBack.Left = Me.Width - 300 - cmdBack.Width
End If
End Sub

Private Sub hl_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo openerror
With hv
Select Case hl.SelectedItem.Key
Case "a"
    .FileName = LibPath & "\about.rtf"
Case "u"
    .FileName = LibPath & "\Readme.rtf"
Case "f"
    .FileName = LibPath & "\Function Introduction.rtf"
Case "e"
    .FileName = LibPath & "\Epinal Mod.rtf"
Case "q"
    .FileName = LibPath & "\Q&A.rtf"
End Select
End With
Exit Sub
openerror:
hv.Text = "Can't load file! check network."
Resume Next
End Sub
