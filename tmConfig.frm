VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form tmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Config TBM"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "tmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Height          =   350
      Left            =   8500
      Picture         =   "tmConfig.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Open a file on disk"
      Top             =   489
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "EXCEL文件(*.xls)|*.xls"
   End
   Begin VB.TextBox txtdbpath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   8775
   End
   Begin VB.CommandButton CommandButton2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7920
      TabIndex        =   3
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   400
      Left            =   6840
      TabIndex        =   2
      Top             =   3360
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Default Value Set"
      Height          =   2175
      Left            =   4560
      TabIndex        =   1
      Top             =   960
      Width           =   4335
      Begin VB.CheckBox chkTolLock 
         Caption         =   "Tolerance Set Default Locked"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkDimSet 
         Caption         =   "&Drawing Dimension Initialize Default Checked"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox chkplf 
         Caption         =   "&Format PartsList if exist Default Checked"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TBM Options"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4335
      Begin VB.CheckBox chklldp 
         Caption         =   "Remenber the &Last ext2 Description"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Use &Relevant drop down List"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkatm 
         Caption         =   "Auto &Trim MX_Revision"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chkpip 
         Caption         =   "&Protect Part Descriptions when changing templates"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
      Begin VB.CheckBox chkalup 
         Caption         =   "Check for &Update automatically"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSForms.ComboBox cbLan 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2566;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbl2 
      Caption         =   "TBM Languages:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Full path of the EXCEL database: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "tmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cd.ShowOpen
Dim tempS As String
tempS = txtdbpath.Text
tempS = IIf(Len(cd.FileName) = 0, tempS, cd.FileName)
txtdbpath.Text = tempS
CommandButton1.SetFocus
End Sub

Private Sub Form_Load()
If TbmLan > 0 Then
    Call LanINI
End If
On Error GoTo errhand
cbLan.AddItem "English"
cbLan.AddItem "简体中文"
cbLan.ListIndex = TbmLan
On Error GoTo 0
With frmtbm     '从tbm读取变量值
chkalup.Value = .alup
chkpip.Value = .pip
chkatm.Value = .atm
chklldp.Value = .lldp

txtdbpath.Text = dbpath

chkTolLock.Value = IIf(.tblock.Value, 1, 0)
chkDimSet.Value = LS("DimSet", 0)
chkplf.Value = LS("plf", 0)


End With
Exit Sub
errhand:
MsgBox Err.Description, , "Form_Load"
Resume Next
End Sub
Private Sub CommandButton2_Click()
Me.Hide
Unload Me
End Sub
Private Sub CommandButton1_Click()
On Error GoTo errhand
With frmtbm
.alup = chkalup.Value   '写回tbm窗休
.pip = chkpip.Value
.atm = chkatm.Value
.lldp = chklldp.Value
Call WS("alup", .alup) '写入注册表
Call WS("pip", .pip)
Call WS("atm", .atm)
Call WS("lldp", .lldp)

Call WS("TolLock", IIf(chkTolLock.Value = 1, "True", "False"))
Call WS("DimSet", chkDimSet.Value)
Call WS("plf", chkplf.Value)

End With
If txtdbpath.Tag = "changed" Then
Call dbTest(txtdbpath.Text)
dbpath = txtdbpath.Text
End If

TbmLan = cbLan.ListIndex
Call WS("Languages", TbmLan)

Me.Hide
Unload frmtbm
Load frmtbm
frmtbm.Show
Unload Me

Exit Sub
errhand:
MsgBox Err.Description, , "Apply_Click"
Resume Next
End Sub



Private Sub txtdbpath_Change()
txtdbpath.Tag = "changed"
End Sub
Sub dbTest(DataPath As String)
'

End Sub
Private Sub LanINI()
On Error GoTo errhand
Me.Caption = LRS(401)
Frame1.Caption = LRS(402)
Frame2.Caption = LRS(403)
lbl2.Caption = LRS(404)
CommandButton1.Caption = LRS(405)
CommandButton2.Caption = LRS(406)
Label1.Caption = LRS(411)
chkalup.Caption = LRS(412)
chkpip.Caption = LRS(413)
chkatm.Caption = LRS(414)
'Check4.Caption = LRS(415) 'unuse
chklldp.Caption = LRS(416)
chkTolLock.Caption = LRS(431)
chkDimSet.Caption = LRS(432)
chkplf.Caption = LRS(433)
Exit Sub
errhand:
MsgBox Err.Description, , "LanINI"
Resume Next
End Sub
