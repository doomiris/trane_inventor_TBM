VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form tmConfig1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Config TBM"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "tmConfig2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4080
      TabIndex        =   3
      Top             =   3720
      Width           =   1000
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   400
      Left            =   3000
      TabIndex        =   2
      Top             =   3720
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Additional Functions"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   5055
      Begin MSForms.CheckBox chkapw 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   4575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "8070;450"
         Value           =   "0"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkdpset 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   4575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "8070;450"
         Value           =   "0"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TBM Options"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin MSForms.CheckBox chklldp 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   4575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "8070;450"
         Value           =   "0"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox CheckBox4 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   4575
         VariousPropertyBits=   746588185
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "8070;450"
         Value           =   "0"
         FontName        =   "Arial"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkatm 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "8070;450"
         Value           =   "0"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkpip 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   4575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "8070;450"
         Value           =   "0"
         Caption         =   "&Protect Part Items when changing templates"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.CheckBox chkalup 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   4575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "8070;450"
         Value           =   "0"
         Caption         =   "Check for &Update automatically"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
End
Attribute VB_Name = "tmConfig1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
With frmtbm     '从tm读取变量值
chkalup.Value = .alup
chkpip.Value = .pip
chkatm.Value = .atm
chkdpset.Value = .dpset
chkdpset.ToolTipText = "This function will set linear dimension precision to 1.1," & vbCrLf & " set angular dimension precision to 0," & vbCrLf & " set ordinate dimension precision to 1.1" & vbCrLf & " when you apply properties using TBM"
chkapw.Value = .apw
chkapw.ToolTipText = "This function will write ""Part_Noun; Description; EXT_DESCP_1"" , ""MX_Revision"" and "" DATA"" to ""Descripttion"" , ""Revision Number"" and  ""Creation Date"" of ""Project"" when you Apply."
chklldp.Value = .lldp
End With
End Sub
Private Sub CommandButton2_Click()
Unload Me
End Sub
Private Sub CommandButton1_Click()
With frmtbm
.alup = chkalup.Value   '写回tm窗休
.pip = chkpip.Value
.atm = chkatm.Value
.dpset = chkdpset.Value
.apw = chkapw.Value
.lldp = chklldp.Value

Call .WS("alup", .alup) '写入注册表
Call .WS("pip", .pip)
Call .WS("atm", .atm)
Call .WS("dpset", .dpset)
Call .WS("apw", .apw)
Call .WS("lldp", .lldp)
End With
Unload Me
End Sub

