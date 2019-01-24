VERSION 5.00
Begin VB.Form frmPPG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Description Generate Option"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPPG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6000
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2535
      Begin VB.Label Label4 
         Caption         =   ">Generate EXT_Desp_2 (3)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   ">Generate EXT_Desp_1 (2)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   ">Generate DESCRIPTION (1)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   ">Generate Part_Noun"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generate Method"
      Height          =   1335
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   2895
      Begin VB.OptionButton chkGenPnoun12 
         Caption         =   "Part_Noun+(1)+(2)"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2535
      End
      Begin VB.OptionButton chkGenPnoun1 
         Caption         =   "Part_Noun+(1)"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton chkGenPnoun 
         Caption         =   "Part_Noun only"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Back"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2040
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   945
   End
   Begin VB.CheckBox chkAllowAutoGen 
      Caption         =   "Allow Auto Generate"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmPPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GenMethod
Private Sub chkAllowAutoGen_Click()
With chkAllowAutoGen
    Frame1.Enabled = .Value
    chkGenPnoun.Enabled = .Value
    chkGenPnoun1.Enabled = .Value
    chkGenPnoun12.Enabled = .Value
End With
End Sub

Private Sub chkGenPnoun_Click()
GenMethod = 1
End Sub

Private Sub chkGenPnoun1_Click()
GenMethod = 2
End Sub

Private Sub chkGenPnoun12_Click()
GenMethod = 3
End Sub

Private Sub Command1_Click()
Call WS("AllowAutoGen", chkAllowAutoGen.Value)
Call WS("GenMethod", GenMethod)
frmtbm.cmdPPDes.Tag = chkAllowAutoGen.Value
frmtbm.pPdes.Tag = GenMethod
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If TbmLan > 0 Then
Call LanINI
End If
chkAllowAutoGen.Value = LS("AllowAutoGen", 1)
Select Case frmtbm.pPdes.Tag
Case 1
    chkGenPnoun.Value = True
    GenMethod = 1
Case 2
    chkGenPnoun1.Value = True
    GenMethod = 2
Case 3
    chkGenPnoun12.Value = True
    GenMethod = 3
End Select
End Sub

Private Sub LanINI()
Me.Caption = LRS(350)
chkAllowAutoGen.Caption = LRS(351)
Frame1.Caption = LRS(352)
chkGenPnoun.Caption = LRS(353)
chkGenPnoun1.Caption = LRS(354)
chkGenPnoun12.Caption = LRS(355)
Label1.Caption = LRS(356)
Label2.Caption = LRS(357)
Label3.Caption = LRS(358)
Label4.Caption = LRS(359)
Command1.Caption = LRS(360)
Command2.Caption = LRS(361)
End Sub
