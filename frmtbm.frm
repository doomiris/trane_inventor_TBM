VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmtbm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TBM"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmtbm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   960
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":0C5D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frm4 
      Caption         =   "Apply With"
      Height          =   1455
      Left            =   6600
      TabIndex        =   4
      Top             =   3600
      Width           =   3375
      Begin VB.CheckBox chkplf 
         Caption         =   "Format Existed &Parts List"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         ToolTipText     =   "Set Column's Width and hide title. See Help For Detail."
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox chkxxxx 
         Caption         =   "&ooxxxx"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         ToolTipText     =   "xxxx"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkDimSet 
         Caption         =   "&Drawing Dimension Precision Initialize"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         ToolTipText     =   "Set Dimension Precision, See Help For Detail."
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame frm3 
      Caption         =   "Project Set"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   6600
      TabIndex        =   3
      Top             =   480
      Width           =   3375
      Begin VB.CommandButton cmdPPDes 
         Caption         =   "Description"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox cmdVerUp 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Increase Version Number"
         Top             =   1220
         Width           =   300
      End
      Begin VB.CheckBox cmdRevUp 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1340
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Increase Revision Number"
         Top             =   1220
         Width           =   300
      End
      Begin VB.TextBox pPdes 
         Appearance      =   0  'Flat
         Height          =   780
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   29
         ToolTipText     =   "Double click here to re-generate description"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox pNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         TabIndex        =   23
         ToolTipText     =   "Double Click here to generate File name."
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox pUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         TabIndex        =   21
         ToolTipText     =   "Double click here to generate User Name"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Pdate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         TabIndex        =   27
         ToolTipText     =   "Double click here to generate current date"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox prevision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         TabIndex        =   25
         ToolTipText     =   "Double Click here to remove version numbers"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Create Date"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   240
         TabIndex        =   41
         Top             =   1740
         Width           =   915
      End
      Begin VB.Label lblver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MX_Revision"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   40
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Designer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1020
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":0FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":16E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":1DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":24D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":2BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":32CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtbm.frx":39C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   794
      ButtonWidth     =   1588
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exist"
            Key             =   "E"
            Description     =   "Exist"
            Object.ToolTipText     =   "Read exist properties from current file."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "C"
            Description     =   "Clear"
            Object.ToolTipText     =   "Clear all input boxes."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Spell"
            Key             =   "S"
            Description     =   "Check Spell"
            Object.ToolTipText     =   "Check Spell"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load"
            Key             =   "L"
            Description     =   "Load"
            Object.ToolTipText     =   "Load properties from other file."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Conf"
            Key             =   "o"
            Description     =   "Config"
            Object.ToolTipText     =   "Config TBM options."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Data"
            Description     =   "DataEdit"
            Object.ToolTipText     =   "Edit Templates Database"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "H"
            Description     =   "Help"
            Object.ToolTipText     =   "Show help and about."
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4440
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Select the name of a tolerance set"
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7832
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8760
      TabIndex        =   37
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   7440
      TabIndex        =   35
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame frm2 
      Caption         =   "Part Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Width           =   3975
      Begin MSForms.ComboBox pxdes2 
         Height          =   330
         Left            =   1365
         TabIndex        =   19
         ToolTipText     =   "Double Click Here Will Convert Texts to Capital Letters."
         Top             =   1680
         Width           =   2475
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4366;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ComboBox pxdes1 
         Height          =   330
         Left            =   1365
         TabIndex        =   17
         ToolTipText     =   "Double Click Here Will Convert Texts to Capital Letters."
         Top             =   1200
         Width           =   2475
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4366;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ComboBox pdes 
         Height          =   330
         Left            =   1365
         TabIndex        =   15
         ToolTipText     =   "Double Click Here Will Convert Texts to Capital Letters."
         Top             =   720
         Width           =   2475
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4366;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ComboBox pnoun 
         Height          =   330
         Left            =   1365
         TabIndex        =   13
         ToolTipText     =   "Double Click Here Will Convert Texts to Capital Letters."
         Top             =   240
         Width           =   2475
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4366;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label lblpxdes2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXT_DESP_2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblpxdes1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXT_DESP_1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblpdes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPTION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   0
         TabIndex        =   30
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lblpnoun 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PART_NOUN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame frm1 
      Caption         =   "Tolerance Set"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "When Lock in ""Pretect"" mode, the Tolerance Set cound only be changed by Selecting an Template"
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox px 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   780
         TabIndex        =   5
         Top             =   840
         Width           =   800
      End
      Begin VB.TextBox pxxx 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   780
         TabIndex        =   7
         Top             =   1800
         Width           =   800
      End
      Begin VB.TextBox pxx 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   780
         TabIndex        =   6
         Top             =   1320
         Width           =   800
      End
      Begin VB.TextBox pholel 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2820
         TabIndex        =   11
         Top             =   1800
         Width           =   800
      End
      Begin VB.TextBox pholeh 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2820
         TabIndex        =   10
         Top             =   1320
         Width           =   800
      End
      Begin VB.TextBox pfinish 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2820
         TabIndex        =   9
         Top             =   840
         Width           =   800
      End
      Begin VB.TextBox pangles 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2820
         TabIndex        =   8
         Top             =   360
         Width           =   800
      End
      Begin MSForms.ToggleButton tblock 
         Height          =   360
         Left            =   360
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Mistakes Prevent Lock"
         Top             =   360
         Width           =   360
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   6
         Size            =   "635;635"
         Value           =   "0"
         PicturePosition =   262148
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X,XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X,X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HOLE-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         TabIndex        =   20
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HOLE+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         TabIndex        =   18
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FINISH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ANGLES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   795
      End
   End
   Begin MSForms.Label sb 
      Height          =   255
      Left            =   240
      TabIndex        =   44
      ToolTipText     =   "See Help for Update Detail!"
      Top             =   5280
      Width           =   6975
      ForeColor       =   255
      VariousPropertyBits=   8388633
      Caption         =   "Title Block Manager"
      PicturePosition =   327683
      Size            =   "12303;450"
      MousePointer    =   2
      FontName        =   "Arial"
      FontEffects     =   1073750020
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Menu mnm 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu DesMoveDown 
         Caption         =   "Move Down"
      End
   End
End
Attribute VB_Name = "frmtbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim oDocs As Documents
Dim oDoc As Document
Dim oPropSets As PropertySets
Dim oPropSet As PropertySet

Public pip As Integer   'Part Desp Protect
Public alup As Integer  'alow update check
Public atm As Integer   'auto trim vision
Public lldp As Integer  'load last desp
Dim IsNewDoc As Boolean
Dim FirstLoad As Boolean
Dim MoveDate As String

Const BIE = "BoxInEdit"

Private Sub Form_Activate()
If FirstLoad = True Then
    If oApp.Documents.Count > 0 Then
        Call AppIni
        Call LoadCurrentProperties              '读取当前文档Properties
    End If
    FirstLoad = False
End If
'If Not LS("ServerVer", c_ver) = c_ver Then
'    If Not LS("VersionAllowed", "unknown") = c_ver Then
'        MsgBox "Current Version " & c_ver & " is Disabled for policy control, Please Update to the Latest Version!", vbCritical, "Update Your TBM"
'    Me.Hide
'    Unload Me
'    Exit Sub
'    'Else:
'    '    Call VersionDisabledCheckk
'    End If
'End If
pnoun.Tag = ""
pdes.Tag = ""
pxdes1.Tag = ""
pxdes2.Tag = ""
End Sub

Private Sub Form_Load()                 '窗体初始化过程

Call LoadOpition                        '读取config信息
Call ComboBoxListInitialize             '初始化Part Items下拉式列表
Call TreeViewInitialize                 '初始化模版树形列表

Call CheckUpdate                        '检查程序更新
FirstLoad = True
End Sub
Private Sub AppIni()

   Set oDoc = oApp.ActiveDocument
    Set oPropSets = oDoc.PropertySets
        IsNewDoc = (oDoc.FullFileName = "")         '判断是否为新创建的文件
    
If oDoc.DocumentType = kDrawingDocumentObject And IsNewDoc Then
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = oDoc
    chkDimSet.Value = 1
        If oDrawDoc.ActiveSheet.PartsLists.Count > 0 Then
            chkplf.Value = 1
        End If
End If
If TbmLan > 0 Then
    Call LocINI                                 '读取显示语言
End If
End Sub

Private Sub LoadOpition()
On Error GoTo loaderror

If oApp.Documents.Count > 0 Then
 Call AppIni
End If
On Error GoTo 0
    Me.Caption = "TBM v" & c_ver                  '设置程序title信息
    pip = LS("pip", 1)
    alup = LS("alup", 1)
    atm = LS("atm", 1)

    lldp = LS("lldp", 1)
    dbpath = LS("DBpath", "c:\Program Files\Autodesk\Inventor 9\addins\TBM\tbmdb.xls")
    tblock.Value = LS("TolLock", True)

    chkDimSet.Value = LS("DimSet", 0)
    chkplf.Value = LS("plf", 0)
    cmdPPDes.Tag = LS("AllowAutoGen", 1)
    pPdes.Tag = LS("GenMethod", 2)

Exit Sub
loaderror:
 MsgBox "You should open one document at least", vbCritical, "Error"
Unload Me
End Sub
Private Function TraneDate(n)
TraneDate = UCase(Format(n, DateForTrane))
End Function
Private Sub ComboBoxListInitialize()
Call dbOpen                             '建立数据库连接
Dim mySQL As String
With rs
mySQL = "select pnoun from [PnounList$];"
.Open mySQL, conn, 1, 3
For i = 1 To .RecordCount                   'Part_noun List
pnoun.AddItem (CStr(rs("pnoun")))
.MoveNext
Next i
.Close
mySQL = "select pdes from [PdesList$];"     'Description List
.Open mySQL, conn, 1, 3
For i = 1 To .RecordCount
pdes.AddItem (CStr(rs("pdes")))
.MoveNext
Next i
.Close
mySQL = "select pxdes1 from [Pxdes1List$];"    'ext1 Description List
.Open mySQL, conn, 1, 3
For i = 1 To .RecordCount
pxdes1.AddItem (CStr(rs("pxdes1")))
.MoveNext
Next i
.Close
mySQL = "select pxdes2 from [Pxdes2List$];"     'ext2 Description List
.Open mySQL, conn, 1, 3
For i = 1 To .RecordCount
pxdes2.AddItem (CStr(rs("pxdes2")))
.MoveNext
Next i
Call dbClose                            '关闭数据库连接
End With
End Sub

Private Sub TreeViewInitialize()
Call dbOpen                             '建立数据库连接
Dim mySQL As String
With tv
mySQL = "select name from [templates$];"
rs.Open mySQL, conn, 1, 3
.Nodes.Add , , "tlist", "Templates List"
For i = 1 To rs.RecordCount             'TreeView List INI
    Dim rsadd As String
    rsadd = "" & rs("name")
    .Nodes.Add "tlist", 4, rsadd, rsadd
    rs.MoveNext
Next i
.Nodes("tlist").Expanded = True
Call dbClose                            '关闭数据库连接
End With
End Sub
Private Sub CmdExit_Click()
Me.Hide
Unload Me
End Sub
Private Sub CmdApply_Click()
cmdApply.Enabled = False
Call WriteProperties
Call TbmLog("G")
If chkDimSet.Value = 1 Then
    Call DimensionPrecisionSet          '修改尺寸精度
End If
If chkplf.Value = 1 Then
    Call PartsListFormatSet              'PartList初始化
End If
Call WS("LastDescription", pxdes2.Text) 'save ext_desp_2
Call ClearMem                       '释放变量以清空占用内存
Me.Hide
Unload Me
End Sub
Private Function oPro(pname As String)
oPro = oPropSet.Item(pname).Value
End Function
Private Sub iPro(pname As String, pvalue As String)
oPropSet.Item(pname).Value = pvalue
End Sub
Private Sub LoadCurrentProperties()
On Error Resume Next
'tolerance set
Set oPropSet = oPropSets.Item("User Defined Properties")
    pangles.Text = oPro("ANGLES")
    pfinish.Text = oPro("FINISH")
    pholel.Text = oPro("HOLE-")
    pholeh.Text = oPro("HOLE+")
If oDoc.UnitsOfMeasure.LengthUnits = kInchLengthUnits Then 'distinguish Inch templates and meter templates
    px.Text = oPro(".X")
    pxx.Text = oPro(".XX")
    pxxx.Text = oPro(".XXX")
    Label7.Caption = ".X"
    Label8.Caption = ".XX"
    Label9.Caption = ".XXX"
Else:
    px.Text = oPro("X,")
    pxx.Text = oPro("X,X")
    pxxx.Text = oPro("X,XX")
    Label7.Caption = "X,"
    Label8.Caption = "X,X"
    Label9.Caption = "X,XX"
End If
'part description
    pnoun.Text = oPro("PART_NOUN")
    pdes.Text = oPro("DESCRIPTION")
    pxdes1.Text = oPro("EXT_DESCP_1")
    pxdes2.Text = oPro("EXT_DESCP_2")
'project set
    prevision.Text = oPro("MX_REVISION")
    Pdate.Text = oPro("DATE")
Set oPropSet = Nothing
Set oPropSet = oPropSets.Item("Design Tracking Properties")
    pNumber.Text = oPro("Part Number")
    pPdes.Text = oPro("Description")
    pUser = oPro("Designer")
Set oPropSet = Nothing
If IsNewDoc = True Then
    Call DefaultInput
End If

End Sub
Private Sub DefaultInput()
If Len(prevision.Text) = 0 Then
    prevision.Text = "01.0"
End If
If atm = 1 Then
    prevision.Text = TrimFileName(prevision.Text)
End If
If pnoun.Text = "Unassigned" Then
    pnoun.Text = ""
End If
If Pdate.Text = "DD-MMM-YYYY" Or Len(Pdate.Text) = 0 Then
    Pdate.Text = TraneDate(Now)
End If
If oDoc.DocumentType = kDrawingDocumentObject And Len(pNumber.Text) = 0 Then
    Dim pTempNumber As String
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = oDoc
    Dim oSheet As Sheet
    Set oSheet = oDoc.ActiveSheet
        If oSheet.DrawingViews.Count > 0 Then
            pTempNumber = TrimFileName(oSheet.DrawingViews.Item(1).ReferencedFile.DisplayName)
            If Len(pTempNumber) = 12 Then
                pTempNumber = Left(pTempNumber, 8)
            End If
        Else:
            pTempNumber = TrimFileName(oDoc.DisplayName)
        End If
        pNumber.Text = pTempNumber
End If
If lldp = 1 Then
    pxdes2.Text = LS("LastDescription", " ")
End If
End Sub


Private Sub WriteProperties()
On Error Resume Next
'tolerance set
Set oPropSet = oPropSets.Item("User Defined Properties")
    Call iPro("ANGLES", pangles.Text)
    Call iPro("FINISH", pfinish.Text)
    Call iPro("HOLE-", pholel.Text)
    Call iPro("HOLE+", pholeh.Text)
If oDoc.UnitsOfMeasure.LengthUnits = kInchLengthUnits Then 'distinguish Inch templates and meter templates
    Call iPro(".X", px.Text)
    Call iPro(".XX", pxx.Text)
    Call iPro(".XXX", pxxx.Text)
Else:
    Call iPro("X,", px.Text)
    Call iPro("X,X", pxx.Text)
    Call iPro("X,XX", pxxx.Text)
End If
'part description
    Call iPro("PART_NOUN", pnoun.Text)
    Call iPro("DESCRIPTION", pdes.Text)
    If oDoc.DocumentType = kPartDocumentObject Or oDoc.DocumentType = kAssemblyDocumentObject Then
        If LS("DesForPLM", 0) = 1 Then
            Dim TempSTR As String
            TempSTR = pnoun.Text
            If Not Len(pdes.Text) = 0 Then
                TempSTR = TempSTR & "; " & pdes.Text
            End If
            If LS("GenMethod", 2) = 3 Then
                If Not Len(pxdes1.Text) = 0 Then
                    TempSTR = TempSTR & "; " & pxdes1.Text
                End If
            End If
            Call iPro("DESCRIPTION", TempSTR)
        End If
    End If
    Call iPro("EXT_DESCP_1", pxdes1.Text)
    Call iPro("EXT_DESCP_2", pxdes2.Text)
'project set
    Call iPro("MX_REVISION", prevision.Text)
    Call iPro("DATE", Pdate.Text)
Set oPropSet = Nothing
Set oPropSet = oPropSets.Item("Design Tracking Properties")
    Call iPro("Designer", pUser.Text)
    Call iPro("Part Number", pNumber.Text)
    Call iPro("Description", pPdes.Text)
    Call iPro("Creation Time", Pdate.Text)
Set oPropSet = Nothing
Set oPropSet = oPropSets.Item("Summary Information")
    Call iPro("Revision Number", prevision.Text)
Set oPropSet = Nothing
End Sub

Private Sub DesMoveDown_click()
Dim tDes As String
tDes = pdes.Text

Dim txdes1 As String
txdes1 = pxdes1.Text

Dim txdes2 As String
txdes2 = pxdes2.Text

Select Case MoveDate
Case "PdesDown"
    If Len(txdes2) = 0 Then
        txdes2 = txdes1
    End If
    txdes1 = tDes
    tDes = ""
Case "Pxdes1Down"
    txdes2 = txdes1
    txdes1 = ""
End Select

pdes.Text = tDes
pxdes1.Text = txdes1
pxdes2.Text = txdes2
End Sub

Private Sub pdate_DblClick()
Pdate.Text = TraneDate(Now)
End Sub

Private Sub pdes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Len(pdes.Text) = 0 Then
    Exit Sub
    End If
    Dim oMenu As Menu
    Set oMenu = Me.mnm
    MoveDate = "PdesDown"
    PopupMenu oMenu
End If
End Sub

Private Sub pxdes1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Len(pxdes1.Text) = 0 Then
    Exit Sub
    End If
    Dim oMenu As Menu
    Set oMenu = Me.mnm
    MoveDate = "Pxdes1Down"
    PopupMenu oMenu
End If
End Sub
Private Sub pnoun_Change()
Call ProDesGen
End Sub
Private Sub pdes_Change()
Call ProDesGen
End Sub
Private Sub pNumber_DblClick()
Dim tempN As String
Dim IsDrawing As Boolean

tempN = oDoc.DisplayName

If oDoc.DocumentType = kDrawingDocumentObject Then
    Dim oDrwDoc As DrawingDocument
    Set oDrwDoc = oDoc
    
    IsDrawing = True
    
    Dim oSheet As Sheet
    Set oSheet = oDrwDoc.ActiveSheet
    
    If oSheet.DrawingViews.Count > 0 Then
        tempN = oSheet.DrawingViews.Item(1).ReferencedFile.DisplayName
        
    End If
Else:
    IsDrawing = False
    Dim tempM As String
    tempM = oDoc.FullFileName
    If Not Len(tempM) = 0 Then
        Dim DotLoc As Integer
        DotLoc = InStrRev(tempM, "\")
        tempN = Mid(tempM, DotLoc + 1, Len(tempM) - DotLoc)
    End If
End If

tempN = TrimFileName(tempN)

If IsDrawing = True And Len(tempN) = 12 Then
    tempN = Left(tempN, 8)
End If

pNumber.Text = tempN

End Sub

Private Sub pPdes_DblClick()
Call ProDesGen
End Sub

Private Sub pxdes1_Change()
Call ProDesGen
End Sub
Private Sub pnoun_DblClick(Cancel As MSForms.ReturnBoolean)
pnoun.Text = Format(pnoun.Text, ">")
End Sub
Private Sub pdes_DblClick(Cancel As MSForms.ReturnBoolean)
pdes.Text = Format(pdes.Text, ">")
End Sub
Private Sub prevision_Change()
prevision.Text = UCase(prevision.Text)
End Sub
Private Sub pUser_DblClick()
pUser.Text = oApp.UserName
End Sub
Private Sub pxdes1_DblClick(Cancel As MSForms.ReturnBoolean)
pxdes1.Text = Format(pxdes1.Text, ">")
End Sub

Private Sub pdes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pPdes.Tag = BIE Then
    Exit Sub
End If
pdes.SetFocus
End Sub
Private Sub pdes_GotFocus()
pPdes.Tag = BIE
End Sub
Private Sub pdes_LostFocus()
pPdes.Tag = ""
End Sub

Private Sub pnoun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pnoun.Tag = BIE Then
    Exit Sub
End If
pnoun.SetFocus
End Sub
Private Sub pnoun_GotFocus()
pnoun.Tag = BIE
End Sub
Private Sub pnoun_LostFocus()
pnoun.Tag = ""
End Sub

Private Sub pxdes1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pxdes1.Tag = BIE Then
    Exit Sub
End If
pxdes1.SetFocus
End Sub
Private Sub pxdes1_GotFocus()
pxdes1.Tag = BIE
End Sub
Private Sub pxdes1_LostFocus()
pxdes1.Tag = ""
End Sub

Private Sub pxdes2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pxdes2.Tag = BIE Then
Exit Sub
End If
pxdes2.SetFocus
End Sub
Private Sub pxdes2_GotFocus()
pxdes2.Tag = BIE
End Sub
Private Sub pxdes2_LostFocus()
pxdes2.Tag = ""
End Sub

Private Sub pxdes2_DblClick(Cancel As MSForms.ReturnBoolean)
pxdes2.Text = Format(pxdes2.Text, ">")
End Sub
Private Sub prevision_DblClick()
prevision.Text = TrimFileName(prevision.Text)
End Sub
Private Sub ScrollBar1_GotFocus()
prevision.SetFocus
End Sub

Private Sub sb_Click()
    Call ShellExecute(Me.hwnd, "Open", "\\shahns07\NPI_Program\Chiller_Projects\Global_Scroll_AC\Member_backup\Dominic\TBM", "", "", 1)
End Sub

Private Sub tblock_Click()
If tblock.Value = False Then
    tblock.Picture = ImageList2.ListImages(1).Picture
    Call TolLock(False)
    If TbmLan > 0 Then
        Text4.Text = LRS(164)
    Else:
        Text4.Text = "unProtect"
    End If
Else:
    tblock.Picture = ImageList2.ListImages(2).Picture
    Call TolLock(True)
    If TbmLan > 0 Then
        Text4.Text = LRS(163)
    Else:
        Text4.Text = "Protect"
    End If
End If
End Sub
Private Sub TolLock(Mode As Boolean)
px.Locked = Mode
pxx.Locked = Mode
pxxx.Locked = Mode
pangles.Locked = Mode
pfinish.Locked = Mode
pholeh.Locked = Mode
pholel.Locked = Mode
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
If Not Len(tv.SelectedItem) = 0 And Not tv.SelectedItem = "Templates List" Then
    Call dbOpen
    Dim mySQL As String
    mySQL = "select * from [templates$] where name='" & tv.SelectedItem & "';"
    On Error GoTo xlserror
    rs.Open mySQL, conn, 1, 3
        pangles.Text = checknull(rs("pangles"))
        pfinish.Text = checknull(rs("pfinish"))
        pholel.Text = checknull(rs("pholel"))
        pholeh.Text = checknull(rs("pholeh"))
        px.Text = checknull(rs("px"))
        pxx.Text = checknull(rs("pxx"))
        pxxx.Text = checknull(rs("pxxx"))
        If pip = 0 Then            'Part Items Protect
            pnoun.Text = checknull(rs("pnoun"))
            pdes.Text = checknull(rs("pdes"))
            pxdes1.Text = checknull(rs("pxdes1"))
            pxdes2.Text = checknull(rs("pxdes2"))
        End If
    Call dbClose
End If
Exit Sub
xlserror:
MsgBox "Template:" & rs("name") & "has wrong data format, refer to Readme sheet and correct it!", vbCritical, "Error"
End Sub


Private Sub DimensionPrecisionSet() '初始化图纸尺寸精度
If oDoc.DocumentType = kDrawingDocumentObject Then
    On Error Resume Next
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = oApp.ActiveDocument
    
    Dim oStyleMgr As DrawingStylesManager
    Set oStyleMgr = oDrawDoc.StylesManager
    With oStyleMgr.ObjectDefaultsStyles.Item(1)
        .LinearDimensionStyle.LinearPrecision = kOneDecimalPlaceLinearPrecision
        '线性尺寸保留1位小数
        .AngularDimensionStyle.AngularPrecision = kZeroDecimalPlaceAngularPrecision
        '角度尺寸保留0位小数
        .OrdinateSetDimensionStyle.LinearPrecision = kOneDecimalPlaceLinearPrecision
        '座标尺寸保留1位小数
    End With
End If
End Sub
Private Sub ProDesGen()
On Error Resume Next
If cmdPPDes.Tag = 1 Then
    Dim ProDes As String
    ProDes = pnoun.Text                         'method 1
    If pPdes.Tag > 1 Then
        If Not Len(pdes.Text) = 0 Then          'method 2
            ProDes = ProDes & "; " & pdes.Text
        End If
        If pPdes.Tag > 2 Then                     'method 3
            If Not Len(pxdes1) = 0 Then
                ProDes = ProDes & "; " & pxdes1.Text
            End If
        End If
    End If
    pPdes.Text = ProDes
End If
End Sub
Private Sub ClearMem()
On Error Resume Next
    Set oPropSet = Nothing
    Set oPropSets = Nothing
    Set oDoc = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call ClearMem                           '释放变量以清空占用内存
End Sub
Private Sub ClearAllInputBox()
pangles.Text = ""
pfinish.Text = ""
pholel.Text = ""
pholeh.Text = ""
px.Text = ""
pxx.Text = ""
pxxx.Text = ""
pnoun.Text = ""
pdes.Text = ""
pxdes1.Text = ""
pxdes2.Text = ""
Pdate.Text = ""
prevision.Text = ""
pPdes.Text = ""
pNumber.Text = ""
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index         'toolbar button events
Case 1
    Call LoadCurrentProperties
Case 2
    Call ClearAllInputBox
Case 3
        frmSpellCheck.Show 1, Me
Case 4
    frmLoad.Show 1, Me
Case 5
    tmConfig.Show 1, Me
Case 6
    On Error Resume Next
    Call ShellExecute(Me.hwnd, "Open", dbpath, "", "", 1)
    On Error GoTo 0
Case 7
    frmHelp.Show 1, Me
End Select
End Sub

Private Sub PartsListFormatSet()
If oDoc.DocumentType = kDrawingDocumentObject Then
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = oDoc
    If oDrawDoc.ActiveSheet.PartsLists.Count > 0 Then
        Dim oPartsList As PartsList
        Set oPartsList = oDrawDoc.ActiveSheet.PartsLists.Item(1)
            oPartsList.ShowTitle = False        '隐藏title
        With oPartsList.PartsListColumns
            .Item("ITEM").Width = 1.5           '给定适当列宽
            .Item("QTY").Width = 1.5
            .Item("PART NUMBER").Width = 4
            .Item("DESCRIPTION").Width = 10.5
        End With
        Set oPartsList = Nothing
    End If
    Set oDrawDoc = Nothing
End If
End Sub
'Private Sub SaveNewDes()
'Call dbOpen
'Dim mySQL As String
'Dim tempN As String
'With rs
'    If Not Len(pnoun.Text) = 0 Then     'check Part_Noun List
'        tempN = UCase(pnoun.Text)
'        mySQL = "select * from [PnounList$] where pnoun='" & tempN & "';"
'        .Open mySQL, conn, 1, 3
'        If rs.RecordCount = 0 Then
'            rs.AddNew
'            rs("pnoun") = tempN
'            rs.Update
'        End If
'        .Close
'        pnoun.AddItem (tempN)
'    End If
'    If Not Len(pdes.Text) = 0 Then     'check Desp List
'        tempN = UCase(pdes.Text)
'        mySQL = "select * from [PdesList$] where pdes='" & tempN & "';"
'        .Open mySQL, conn, 1, 3
'        If rs.RecordCount = 0 Then
'            rs.AddNew
'            rs("pdes") = tempN
'            rs.Update
'        End If
'        .Close
'        pdes.AddItem (tempN)
'    End If
'    If Not Len(pxdes1.Text) = 0 Then     'check ext_desp_1 List
'        tempN = UCase(pxdes1.Text)
'        mySQL = "select * from [Pxdes1List$] where pxdes1='" & tempN & "';"
'        .Open mySQL, conn, 1, 3
'        If rs.RecordCount = 0 Then
'            rs.AddNew
'            rs("pxdes1") = tempN
'            rs.Update
'        End If
'        .Close
'        pxdes1.AddItem (tempN)
'    End If
'    If Not Len(pxdes2.Text) = 0 Then     'check ext_desp_2 List
'        tempN = UCase(pxdes2.Text)
'        mySQL = "select * from [Pxdes2List$] where pxdes2='" & tempN & "';"
'        .Open mySQL, conn, 1, 3
'        If rs.RecordCount = 0 Then
'            rs.AddNew
'            rs("pxdes2") = tempN
'            rs.Update
'        End If
'        .Close
'        pxdes2.AddItem (tempN)
'    End If
'
'    Call dbClose
'End With
'Exit Sub
'
'End Sub

Private Sub cmdPPDes_Click()
frmPPG.Show 1, Me
End Sub

Private Sub cmdRevUp_Click()
If cmdRevUp.Value = 0 Then
    Dim cMXv As String
    Dim DotPos As Integer
    Dim Rev As String
    cMXv = prevision.Text
    DotPos = InStr(1, cMXv, ".")
    If DotPos > 0 Then
        Rev = Left(cMXv, DotPos - 1)
    Else:
        Rev = IIf(Len(cMXv) = 0, 0, cMXv)
    End If
    
    If IsNumeric(Rev) Then
        Rev = CInt(Rev) + 1
        Rev = IIf(Rev < 10, "0" & Rev, Rev)
    Else:
        If Rev Like "[A-Y]" Then
        Rev = Chr(Asc(UCase(Rev)) + 1)
        End If
    End If
    prevision.Text = Rev & ".0"
End If
    cmdRevUp.Value = 0
    prevision.SetFocus

End Sub

Private Sub cmdVerUp_Click()
If cmdVerUp.Value = 0 Then
    Dim cMXv As String
    Dim DotPos As Integer
    Dim Rev As String
    Dim VerN As String
    cMXv = prevision.Text
        If Len(cMXv) = 0 Then
            cMXv = "01"
        End If
    DotPos = InStr(1, cMXv, ".")
        If DotPos < 1 Then
            Rev = cMXv
            VerN = -1
        Else:
            Rev = Left(cMXv, DotPos - 1)
            VerN = Right(cMXv, Len(cMXv) - DotPos)
        End If
    VerN = IIf(Len(VerN) = 0, -1, VerN)
    cMXv = Rev & "." & VerN + 1
    prevision = cMXv
End If
cmdVerUp.Value = 0
prevision.SetFocus

End Sub
Private Sub LocINI()
On Error GoTo errhand
'工具按钮
With tb.Buttons
    Dim i As Integer
    For i = 1 To 7
        .Item(i).Caption = LRS(100 + i)
        .Item(i).ToolTipText = LRS(150 + i)
    Next
End With
'框体
frm1.Caption = LRS(111)
frm2.Caption = LRS(112)
frm3.Caption = LRS(113)
frm4.Caption = LRS(114)
'按钮和TV
cmdApply.Caption = LRS(115)
CmdExit.Caption = LRS(116)
tv.Nodes.Item(1).Text = LRS(117)
tv.ToolTipText = LRS(118)
'额外选项
chkDimSet.Caption = LRS(121)
chkplf.Caption = LRS(122)
'Project内容
Label10.Caption = LRS(131)
Label5.Caption = LRS(132)
lblver.Caption = LRS(133)
lbldate.Caption = LRS(134)
cmdPPDes.Caption = LRS(135)
cmdPPDes.ToolTipText = LRS(136)
'公差锁
tblock.ToolTipText = LRS(161)
Text4.ToolTipText = LRS(162)
'描述栏的tooltip
pnoun.ToolTipText = LRS(171)
pdes.ToolTipText = LRS(172)
pxdes1.ToolTipText = LRS(173)
pxdes2.ToolTipText = LRS(174)
'project项的tooltip
pUser.ToolTipText = LRS(181)
pNumber.ToolTipText = LRS(182)
prevision.ToolTipText = LRS(183)
cmdRevUp.ToolTipText = LRS(184)
cmdVerUp.ToolTipText = LRS(185)
Pdate.ToolTipText = LRS(186)
pPdes.ToolTipText = LRS(187)
'额外项的tooltip
chkDimSet.ToolTipText = LRS(191)
chkplf.ToolTipText = LRS(192)
Exit Sub
errhand:
'MsgBox Err.Description, , Err.Source
Resume Next
End Sub
