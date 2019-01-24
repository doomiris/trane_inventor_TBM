VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCharms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TBM for Epinal"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkPLini 
      Caption         =   "Format Existed &Parts List"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   4815
   End
   Begin VB.CheckBox chkTrimZero 
      Caption         =   "&Dimensions don't Trailing Zero(NOTE: this's NOT for Salvagnini )"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   5415
   End
   Begin VB.CheckBox chkAddTol 
      Caption         =   "Add Charms &Tolerance Automatically"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton CommandButton4 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CommandButton CommandButton3 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   3120
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Part Descriptions"
      Height          =   2295
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   3975
      Begin MSForms.ComboBox tbDesp3 
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   1680
         Width           =   2535
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4471;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ComboBox tbDesp2 
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4471;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ComboBox tbDesp1 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   2535
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4471;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ComboBox tbNum 
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   2535
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4471;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Arial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label Label8 
         Height          =   255
         Left            =   115
         TabIndex        =   10
         Top             =   1800
         Width           =   1005
         Caption         =   "Description 3"
         Size            =   "1773;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label7 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1005
         Caption         =   "Description 2"
         Size            =   "1773;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label6 
         Height          =   255
         Left            =   115
         TabIndex        =   8
         Top             =   840
         Width           =   1005
         Caption         =   "Description 1"
         Size            =   "1773;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1000
         Caption         =   "Part Number"
         Size            =   "1764;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Create Set"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3615
      Begin VB.TextBox tbSim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox tbDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox tbAuthor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox tbVer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   " %%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   735
         Caption         =   "Similar To"
         Size            =   "1296;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   975
         Caption         =   "MX_Revision"
         Size            =   "1720;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   615
         Caption         =   "Date"
         Size            =   "1085;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   735
         Caption         =   "Author"
         Size            =   "1296;450"
         FontName        =   "Arial"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCharms.frx":06EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCharms.frx":0DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCharms.frx":14DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCharms.frx":1BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCharms.frx":22D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   794
      ButtonWidth     =   1588
      ButtonHeight    =   794
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exist"
            Description     =   "Exist"
            Object.ToolTipText     =   "Load exist properties from current file."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Description     =   "Clear"
            Object.ToolTipText     =   "Clear All Text Box"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Description     =   "Help"
            Object.ToolTipText     =   "Show help and about."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TBM"
            Description     =   "Switch to TBM mode"
            Object.ToolTipText     =   "Switch to TBM mode"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCharms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oDrawDoc As DrawingDocument
Dim oSheet As Sheet
Dim oTitleBlockDef As TitleBlockDefinition
Dim oTitleBlock As TitleBlock
Dim oPropSets As PropertySets
Dim oPropSet As PropertySet
Dim StyTemName As String
Const BIE = "BoxInEdit"

Private Sub Form_Load()
If TbmLan > 0 Then
Call LanINI
End If
Call LoadOption
Call LoadExist
Call LoadSave
Call LoadPartNumList
End Sub
Private Sub LoadOption()
Set oDrawDoc = oApp.ActiveDocument
Set oSheet = oDrawDoc.ActiveSheet
Set oTitleBlockDef = oDrawDoc.TitleBlockDefinitions.Item(1)
Set oPropSets = oDrawDoc.PropertySets
chkAddTol.Value = LS("EpinalAddTol", "1")
chkTrimZero.Value = LS("EpinalTrimZero", "1")
chkPLini.Value = LS("EpinalPLini", "0")
StyTemName = LS("EpinalStyleTemName", "Copie - DEFAULT-ISO")
If oDrawDoc.ActiveSheet.PartsLists.Count > 0 Then
    chkPLini.Value = 1
End If
End Sub
Private Function CharmsDate(n)
CharmsDate = Format(n, DateForTrane)
End Function
Private Sub LoadSave()
tbAuthor.Text = oApp.UserName
tbDesp3.Text = CheckBlank(tbDesp3.Text, LS("LastEpinalDescription", "CXAM"))
tbDate.Text = CheckBlank(tbDate.Text, CharmsDate(Now))
tbVer.Text = CheckBlank(tbVer.Text, "/")

Dim pTempNumber As String
If oSheet.DrawingViews.Count > 0 Then
    pTempNumber = TrimFileName(oSheet.DrawingViews.Item(1).ReferencedFile.DisplayName)
    If Len(pTempNumber) = 12 Then
        pTempNumber = Left(pTempNumber, 8)
    End If
Else:
    pTempNumber = TrimFileName(oDrawDoc.DisplayName)
End If

tbNum.Text = CheckBlank(tbNum.Text, pTempNumber)

End Sub
Private Sub SaveOption()
'Call WS("oName", tbAuthor.Text)
Call WS("LastEpinalDescription", tbDesp3.Text)
'Call WS("LastProNum", Left(tbNum.Text, 4))
Call WS("EpinalTrimZero", chkTrimZero.Value)
Call WS("EpinalAddTol", chkAddTol.Value)

Set oPropSet = Nothing
Set oPropSets = Nothing
Set oTitleBlock = Nothing
Set oTitleBlockDef = Nothing
Set oSheet = Nothing
Set oDrawDoc = Nothing

End Sub
Private Function CheckBlank(Target As String, Value As String)
If Len(Target) = 0 Then
    CheckBlank = Value
Else:
    CheckBlank = Target
End If
End Function
Private Sub LoadExist()
Set oTitleBlock = oSheet.TitleBlock
Dim oTextBoxes As TextBoxes
Set oTextBoxes = oSheet.TitleBlock.Definition.Sketch.TextBoxes
tbVer.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(25))
tbAuthor.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(26))
tbDate.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(27))
tbNum.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(28))
tbSim.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(31))
tbDesp1.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(39))
tbDesp2.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(40))
tbDesp3.Text = oSheet.TitleBlock.GetResultText(oTextBoxes.Item(41))
End Sub
Private Sub WriteProperty()
Dim sPromptStrings(1 To 8) As String
sPromptStrings(1) = "" & tbVer.Text
sPromptStrings(2) = "" & tbAuthor.Text
sPromptStrings(3) = "" & tbDate.Text
sPromptStrings(4) = "" & tbNum.Text
sPromptStrings(5) = "" & tbSim.Text
sPromptStrings(6) = "" & tbDesp1.Text
sPromptStrings(7) = "" & tbDesp2.Text
sPromptStrings(8) = "" & tbDesp3.Text
oSheet.TitleBlock.Delete
Set oTitleBlock = oSheet.AddTitleBlock(oTitleBlockDef, , sPromptStrings)

Set oPropSet = oPropSets.Item("Design Tracking Properties")
    Call iPro("Designer", tbAuthor.Text)
    Call iPro("Part Number", tbNum.Text)
Set oPropSet = Nothing

End Sub
Private Sub LoadPartNumList()
Dim oDocs As Documents
Set oDocs = oApp.Documents
Dim oOpenDoc As Document
For i = 1 To oDocs.Count
Set oOpenDoc = oDocs.Item(i)
If oOpenDoc.DocumentType = kAssemblyDocumentObject Or oOpenDoc.DocumentType = kPartDocumentObject Then
    If oOpenDoc.Views.Count = 1 And Len(oOpenDoc.DisplayName) > 8 Then
    tbNum.AddItem Left(oOpenDoc.DisplayName, 8)
    End If
End If
Next i

'If tbNum.Text = LS("LastProNum", "5720") Or Len(tbNum.Text) = 0 Then
'    If tbNum.ListCount > 0 Then
'        tbNum.ListIndex = 0
'    End If
'End If
End Sub
Private Sub ClearAllBox()
tbAuthor.Text = ""
tbVer.Text = ""
tbDate.Text = ""
tbSim.Text = ""
tbNum.Text = ""
tbDesp1.Text = ""
tbDesp2.Text = ""
tbDesp3.Text = ""
End Sub
Private Sub CommandButton3_Click()
CommandButton3.Enabled = False
Call WriteProperty
If chkAddTol.Value = 1 Then
    Call AddTol
End If

Call TrimZero       'handle within function

If chkPLini.Value = 1 Then
    Call PartsListFormatSet
End If
Call TbmLog("E")
Call SaveOption
Me.Hide
Unload Me
End Sub
Private Sub AddTol()
On Error Resume Next
Set oPropSet = oPropSets.Item("User Defined Properties")
    Call iPro("ANGLES", "30")
    Call iPro("FINISH", "3,2")
    Call iPro("HOLE-", "")
    Call iPro("HOLE+", "")
    Call iPro("X,", "1")
    Call iPro("X,X", "0,4")
    Call iPro("X,XX", "0,12")
End Sub
Private Sub TrimZero()
Dim oStyMgr As DrawingStylesManager
Set oStyMgr = oDrawDoc.StylesManager
        If chkTrimZero.Value = 1 Then
            oStyMgr.DimensionStyles.Item(StyTemName).TrailingZeroDisplay = False
        Else:
            oStyMgr.DimensionStyles.Item(StyTemName).TrailingZeroDisplay = True
        End If
End Sub
Private Sub iPro(pname As String, pvalue As String)
oPropSet.Item(pname).Value = pvalue
End Sub
Private Sub CommandButton4_Click()
Me.Hide
Unload Me
End Sub
Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Call LoadExist
Case 2
    Call ClearAllBox
Case 3
    frmHelp.Show 1, Me
Case 4
    Me.Hide
    frmtbm.Show 1
    Unload Me
End Select
End Sub

Private Sub tbAuthor_DblClick()
tbAuthor.Text = oApp.UserName
End Sub
Private Sub tbDate_DblClick()
tbDate.Text = CharmsDate(Now)
End Sub

Private Sub tbDesp1_DblClick(Cancel As MSForms.ReturnBoolean)
tbDesp1.Text = UCase(tbDesp1.Text)
End Sub

Private Sub tbDesp1_GotFocus()
tbDesp1.Tag = BIE
End Sub

Private Sub tbDesp1_LostFocus()
tbDesp1.Tag = ""
End Sub

Private Sub tbDesp1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tbDesp1.Tag = BIE Then
    Exit Sub
End If
tbDesp1.SetFocus
End Sub

Private Sub tbDesp2_Change()
tbDesp2.Text = UCase(tbDesp2.Text)
End Sub

Private Sub tbDesp2_GotFocus()
tbDesp2.Tag = BIE
End Sub

Private Sub tbDesp2_LostFocus()
tbDesp2.Tag = ""
End Sub

Private Sub tbDesp2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tbDesp2.Tag = BIE Then
    Exit Sub
End If
tbDesp2.SetFocus
End Sub

Private Sub tbDesp3_DblClick(Cancel As MSForms.ReturnBoolean)
tbDesp3.Text = LS("LastEpinalDescription", " ")
End Sub

Private Sub tbDesp3_GotFocus()
tbDesp3.Tag = BIE
End Sub

Private Sub tbDesp3_LostFocus()
tbDesp3.Tag = ""
End Sub

Private Sub tbDesp3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tbDesp3.Tag = BIE Then
    Exit Sub
End If
tbDesp3.SetFocus
End Sub

Private Sub tbNum_GotFocus()
tbNum.Tag = BIE
End Sub

Private Sub tbNum_LostFocus()
tbNum.Tag = ""
End Sub

Private Sub tbNum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tbNum.Tag = BIE Then
    Exit Sub
End If
tbNum.SetFocus
End Sub

Private Sub tbVer_DblClick()
tbVer.Text = "/"
End Sub
Private Sub PartsListFormatSet()

    If oDrawDoc.ActiveSheet.PartsLists.Count > 0 Then
        Dim oPartsList As PartsList
        Set oPartsList = oDrawDoc.ActiveSheet.PartsLists.Item(1)
            oPartsList.ShowTitle = False        '隐藏title
        With oPartsList.PartsListColumns
            .Item(1).Width = 1.5
            .Item(2).Width = 3.75           '给定适当列宽
            .Item(3).Width = 1.5
            .Item(4).Width = 10.5
            If .Count > 4 Then
            .Item(5).Remove
            End If
        End With
        Set oPartsList = Nothing
    End If
    
End Sub
Private Sub LanINI()
With tb.Buttons
    For i = 1 To 4
        .Item(i).Caption = LRS(500 + i)
        .Item(i).ToolTipText = LRS(510 + i)
    Next
End With
Frame3.Caption = LRS(505)
Frame1.Caption = LRS(506)
CommandButton3.Caption = LRS(507)
CommandButton4.Caption = LRS(508)

chkAddTol.Caption = LRS(521)
chkTrimZero.Caption = LRS(522)
chkPLini.Caption = LRS(523)

End Sub
