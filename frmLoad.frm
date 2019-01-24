VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load From File"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkEpinal 
      Caption         =   "Load for Epinal"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CommandButton3 
      Cancel          =   -1  'True
      Caption         =   "&Back"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Load Option"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
      Begin VB.CheckBox chkLfp 
         Caption         =   "Load From Project Table When Nothing Exist"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   5295
      End
      Begin VB.CheckBox chkOnlyBlanks 
         Caption         =   "Only insert &blanks (whose text box is empty)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox chkLoadTol 
         Caption         =   "Auto Select a Tolerance Set"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   680
         Width           =   5295
      End
   End
   Begin VB.CommandButton CommandButton1 
      Height          =   360
      Left            =   5280
      Picture         =   "frmLoad.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open a file on disk"
      Top             =   720
      Width           =   360
   End
   Begin MSComDlg.CommonDialog cdbox 
      Left            =   240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Inventor文件(*.idw;*.ipt;*.iam)|*.iam;*.idw;*.ipt"
   End
   Begin MSForms.ComboBox cbfile 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   5055
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "8916;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "Select a file in the drop down list or open a new one to continue"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oDocs As Documents
Dim oDoc As Document
Dim oPropSets As PropertySets
Dim oPropSet As PropertySet
Dim OpenFlag As Boolean

Private Sub Form_Load()
If TbmLan > 0 Then
    Call LanINI
End If
Call FileLoadConfig
Call FileListINI
End Sub

Private Sub FileLoadConfig()

Set oDocs = oApp.Documents

chkLoadTol.Value = LS("LoadTol", 0)
chkOnlyBlanks.Value = LS("OnlyBlanks", 0)
chkLfp.Value = LS("LoadFromProject", 0)

End Sub
Private Sub FileListINI()

For Each oDoc In oDocs
If Not oDoc.DocumentType = kUnknownDocumentObject Then
        If oDoc.Views.Count = 1 Then                ' what's this.....file open?
            If Not Len(oDoc.FullFileName) = 0 Then
                cbfile.AddItem oDoc.FullFileName
            End If
        End If
End If
Next
If cbfile.ListCount > 0 Then
cbfile.ListIndex = 0
End If
End Sub
Private Sub CommandButton1_Click()
On Error Resume Next
Dim oFileLoc As FileLocations
Set oFileLoc = oApp.FileLocations
cdbox.InitDir = oFileLoc.Workspace
cdbox.Action = 1
cbfile.Text = cdbox.FileName
End Sub
Private Sub CommandButton2_Click()
CommandButton2.Enabled = False
If Not Trim(cbfile.Text) = "" Then
    Call LoadFromFile
    If OpenFlag = True Then
        Call LoadConfigSave
        Me.Hide
        Unload Me
    End If
Else:
        Call LoadConfigSave
        Me.Hide
        Unload Me
End If
End Sub

Private Sub CommandButton3_Click()
Call LoadConfigSave
Me.Hide
Unload Me
End Sub
Private Sub LoadConfigSave()

Call WS("LoadTol", chkLoadTol.Value)
Call WS("OnlyBlanks", chkOnlyBlanks.Value)
Call WS("LoadFromProject", chkLfp.Value)
End Sub

Private Sub LoadFromFile()
Set oDoc = Nothing
    On Error GoTo openerror
    Dim tDoc As Document
    Set tDoc = oApp.Documents.Open(cbfile.Text, False)
    On Error GoTo 0
Set oPropSets = tDoc.PropertySets
Set oPropSet = oPropSets.Item("User Defined Properties")

Dim Lpnoun As String, Lpdes As String, Lpxdes1 As String, Lpxdes2 As String
On Error Resume Next
Lpnoun = oPro("PART_NOUN")
Lpdes = oPro("DESCRIPTION")
Lpxdes1 = oPro("EXT_DESCP_1")
Lpxdes2 = oPro("EXT_DESCP_2")
On Error GoTo 0
If chkLfp.Value = 1 Then
     If Len(Lpnoun) = 0 Then
         Set oPropSet = oPropSets.Item("Design Tracking Properties")
         Dim ProjectDes As String
         ProjectDes = oPro("Description")
            If Not Len(ProjectDes) = 0 Then
                Dim DesStr
                DesStr = Split(ProjectDes, ";")
                For i = 0 To UBound(DesStr)
                    Select Case i
                        Case 0
                                Lpnoun = DesStr(0)
                        Case 1
                            If Not Len(DesStr(1)) = 0 Then
                                Lpdes = DesStr(1)
                            End If
                        Case 2
                            If Not Len(DesStr(2)) = 0 Then
                                Lpxdes1 = DesStr(2)
                            End If
                        Case 3
                            If Not Len(DesStr(3)) = 0 Then
                                Lpxdes2 = DesStr(3)
                            End If
                    End Select
                Next
        End If
    End If
End If
With frmtbm
On Error Resume Next

    .pnoun.Text = BlankCheck(.pnoun.Text, Lpnoun)
        
If InStr(1, Lpdes, ";", vbTextCompare) > 1 Then
    Dim TempSTR
    TempSTR = Split(Lpdes, ";")
        If UBound(TempSTR) > 0 Then
            Lpdes = TempSTR(1)
        End If
End If
    .pdes.Text = BlankCheck(.pdes.Text, Lpdes)
        
    .pxdes1.Text = BlankCheck(.pxdes1.Text, Lpxdes1)
    .pxdes2.Text = BlankCheck(.pxdes2.Text, Lpxdes2)
On Error GoTo 0
End With

If chkLoadTol.Value = 1 Then
    Call LoadTol
End If


Set oPropSet = Nothing
Set oPropSets = Nothing
Set tDoc = Nothing
OpenFlag = True

Exit Sub
openerror:
MsgBox "Cannot find file: " & cbfile.Text, vbCritical, "Error"
OpenFlag = False
Exit Sub
End Sub
Private Sub LoadTol()
With frmtbm
On Error Resume Next
'On Error GoTo errhand
Dim tempS As String
tempS = .pnoun.Text
    If Not Len(tempS) = 0 Or Not UCase(tempS) = "UNASSIGNED" Then
        Call dbOpen
        Dim mySQL As String
        mySQL = "select * from [templates$] where name='" & tempS & "';"
        rs.Open mySQL, conn, 1, 3
        .pangles.Text = checknull(rs("pangles"))
        .pfinish.Text = checknull(rs("pfinish"))
        .pholel.Text = checknull(rs("pholel"))
        .pholeh.Text = checknull(rs("pholeh"))
        .px.Text = checknull(rs("px"))
        .pxx.Text = checknull(rs("pxx"))
        .pxxx.Text = checknull(rs("pxxx"))
        Call dbClose
    End If
End With
Exit Sub
'errhand:
'MsgBox "Tol Sel Err:" & Err.Description
'Resume Next
End Sub

Private Function BlankCheck(s1, s2)
If chkOnlyBlanks.Value = 1 Then
    If Len(s1) = 0 Then
        BlankCheck = s2
    Else:
        BlankCheck = s1
    End If
Else:
    BlankCheck = s2
End If
End Function
Private Function oPro(pname As String)
oPro = "" & oPropSet.Item(pname).Value
End Function
Private Sub Form_Unload(Cancel As Integer)
Call LoadConfigSave
End Sub
Private Sub LanINI()
Me.Caption = LRS(301)
Label1.Caption = LRS(302)
Frame1.Caption = LRS(303)
chkOnlyBlanks.Caption = LRS(304)
chkLoadTol.Caption = LRS(305)
chkLfp.Caption = LRS(306)
CommandButton2.Caption = LRS(307)
CommandButton3.Caption = LRS(308)
CommandButton1.ToolTipText = LRS(309)
End Sub
