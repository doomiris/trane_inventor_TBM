VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSpellCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpellCheck"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SpellCheck.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDict 
      Caption         =   "&Dict"
      Height          =   350
      Left            =   2160
      TabIndex        =   8
      Top             =   4200
      Width           =   800
   End
   Begin VB.TextBox tb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Preview Text"
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   3000
      TabIndex        =   2
      Top             =   4200
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Back"
      Height          =   350
      Left            =   3840
      TabIndex        =   3
      Top             =   4200
      Width           =   800
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
      ForeColor       =   255
      VariousPropertyBits=   8388627
      Caption         =   "No Spell Error found!"
      Size            =   "3413;450"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox SugList 
      Height          =   2820
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   2100
      VariousPropertyBits=   746586139
      ForeColor       =   16384
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3704;4579"
      MatchEntry      =   0
      ListStyle       =   1
      SpecialEffect   =   0
      FontName        =   "Times New Roman"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox WordList 
      Height          =   2820
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2100
      VariousPropertyBits=   746586139
      ForeColor       =   255
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3704;4579"
      MatchEntry      =   0
      ListStyle       =   1
      SpecialEffect   =   0
      FontName        =   "Times New Roman"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "Spell Suggestion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Iffy Words:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim wDoc As New Word.Document
Dim oSugs As SpellingSuggestions
Dim i As Integer
Dim TempSTR As String
Dim WordFrom As String
Dim oWordStr As String
Dim WC() As String
Dim WordCount As Integer
Dim tempNum As Integer
Dim IniDone As Boolean
Dim PleaseWait As String

Private Sub Form_Load()
PleaseWait = "Please Wait..."
If TbmLan > 0 Then
    Call LanINI
End If
IniDone = False
WordFrom = "pnoun"
oWordStr = LCase("" & frmtbm.pnoun.Text)
Call ErrorWordRed

WordFrom = "pdes"
oWordStr = LCase("" & frmtbm.pdes.Text)
Call ErrorWordRed

WordFrom = "pxdes1"
oWordStr = LCase("" & frmtbm.pxdes1.Text)
Call ErrorWordRed

WordFrom = "pxdes2"
oWordStr = LCase("" & frmtbm.pxdes2.Text)
Call ErrorWordRed

If WordList.ListCount > 0 Then
WordList.ListIndex = 0
Call WordList_Click
Else:
lbl1.Visible = True
End If
IniDone = True


End Sub
Private Sub ErrorWordRed()
If oWordStr = "" Then
Exit Sub
Else:
wDoc.Range = oWordStr
End If
With wDoc.Range.Words
tempNum = 0
For i = 1 To .Count                             '检查传送过来的每个单词
    Set oSugs = .Item(i).GetSpellingSuggestions(App.Path & "\InvUser.dic")
        If oSugs.Count > 0 Then                 '如果存在拼写建议则加入iffy列表
        WordList.AddItem Trim(.Item(i).Text)
        tempNum = tempNum + 1                   '并统计个数
        End If
Next
End With

WordCount = WordList.ListCount                  'iffy列表单词数

ReDim Preserve WC(3, WordCount)                 '4列 数组定义 并 保留上次的数据

For i = WordCount - tempNum To WordCount - 1
WC(1, i) = WordFrom                             '数据来源
WC(2, i) = UCase(SugList.Text)                  '首选建议
WC(3, i) = WordList.Text                        '原始单词
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
Command4.Enabled = False
Command3.Enabled = False
Call ExitSpell
End Sub

Private Sub WordList_Click()

    SugList.Clear
    wDoc.Range = WordList.List(WordList.ListIndex)
    Set oSugs = wDoc.Range.GetSpellingSuggestions(App.Path & "\InvUser.dic")
    For i = 1 To oSugs.Count
    SugList.AddItem oSugs.Item(i)
    Next

If SugList.ListCount > 0 Then

    SugList.ListIndex = Val(WC(0, WordList.ListIndex))
    WC(2, WordList.ListIndex) = UCase(SugList.List(SugList.ListIndex))
    WC(3, WordList.ListIndex) = WordList.List(WordList.ListIndex)

End If
Call TextPreview
End Sub

Private Sub SugList_Click()
WC(0, WordList.ListIndex) = SugList.ListIndex
WC(2, WordList.ListIndex) = UCase(SugList.Text)
WC(3, WordList.ListIndex) = WordList.Text
    If IniDone = True Then
    WordList.SetFocus
    End If
Call TextPreview
End Sub


Private Sub TextPreview()
tb1.Text = UCase(SugList.List(0))
End Sub
Private Sub ExitSpell()
If wDoc.Application.Visible = True Then
wDoc.Close (False)
Else:
wDoc.Application.Quit savechanges:=False
End If
Set wDoc = Nothing
Erase WC()
End Sub
Private Sub Command3_Click()
lbl1.Visible = True
lbl1.Caption = PleaseWait
Command4.Enabled = False
Command3.Enabled = False
Call ExitSpell
Unload Me
frmtbm.pnoun.SetFocus
End Sub

Private Sub Command4_Click()
lbl1.Visible = True

lbl1.Caption = PleaseWait

Command4.Enabled = False
Command3.Enabled = False
If WordList.ListCount > 0 Then
    With frmtbm
    For i = 0 To WordList.ListCount
        Select Case WC(1, i)
        Case "pnoun"
            .pnoun.Text = Replace(.pnoun.Text, WC(3, i), WC(2, i), 1, -1, vbTextCompare)
        Case "pdes"
            .pdes.Text = Replace(.pdes.Text, WC(3, i), WC(2, i), 1, -1, vbTextCompare)
        Case "pxdes1"
            .pxdes1.Text = Replace(.pxdes1.Text, WC(3, i), WC(2, i), 1, -1, vbTextCompare)
        Case "pxdes2"
            .pxdes2.Text = Replace(.pxdes2.Text, WC(3, i), WC(2, i), 1, -1, vbTextCompare)
        End Select
    Next
    End With
End If
Call ExitSpell
Unload Me
frmtbm.pnoun.SetFocus
End Sub
Private Sub cmdDict_Click()
    On Error Resume Next
    Call ShellExecute(Me.hwnd, "Open", App.Path & "\InvUser.dic", "", App.Path, 1)
    On Error GoTo 0

End Sub
Private Sub LanINI()
Me.Caption = LRS(251)
Label1.Caption = LRS(252)
Label2.Caption = LRS(253)
cmdDict.Caption = LRS(254)
Command4.Caption = LRS(255)
Command3.Caption = LRS(256)
lbl1.Caption = LRS(257)
PleaseWait = LRS(258)
tb1.ToolTipText = LRS(261)

End Sub

