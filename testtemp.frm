VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmtbm 
   Caption         =   "tbm"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8880
   StartUpPosition =   3  '窗口缺省
   Begin MSForms.ScrollBar ScrollBar1 
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   3360
      Width           =   495
      Size            =   "873;873"
      Min             =   1
      Max             =   100
      Position        =   1
      SmallChange     =   -1
      LargeChange     =   -1
   End
   Begin MSForms.TextBox prevision 
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
      VariousPropertyBits=   746604571
      Size            =   "5106;873"
      Value           =   "01"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   5400
      Width           =   2175
      Size            =   "3836;1296"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox pxdes2 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
      VariousPropertyBits=   746604571
      Size            =   "5953;1085"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox pxdes1 
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
      VariousPropertyBits=   746604571
      Size            =   "5741;1085"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox pdes 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3255
      VariousPropertyBits=   746604571
      Size            =   "5741;661"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox pnoun 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      VariousPropertyBits=   746604571
      Size            =   "5741;873"
      FontName        =   "宋体"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frmtbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
frmSpellCheck.Show 1, Me
End Sub

Private Sub Form_Load()
With prevision
If IsNumeric(.Text) Then
ScrollBar1.Value = .Text
Else:
    If .Text Like "[A-Z]" Or .Text Like "[a-z]" Then
    ScrollBar1.Value = Asc(UCase(.Text))
    End If
End If
End With
End Sub

Private Sub prevision_Change()
With prevision
If .Text Like "[A-Z]" Or .Text Like "[a-z]" Then
    ScrollBar1.Value = Asc(UCase(.Text))
    ScrollBar1.Min = 65
    ScrollBar1.Max = 65 + 25
Else:
    If IsNumeric(.Text) Then
        If .Text > 0 And .Text <= 100 Then
        ScrollBar1.Min = 1
        ScrollBar1.Max = 100
        ScrollBar1.Value = Int(.Text)
        End If
    End If
End If
End With
End Sub

Private Sub ScrollBar1_Change()
With prevision
If IsNumeric(.Text) Then
    If .Text > 0 And .Text <= 100 Then
        If ScrollBar1.Value < 10 Then
        .Text = "0" & ScrollBar1.Value
        Else:
        .Text = ScrollBar1.Value
        End If
    End If
Else:
    If .Text Like "[A-Z]" Or .Text Like "[a-z]" Then
    .Text = Chr(ScrollBar1.Value)
    End If
End If
End With
End Sub


