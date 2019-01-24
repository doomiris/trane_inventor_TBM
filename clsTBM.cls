VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ApplicationAddInServer

Dim WithEvents tbmButtonDef As ButtonDefinition
Attribute tbmButtonDef.VB_VarHelpID = -1
'Private WithEvents oFileUIEvents As FileUIEvents
Private Property Get ApplicationAddInServer_Automation() As Object
Set ApplicationAddInServer_Automation = Nothing
End Property
Private Sub ApplicationAddInServer_Deactivate()
'Unload All
'Erase c_ver
'Erase c_date
'Erase DateForTrane
'Erase dbpath

tbmButtonDef.Delete
Set tbmButtonDef = Nothing
Set oApp = Nothing
End Sub
Private Sub ApplicationAddInServer_ExecuteCommand(ByVal CommandID As Long)
 ' Not used.
End Sub
Private Sub ApplicationAddInServer_Activate(ByVal AddInSiteObject As Inventor.ApplicationAddInSite, ByVal FirstTime As Boolean)

Set oApp = AddInSiteObject.Application
TbmLan = LS("Languages", 0)
'Set oFileUIEvents = oapp.FileUIEvents


    Dim oCommandMgr As CommandManager
    Set oCommandMgr = oApp.CommandManager
    
    Dim oControlDefs As ControlDefinitions
    Set oControlDefs = oCommandMgr.ControlDefinitions
    
    Dim oIcon As IPictureDisp
    Set oIcon = LoadPicture(App.Path & "\tbm.ico")

    Set tbmButtonDef = oControlDefs.AddButtonDefinition("TBM", "TitleBlockManager", kNonShapeEditCmdType, , "TitleBlockManager", "Title Block Manager", oIcon, oIcon)
    Dim oButtonControl As CommandBarControl
    
If FirstTime = True Then
    Dim oCustomCmdBar As CommandBar
    Set oCustomCmdBar = oApp.UserInterfaceManager.CommandBars.Add("ADDINS", "DominicAddinCmdBar ", kRegularCommandBar)
    Set oButtonControl = oCustomCmdBar.Controls.AddButton(tbmButtonDef)
    oCustomCmdBar.Visible = True
End If
If LS("PreLoad", "1") = "1" Then
    Load frmtbm
End If
End Sub

Private Sub tbmButtonDef_OnExecute(ByVal Context As Inventor.NameValueMap)
If oApp.Documents.Count = 0 Then
    oApp.StatusBarText = "Cann't open TBM. No document open."
    Exit Sub
End If

    Dim oDoc As Document
    Set oDoc = oApp.ActiveDocument
    
    If oDoc.DocumentType = kDrawingDocumentObject Then
        Dim oDrawDoc As DrawingDocument
        Set oDrawDoc = oDoc
        Dim oSheet As Sheet
        Set oSheet = oDrawDoc.ActiveSheet
        If Left(oSheet.name, 5) = "Feuil" And oDrawDoc.TitleBlockDefinitions.Item(1).name = "Trane_ISO" Then
                frmCharms.Show 1
        Else:
                frmtbm.Show 1
        End If
    Else:
        frmtbm.Show 1
    End If

End Sub
