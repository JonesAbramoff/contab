VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Wizard Template"
Option Explicit

'Implements IDTExtensibility
'
'Dim mcbMenuCommandBar         As Office.CommandBarControl  'command bar object
'Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
'Dim mfrmWizard As frmWizardEmpresa
'Dim VBInstance As VBIDE.VBE
'
'
''------------------------------------------------------
''this method adds the Add-In to the VB menu
''it is called by the VB addin manager
''------------------------------------------------------
'Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal LoadMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
'   On Error GoTo error_handler
'
'   Set VBInstance = VBInst
'
'   If LoadMode = vbext_cm_External Then
'       'Used by the wizard toolbar to start this wizard
'       LoadMe
'   Else
'       Set mcbMenuCommandBar = AddToAddInCommandBar(VBInst, LoadResString(15), LoadResPicture(5000, 0))
'       'sink the event
'       Set Me.MenuHandler = VBInst.Events.CommandBarEvents(mcbMenuCommandBar)
'   End If
'
'   Exit Sub
'
'error_handler:
'   MsgBox Err.Description
'End Sub
'
''------------------------------------------------------
''this method removes the Add-In from the VB menu
''it is called by the VB addin manager
''------------------------------------------------------
'Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
'    'delete the command bar entry
'    mcbMenuCommandBar.Delete
'End Sub
'
'Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'  'stub needed
'End Sub
'
'Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'  'stub needed
'End Sub
'
''this event fires when the menu is clicked in the IDE
'Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'    LoadMe
'End Sub
'
'Private Sub LoadMe()
'    Set mfrmWizard = New frmWizard
'    'pass the vb instance to the wizard module
'    Set mfrmWizard.VBInst = VBInstance
'    'load and show the form
'    mfrmWizard.Show vbModal
'    Set mfrmWizard = Nothing
'End Sub
'

Public Property Get objfrmWizardEmpresa() As Object
    Set objfrmWizardEmpresa = frmWizardEmpresa
End Property

Public Property Get objfrmWizardFilial() As Object
    Set objfrmWizardFilial = frmWizardFilial
End Property

