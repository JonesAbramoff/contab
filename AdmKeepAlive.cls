VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmKeepAlive"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarColObjs As New Collection

Private Sub Class_Initialize()
    
    'Set SistGlob_objContabInt = CreateObject("GlobaisContab.ClassContabInt")
            
End Sub

Public Sub Add(ByVal vData As Object)

    mvarColObjs.Add vData
    
End Sub

Public Sub Remove(vData As Object)
Dim iIndice As Integer, vAux As Object

    For Each vAux In mvarColObjs
    
        iIndice = iIndice + 1
        If vAux Is vData Then
            Call mvarColObjs.Remove(iIndice)
            Exit For
        End If
        
    Next
    
End Sub

