VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ObtemTela(sNomeTela As String, objForm As Object, iTelaNova As Integer) As Long
'iTelaNova =1 : tela criada agora, iTelaNova = 0: tela já estava carregada

Dim objFormTemp As Object, lErro As Long
Dim sTela As String
Dim colBrowseCampo As New Collection

On Error GoTo Erro_ObtemTela

    For Each objFormTemp In Forms
    
        If objFormTemp.Name = "PrincPopUp" Then
        
            If objFormTemp.sNomeTelaOcx = sNomeTela Then
            
                Set objForm = objFormTemp
                iTelaNova = 0
                Exit For
                
            End If
        
        End If
        
    Next

    If objForm Is Nothing Then
    
'        If giTelaTamanhoVariavel = 0 Then

            If InStr(sNomeTela, "Lista") <> 0 Then

                sTela = Mid(sNomeTela, InStr(sNomeTela, ".") + 1)

                If UCase(right(sTela, 3)) = "OCX" Then sTela = left(sTela, Len(sTela) - 3)

                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
                If lErro <> SUCESSO Then Error 1000

                If colBrowseCampo.Count = 0 Then GoTo Label_ObtemTelaMdiChild

                Set objForm = New PrincPopUpSize
            Else
Label_ObtemTelaMdiChild:
                If giTelaTamanhoVariavel = 0 Then
    
                    Set objForm = New PrincPopUp
                    
                Else
                
                    Set objForm = New PrincPopUpSize
                    
                End If
            End If

            lErro = objForm.Iniciar(sNomeTela)
            If lErro <> SUCESSO Then Error 41939

            iTelaNova = 1

 '       Else

     

'            Set objForm = New PrincMDIChildSize
'
'            sTela = Mid(sNomeTela, InStr(sNomeTela, ".") + 1)
'
'            If UCase(right(sTela, 3)) = "OCX" Then sTela = left(sTela, Len(sTela) - 3)
'
'            If InStr(sNomeTela, "Lista") <> 0 Then
'
'                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
'                If lErro <> SUCESSO Then Error 1000
'
'            End If
'
'
'            lErro = objForm.Iniciar(sNomeTela)
'            If lErro <> SUCESSO Then Error 41939
'
'            iTelaNova = 1

'        End If
    
    
    
    
    
    
    
    
    
    
    
    
'        Set objForm = New PrincPopUp
'
'        lErro = objForm.Iniciar(sNomeTela)
'        If lErro <> SUCESSO Then Error 56582
'
'        iTelaNova = 1
        
    End If
    
    ObtemTela = SUCESSO
    
    Exit Function
    
Erro_ObtemTela:

    ObtemTela = Err
    
    Select Case Err
    
        Case 56582
            Unload objForm
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153601)
     
    End Select
     
    Exit Function
        
End Function

Public Sub UnloadTela(objForm As Object)
    Unload objForm
End Sub

