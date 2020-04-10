VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProjetosRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get Demonstrativo(ByVal objCT As Object) As Object
     Set Demonstrativo = objCT.objUserControl.Controls("Demonstrativo")
End Property

Public Function PRJ_Limpa_Tela_Projetos(ByVal objCT As Object) As Long

    Demonstrativo(objCT).Caption = "000"

    PRJ_Limpa_Tela_Projetos = SUCESSO
End Function

Public Function PRJ_Traz_Projetos_Tela(ByVal objCT As Object, ByVal objProjeto As ClassProjetos) As Long
    
Dim lErro As Long
Dim objProjetoRL As ClassProjetosRL

On Error GoTo Erro_PRJ_Traz_Projetos_Tela

    lErro = CF("Projetos_le_Customizado", objProjeto)
    If lErro <> SUCESSO Then gError 187493
    
    If Not (objProjeto.objInfoUsu Is Nothing) Then
    
        Set objProjetoRL = objProjeto.objInfoUsu
    
        Demonstrativo(objCT).Caption = Format(objProjetoRL.lDemonstrativo, "000")
    
    End If
    
    PRJ_Traz_Projetos_Tela = SUCESSO
    
    Exit Function

Erro_PRJ_Traz_Projetos_Tela:

    PRJ_Traz_Projetos_Tela = gErr

    Select Case gErr
    
        Case 187493

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185601)

    End Select

    Exit Function
    
End Function
