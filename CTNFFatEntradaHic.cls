VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFatEntradaHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- inicio dos properties get dos controles da tela

Public Property Get TaxaConversao(ByVal objCT As Object) As Object
     Set TaxaConversao = objCT.objUserControl.Controls("TaxaConversao")
End Property

'--- fim dos properties get dos controles da tela

Public Sub TaxaConversao_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TaxaConversao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TaxaConversao_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(TaxaConversao(objCT).Text)) > 0 Then

        'Critica se é valor Positivo
        lErro = Valor_Positivo_Critica_Double(TaxaConversao(objCT).Text)
        If lErro <> SUCESSO Then gError 103378
    
        'Põe o valor formatado na tela
        TaxaConversao(objCT).Text = Format(TaxaConversao(objCT).Text, FORMATO_TAXA_CONVERSAO_MOEDA)
        
    End If
    
    Exit Sub
     
Erro_TaxaConversao_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case 103378

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179261)
     
    End Select
     
    Exit Sub

End Sub

Function NFEnt_Traz_NFiscal_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFEnt_Traz_NFiscal_Tela

    If objNFiscal.dTaxaMoeda <> 0 Then
        TaxaConversao(objCT).Text = Format(objNFiscal.dTaxaMoeda, FORMATO_TAXA_CONVERSAO_MOEDA)
    Else
        TaxaConversao(objCT).Text = ""
    End If
    
    NFEnt_Traz_NFiscal_Tela = SUCESSO
     
    Exit Function
    
Erro_NFEnt_Traz_NFiscal_Tela:

    NFEnt_Traz_NFiscal_Tela = gErr
     
    Select Case gErr
          
        Case 132361
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179262)
     
    End Select
     
    Exit Function

End Function

Function NFEnt_Move_Tela_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objNFiscalHic As ClassNFiscalHic

On Error GoTo Erro_NFEnt_Move_Tela_Memoria

    Set objNFiscalHic = New ClassNFiscalHic

    objNFiscal.dTaxaMoeda = StrParaDbl(TaxaConversao(objCT).Text)
        
    NFEnt_Move_Tela_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFEnt_Move_Tela_Memoria:

    NFEnt_Move_Tela_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179263)
     
    End Select
     
    Exit Function

End Function

Function NFEnt_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    'liberar objetos globais referentes a customizacao
    
    NFEnt_Form_Unload = SUCESSO
    
End Function



