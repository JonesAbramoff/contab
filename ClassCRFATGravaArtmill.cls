VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Vendedor_Grava_Customizacao(ByVal objVendedor As ClassVendedor) As Long
'grava a customizacao de vendedor

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Vendedor_Grava_Customizacao

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183137
    Next
        
    'Altera registro na tabela Vendedores
    lErro = Comando_Executar(alComando(1), "UPDATE FiliaisClientes SET ComissaoVendas = ? WHERE Vendedor = ?", objVendedor.dPercComissao, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 183138

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Vendedor_Grava_Customizacao = SUCESSO

    Exit Function

Erro_Vendedor_Grava_Customizacao:

    Vendedor_Grava_Customizacao = gErr

        Select Case gErr

        Case 183137
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 183138
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_VENDEDOR", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183139)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

