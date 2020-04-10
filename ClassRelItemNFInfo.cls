VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelItemNFInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lUltNumIntDoc As Long
Private dUltIPIBaseCalculo As Double
Private dUltIPICredito As Double

Private lComando As Long

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

'***** 06/06/01 Livro modelo 3 - Maristela (Revisar)

Public Function ItemNF_ObterIPIBase(lNumIntDoc As Long, dIPIBaseCalculo As Double) As Long
'Para obter a Base de Calculo

Dim lErro As Long

On Error GoTo Erro_ItemNF_ObterIPIBase
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelItemNF_IPIBaseCredito(lNumIntDoc)
        If lErro <> SUCESSO Then gError 90337
    
    End If
    
    dIPIBaseCalculo = dUltIPIBaseCalculo
    
    ItemNF_ObterIPIBase = SUCESSO
    
    Exit Function
    
Erro_ItemNF_ObterIPIBase:

    ItemNF_ObterIPIBase = gErr
    
    Select Case gErr

        Case 90337
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153176)

    End Select

    Exit Function

End Function

Public Function ItemNF_ObterIPICredito(lNumIntDoc As Long, dIPICredito As Double) As Long
'Para obter o IPICredito

Dim lErro As Long

On Error GoTo Erro_ItemNF_ObterIPICredito
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelItemNF_IPIBaseCredito(lNumIntDoc)
        If lErro <> SUCESSO Then gError 90338
    
    End If
    
    dIPICredito = dUltIPICredito
    
    ItemNF_ObterIPICredito = SUCESSO
    
    Exit Function
    
Erro_ItemNF_ObterIPICredito:

    ItemNF_ObterIPICredito = gErr
    
    Select Case gErr

        Case 90338
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153177)

    End Select

    Exit Function

End Function

Private Function ObterRelItemNF_IPIBaseCredito(lNumIntDoc As Long) As Long
'Para obter valores e torna-los disponiveis para as outras funcoes

Dim lErro As Long
Dim dIPIBaseCalculo As Double
Dim dIPICredito As Double
Dim dQuantidade As Double
Dim dPrecoUnitario As Double

On Error GoTo Erro_ObterRelItemNF_IPIBaseCredito

    lErro = Obter_IPIBaseCalculo_IPICredito_NumIntDoc(lNumIntDoc, dIPIBaseCalculo, dIPICredito, dQuantidade, dPrecoUnitario, lComando)
    If lErro <> SUCESSO Then gError 90339
    
    lUltNumIntDoc = lNumIntDoc
    
    If dIPICredito > 0 Then
        dUltIPIBaseCalculo = dIPIBaseCalculo
    ElseIf dIPICredito = 0 Then
        dUltIPIBaseCalculo = dQuantidade * dPrecoUnitario
    End If
    
    dUltIPICredito = dIPICredito
    
    ObterRelItemNF_IPIBaseCredito = SUCESSO
    
    Exit Function
    
Erro_ObterRelItemNF_IPIBaseCredito:

    ObterRelItemNF_IPIBaseCredito = gErr
    
    Select Case gErr

        Case 90339
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153178)

    End Select

    Exit Function
    
End Function

Function Obter_IPIBaseCalculo_IPICredito_NumIntDoc(lNumIntDoc As Long, dIPIBaseCalculo As Double, dIPICredito As Double, dQuantidade As Double, dPrecoUnitario, lComando As Long) As Long
'Obtem a IPIBaseCalculo e o IPICredito do NumIntDoc da tabela ItensNFiscal(passada por parametro)

Dim lErro As Long

On Error GoTo Erro_Obter_IPIBaseCalculo_IPICredito_NumIntDoc

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 90340
    End If
    
    'Lê da tabela TributacaoItemNF a IPIBaseCalculo e o IPICredito
    lErro = Comando_Executar(lComando, "SELECT TributacaoItemNF.IPIBaseCalculo, TributacaoItemNF.IPICredito, ItensNFiscal.Quantidade, ItensNFiscal.PrecoUnitario " & _
            "FROM ItensNFiscal, TributacaoItemNF WHERE ItensNFiscal.NumIntDoc = ? AND TributacaoItemNF.Item = ItensNFiscal.Item AND TributacaoItemNF.NumIntNF = ItensNFiscal.NumIntNF ", dIPIBaseCalculo, dIPICredito, dQuantidade, dPrecoUnitario, lNumIntDoc)
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90341

    Obter_IPIBaseCalculo_IPICredito_NumIntDoc = SUCESSO

    Exit Function

Erro_Obter_IPIBaseCalculo_IPICredito_NumIntDoc:

    Obter_IPIBaseCalculo_IPICredito_NumIntDoc = gErr

    Select Case gErr

        Case 90340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153179)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

