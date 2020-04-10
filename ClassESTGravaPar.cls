VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function RelRotulo_Customiza(ByVal objOP As ClassOrdemDeProducao, lNumIntRel As Long) As Long
'Grava um registro para cada batelada

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim objOPOper As ClassOrdemProducaoOperacoes
Dim objOPInsumo As ClassOrdemProducaoInsumos
Dim objItemOP As ClassItemOP
Dim dQuantidade As Double
Dim iBateladas As Integer
Dim dBateladas As Double
Dim iSeq As Integer

On Error GoTo Erro_RelRotulo_Customiza

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189055

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 189056
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_PROX_REL_ROTULO", lNumIntRel)
    If lErro <> SUCESSO Then gError 189057
    
    For Each objItemOP In objOP.colItens
        For Each objOPOper In objItemOP.colOrdemProducaoOperacoes
            For Each objOPInsumo In objOPOper.colOPInsumos
            
                dQuantidade = 0
                lErro = Comando_Executar(alComando(1), "SELECT Quantidade FROM ProdutoKit WHERE ProdutoRaiz = ? AND Versao = ? AND Nivel = ?", dQuantidade, objOPInsumo.sProduto, objOPInsumo.sVersaoKitComp, KIT_NIVEL_RAIZ)
                If lErro <> AD_SQL_SUCESSO Then gError 189058
            
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 189059

                If dQuantidade <> 0 Then
                    dBateladas = objOPInsumo.dQuantidade / dQuantidade
                    
                    If Abs(dBateladas - Fix(dBateladas)) > QTDE_ESTOQUE_DELTA Then
                        iBateladas = Fix(dBateladas) + 1
                    Else
                        iBateladas = Fix(dBateladas)
                    End If
                Else
                    iBateladas = 0
                End If
            
                For iIndice = 1 To iBateladas
                
                    iSeq = iSeq + 1
                
                    'Insere no BD
                    lErro = Comando_Executar(alComando(2), "INSERT INTO RelRotuloPar(NumIntRel, NumIntDocOper, Produto, Seq) VALUES (?,?,?,?)", lNumIntRel, objOPInsumo.lNumIntDocOper, objOPInsumo.sProduto, iSeq)
                    If lErro <> AD_SQL_SUCESSO Then gError 189060
                Next
                
            Next
            
        Next
        
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189061

    RelRotulo_Customiza = SUCESSO

    Exit Function

Erro_RelRotulo_Customiza:

    RelRotulo_Customiza = gErr

    Select Case gErr

        Case 189055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 189056
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 189057
        
        Case 189058, 189059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)
            
        Case 189060
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELROTULO", gErr)

        Case 189061
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189062)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback

    Exit Function

End Function

