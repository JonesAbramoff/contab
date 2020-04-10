VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelItemNFLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'para obter informacoes de lotes vinculados na itens de nf para impressao da nf

Private lUltNumIntDoc As Long
Private lComando As Long
Private giQuantLotes As Integer
Private gasLotes(1 To 20) As String
Private gadQtd(1 To 20) As Double
Private gadtValidade(1 To 20) As Date

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

Public Function ItemNF_ObterLote(ByVal lNumIntDoc As Long, ByVal iSeq As Integer, sLote As String) As Long
'Para obter a Base de Calculo
'para iSeq = 0 só preencher se houver apenas um lote para o item

Dim lErro As Long

On Error GoTo Erro_ItemNF_ObterLote
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelItemNF_LoteInfo(lNumIntDoc)
        If lErro <> SUCESSO Then gError 130338
        
    End If
    
    If giQuantLotes = 0 Or iSeq > giQuantLotes Then
    
        sLote = ""
        
    Else
    
        If iSeq = 0 Then
        
            If giQuantLotes > 1 Then
                sLote = ""
            Else
                sLote = gasLotes(1)
            End If
            
        Else
        
            If giQuantLotes = 1 Then
                sLote = ""
            Else
                sLote = gasLotes(iSeq)
            End If
            
        End If
    
    End If
    
    ItemNF_ObterLote = SUCESSO
    
    Exit Function
    
Erro_ItemNF_ObterLote:

    ItemNF_ObterLote = gErr
    
    Select Case gErr

        Case 130338
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153186)

    End Select

    Exit Function

End Function

Public Function ItemNF_ObterLoteQtd(ByVal lNumIntDoc As Long, ByVal iSeq As Integer, dQuantidade As Double) As Long
'Para obter a Base de Calculo
'para iSeq = 0 só preencher se houver apenas um lote para o item

Dim lErro As Long

On Error GoTo Erro_ItemNF_ObterLoteQtd
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelItemNF_LoteInfo(lNumIntDoc)
        If lErro <> SUCESSO Then gError 130339
    
    End If
    
    If giQuantLotes = 0 Or iSeq > giQuantLotes Then
    
        dQuantidade = 0
        
    Else
    
        If iSeq = 0 Then
        
            If giQuantLotes > 1 Then
                dQuantidade = 0
            Else
                dQuantidade = gadQtd(1)
            End If
            
        Else
        
            If giQuantLotes = 1 Then
                dQuantidade = 0
            Else
                dQuantidade = gadQtd(iSeq)
            End If
        End If
    
    End If
    
    ItemNF_ObterLoteQtd = SUCESSO
    
    Exit Function
    
Erro_ItemNF_ObterLoteQtd:

    ItemNF_ObterLoteQtd = gErr
    
    Select Case gErr

        Case 130339
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153187)

    End Select

    Exit Function

End Function

Private Function ObterRelItemNF_LoteInfo(ByVal lNumIntItemNF As Long) As Long
'Função que vai ao BD e lê info sobre Lotes para um determinado item de nf passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDocMovEst As Long
Dim iIndice As Integer
Dim sRetornoCodLote As String
Dim sSQL As String, dRetornoQtdeLote As Double, dtValidade As Date

On Error GoTo Erro_ObterRelItemNF_LoteInfo
 
    lUltNumIntDoc = 0
    
    giQuantLotes = 0
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 111301
    
    'Inicializar a String
    sRetornoCodLote = String(STRING_LOTE_RASTREAMENTO, 0)
    
    'Lê os Lotes Fazendo o Join das tabelas ItensNiscal , Movimentos Estoque , RastreamentoMovto , RastreamentoLote
    lErro = Comando_Executar(lComando, "SELECT MIN(RastreamentoLote.DataValidade), RastreamentoLote.Lote, SUM(RastreamentoMovto.Quantidade) FROM RastreamentoMovto, RastreamentoLote, MovimentoEstoque, ItensNFiscal " _
    & "WHERE ItensNFiscal.NumIntDoc = ? AND ItensNFiscal.NumIntDoc = MovimentoEstoque.NumIntDocOrigem AND MovimentoEstoque.TipoNumIntDocOrigem = 1 AND RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastreamentoMovto.TipoDocOrigem = 0 AND " _
    & "RastreamentoMovto.NumIntDocOrigem = MovimentoEstoque.NumIntDoc GROUP BY RastreamentoLote.Lote ORDER BY RastreamentoLote.Lote", dtValidade, sRetornoCodLote, dRetornoQtdeLote, lNumIntItemNF)
    If lErro <> AD_SQL_SUCESSO Then gError 111302
     
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111303
    
    Do While lErro = SUCESSO
    
        giQuantLotes = giQuantLotes + 1
        gadtValidade(giQuantLotes) = dtValidade
        gasLotes(giQuantLotes) = sRetornoCodLote
        gadQtd(giQuantLotes) = dRetornoQtdeLote
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111305
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)
       
    lUltNumIntDoc = lNumIntItemNF
        
    ObterRelItemNF_LoteInfo = SUCESSO

    Exit Function

Erro_ObterRelItemNF_LoteInfo:

    ObterRelItemNF_LoteInfo = gErr

    Select Case gErr

        Case 111301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111302, 111303, 111305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE_VAZIO_NF", gErr, lNumIntItemNF)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153188)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ItemNF_ObterLoteVal(ByVal lNumIntDoc As Long, ByVal iSeq As Integer, dtValidade As Date) As Long
'Para obter a Base de Calculo
'para iSeq = 0 só preencher se houver apenas um lote para o item

Dim lErro As Long

On Error GoTo Erro_ItemNF_ObterLoteVal
    
    If lNumIntDoc <> lUltNumIntDoc Then
        
        lErro = ObterRelItemNF_LoteInfo(lNumIntDoc)
        If lErro <> SUCESSO Then gError 130339
    
    End If
    
    If giQuantLotes = 0 Or iSeq > giQuantLotes Then
    
        dtValidade = DATA_NULA
        
    Else
    
        If iSeq = 0 Then
        
            If giQuantLotes > 1 Then
                dtValidade = DATA_NULA
            Else
                dtValidade = gadtValidade(1)
            End If
            
        Else
        
            If giQuantLotes = 1 Then
                dtValidade = DATA_NULA
            Else
                dtValidade = gadtValidade(iSeq)
            End If
            
        End If
    
    End If
    
    ItemNF_ObterLoteVal = SUCESSO
    
    Exit Function
    
Erro_ItemNF_ObterLoteVal:

    ItemNF_ObterLoteVal = gErr
    
    Select Case gErr

        Case 130339
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153187)

    End Select

    Exit Function

End Function

