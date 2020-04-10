VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelPedCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Criada para atender a pedidos de compras

Dim lComando As Long
Dim lNumIntPCAnt As Long
Dim iSeqAnt As Long
Dim colNotasAnt As Collection

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
        
End Sub

Public Function ObterTextoNotasPC(sTexto As String, ByVal lNumIntPC As Long, ByVal iSeq As Integer) As Long
'Função que vai ao BD e le a nota do pedido de numero lNumIntPC + sequencial iSq
'ela retorna em sTexto o conteudo dessa nota lida
'sTexto (OUTPUT)
'lNumIntPC (INPUT)
'iSeq (INPUT)
'Retorna o codigo de erro (LONG)

Dim lErro As Long

On Error GoTo Erro_ObterTextoNotasPC
 
    'se o pedido de compra passado como parametro for diferente do corrente
    If lNumIntPC <> lNumIntPCAnt Then
               
        'le o pedido passado como parametro visando obter a colecao de textos
        'do mesmo...
        'essa colecao esta ordenada por iSeq
        lErro = NotasPedCompra_Le_PedCompra(lNumIntPC, colNotasAnt, lComando)
        If lErro <> SUCESSO Then gError 111840
                        
        'Guarda o numero interno do pedido de compra que acaba de ser lido..
        'esse pedido de compra sera considerado como pedido corrente
        lNumIntPCAnt = lNumIntPC
    
    End If
    
    'Retorna o texto do sequencial em questao
    If iSeq >= 1 And iSeq <= colNotasAnt.Count Then
        sTexto = colNotasAnt(iSeq)
    Else
        sTexto = ""
    End If
        
    ObterTextoNotasPC = SUCESSO

    Exit Function

Erro_ObterTextoNotasPC:

    ObterTextoNotasPC = gErr

    Select Case gErr
        
        Case 111840
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153318)

    End Select

    Exit Function

End Function

Function NotasPedCompra_Le_PedCompra(lNumIntPC As Long, colNotasAnt As Collection, lComando As Long) As Long
'Le as notas relativas ao pedido lNumIntPC (INPUT)
'Retorna em colNotasAnt (OUTPUT) os textos relativos as notas
'mantem os indices da colecao relacionados com o sequencial do pedido
'lComando (INPUT)

Dim lErro As Long
Dim colNotas As New Collection
Dim sNotaPC As String

On Error GoTo Erro_NotasPedCompra_Le_PedCompra

    'se a conexao ainda nao foi aberta... abre...
    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 111841
    End If
    
    'inicializa a string
    sNotaPC = String(150, 0)

    'obtem o conjunto de notas do pedido em questao
    lErro = Comando_Executar(lComando, "SELECT NotaPC FROM NotasPedCompras WHERE NumIntPedido = ? ORDER BY Sequencial", sNotaPC, lNumIntPC)
    If lErro <> AD_SQL_SUCESSO Then gError 111843
    
    'le a primeira nota
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111844
    
    'enquanto for achando notas
    Do While lErro = AD_SQL_SUCESSO
    
        'adiciona a nota na colecao
        colNotas.Add sNotaPC
    
        'busca a proxima nota
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111845
    
    Loop
        
    'retorna a colecao de textos (notas)
    Set colNotasAnt = colNotas
        
    NotasPedCompra_Le_PedCompra = SUCESSO
    
    Exit Function
 
Erro_NotasPedCompra_Le_PedCompra:

    NotasPedCompra_Le_PedCompra = gErr

    Select Case gErr
    
        Case 111841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 111843, 111844, 111845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTASPEDCOMPRAS", gErr, lNumIntPC)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153319)
    
    End Select
    
    Exit Function

End Function

