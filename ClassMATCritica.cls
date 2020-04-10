VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATCritica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Function Produto_Critica_Estoque(ByVal sProduto As String, ByVal objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL e participa do ESTOQUE.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido

Dim lErro As Long
Dim sProdutoFormatado As String

On Error GoTo Erro_Produto_Critica_Estoque

    If Len(sProduto) > 0 Then
    
        'Formata o código do Produto como no BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro Then Error 25075
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            'Carrega sProdutoFormatado em objProduto
            objProduto.sCodigo = sProdutoFormatado
            
            'Tenta ler o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 25076
    
            'Ausência de Produto no BD
            If lErro = 28030 Then Error 25077
            
            'Critica se é produto ATIVO
            If objProduto.iAtivo <> Ativo Then Error 25078
            
            'Critica se é produto NAO_GERENCIAL
            If objProduto.iGerencial <> NAO_GERENCIAL Then Error 25079
            
            'Critica se é produto de ESTOQUE (c/ reserva ou sem)
            If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then Error 25080
            
        
        End If
    
    End If
    
    Produto_Critica_Estoque = SUCESSO
    
    Exit Function

Erro_Produto_Critica_Estoque:

    Produto_Critica_Estoque = Err
    
    Select Case Err
    
        Case 25075, 25076  'Erro criticado na rotina chamada
        
        Case 25077  'Produto não está cadastrado
        
        Case 25078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", Err, sProduto)
        
        Case 25079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
        
        Case 25080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", Err, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151558)
        
    End Select
    
    Exit Function

End Function

Function InvLote_Critica_Atualizado(ByVal objInvLote As ClassInvLote, iLoteAtualizado As Integer) As Long
'testa se o Invlote já foi atualizado. Se já foi iLoteAtualizado = LOTE_ATUALIZADO senao retorna 0.

Dim lComando As Long
Dim lErro As Long
Dim iLote As Integer

On Error GoTo Erro_InvLote_Critica_Atualizado

    lComando = 0

    iLoteAtualizado = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41232

    'pesquisa no BD o lote em questão
    lErro = Comando_Executar(lComando, "SELECT Lote FROM InvLote WHERE FilialEmpresa =? AND Lote=?", iLote, objInvLote.iFilialEmpresa, objInvLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 41233

    'le o lote, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41236

    If lErro = AD_SQL_SUCESSO Then iLoteAtualizado = LOTE_ATUALIZADO

    Call Comando_Fechar(lComando)

    InvLote_Critica_Atualizado = SUCESSO

    Exit Function

Erro_InvLote_Critica_Atualizado:

    InvLote_Critica_Atualizado = Err

    Select Case Err

        Case 41232
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41233, 41236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVLOTE", Err, objInvLote.iFilialEmpresa, objInvLote.iLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151559)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function InventarioPendente_Critica_Lote(ByVal objInvLote As ClassInvLote, iNumIguais As Integer) As Long
'verifica se o numero de lotes em InventarioLotePedente é igual ao de InvLotePendente
 
Dim lErro As Long
Dim lComando1 As Long
Dim sProduto As String
Dim sComando_SQL As String
Dim iNumItensAtual As Integer
    
On Error GoTo Erro_InventarioPendente_Critica_Lote
    
    'abre o comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 52163
    
    sProduto = String(STRING_PRODUTO, 0)
    iNumItensAtual = 0
    
    'monta o comando SQL
    sComando_SQL = "SELECT InventarioPendente.Produto FROM InventarioPendente, Produtos, Almoxarifado WHERE InventarioPendente.Produto=Produtos.Codigo AND InventarioPendente.Almoxarifado=Almoxarifado.Codigo AND InventarioPendente.Lote = ? AND InventarioPendente.FilialEmpresa = ? ORDER BY InventarioPendente.NumIntDoc"
    
    lErro = Comando_Executar(lComando1, sComando_SQL, sProduto, objInvLote.iLote, giFilialEmpresa)
    If lErro <> SUCESSO Then Error 52166

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52167

    Do While lErro <> AD_SQL_SEM_DADOS
        
        iNumItensAtual = iNumItensAtual + 1
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52169

    Loop
    
    If iNumItensAtual = objInvLote.iNumItensAtual Then
         iNumIguais = IGUAL
    Else
        iNumIguais = DIFERENTE
        objInvLote.iNumItensAtual = iNumItensAtual
    End If
        
    Call Comando_Fechar(lComando1)
        
    InventarioPendente_Critica_Lote = SUCESSO
    
    Exit Function

Erro_InventarioPendente_Critica_Lote:

    InventarioPendente_Critica_Lote = Err

    Select Case Err

        Case 52163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52166, 52167, 52169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOPENDENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151560)

    End Select
    
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function Produto_Critica_Compra(ByVal sProduto As String, ByVal objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL e é COMPRÁVEL.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido

Dim lErro As Long
Dim sProdutoFormatado As String
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Produto_Critica_Compra

    If Len(sProduto) > 0 Then
    
        'Formata o código do Produto como no BD
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro Then Error 25533
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            'Carrega sProdutoFormatado em objProduto
            objProduto.sCodigo = sProdutoFormatado
            
            'Tenta ler o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 25599
    
            'Ausência de Produto no BD
            If lErro = 28030 Then Error 25605
            
            'Critica se é produto ATIVO
            If objProduto.iAtivo <> Ativo Then Error 25606
            
            'Critica se é produto NAO_GERENCIAL
            If objProduto.iGerencial <> NAO_GERENCIAL Then Error 25607
            
            'Critica se é produto comprável
            'If objProduto.iCompras = PRODUTO_PRODUZIVEL Then Error 25608
            
            'Critica se o produto está cadastrado nessa Filial
            objProdutoFilial.sProduto = objProduto.sCodigo
            objProdutoFilial.iFilialEmpresa = giFilialEmpresa
            
            lErro = CF("ProdutoFilial_Le", objProdutoFilial)
            If lErro <> SUCESSO And lErro <> 28261 Then Error 25623
            
            'Se Produto não está cadastrado na Filial, erro
            If lErro = 28261 Then Error 25624
        
        End If
    
    End If
    
    Produto_Critica_Compra = SUCESSO
    
    Exit Function

Erro_Produto_Critica_Compra:

    Produto_Critica_Compra = Err
    
    Select Case Err
    
        Case 25533, 25599, 25623  'Erro criticado na rotina chamada
        
        Case 25605  'Produto não está cadastrado
        
        Case 25606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", Err, sProduto)
        
        Case 25607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
        
        Case 25608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", Err, sProduto)
        
        Case 25624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_NAO_CADASTRADO", Err, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151561)
        
    End Select
    
    Exit Function

End Function

