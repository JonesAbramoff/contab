VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConstCust"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Private SistGlob_STRING_ENDERECO As Integer
'Private SistGlob_STRING_BAIRRO As Integer
'Private SistGlob_STRING_CIDADE As Integer
'
'Private SistGlob_STRING_CLIENTE_RAZAO_SOCIAL As Integer
'Private SistGlob_STRING_CLIENTE_NOME_REDUZIDO As Integer
'
'Private SistGlob_STRING_TRANSPORTADORA_NOME As Integer
'Private SistGlob_STRING_TRANSPORTADORA_NOME_REDUZIDO As Integer

Public Property Let STRING_TRANSPORTADORA_NOME(vData As Integer)
    SistGlob_STRING_TRANSPORTADORA_NOME = vData
End Property

Public Property Get STRING_TRANSPORTADORA_NOME() As Integer

    If SistGlob_STRING_TRANSPORTADORA_NOME = 0 Then
        STRING_TRANSPORTADORA_NOME = 50
    Else
        STRING_TRANSPORTADORA_NOME = SistGlob_STRING_TRANSPORTADORA_NOME
    End If
    
End Property

Public Property Let STRING_PRODUTO_DESCRICAO_TELA(vData As Integer)
    SistGlob_STRING_PRODUTO_DESCRICAO_TELA = vData
End Property

Public Property Get STRING_PRODUTO_DESCRICAO_TELA() As Integer

    If SistGlob_STRING_PRODUTO_DESCRICAO_TELA = 0 Then
        STRING_PRODUTO_DESCRICAO_TELA = 50
    Else
        STRING_PRODUTO_DESCRICAO_TELA = SistGlob_STRING_PRODUTO_DESCRICAO_TELA
    End If
    
End Property

Public Property Let STRING_TRANSPORTADORA_NOME_REDUZIDO(vData As Integer)
    SistGlob_STRING_TRANSPORTADORA_NOME_REDUZIDO = vData
End Property

Public Property Get STRING_TRANSPORTADORA_NOME_REDUZIDO() As Integer

    If SistGlob_STRING_TRANSPORTADORA_NOME_REDUZIDO = 0 Then
        STRING_TRANSPORTADORA_NOME_REDUZIDO = 20
    Else
        STRING_TRANSPORTADORA_NOME_REDUZIDO = SistGlob_STRING_TRANSPORTADORA_NOME_REDUZIDO
    End If
    
End Property

Public Property Let STRING_ENDERECO(vData As Integer)
    SistGlob_STRING_ENDERECO = vData
End Property

Public Property Get STRING_ENDERECO() As Integer

    If SistGlob_STRING_ENDERECO = 0 Then
        STRING_ENDERECO = 40
    Else
        STRING_ENDERECO = SistGlob_STRING_ENDERECO
    End If
    
End Property

Public Property Let STRING_BAIRRO(vData As Integer)
    SistGlob_STRING_BAIRRO = vData
End Property

Public Property Get STRING_BAIRRO() As Integer

    If SistGlob_STRING_BAIRRO = 0 Then
        STRING_BAIRRO = 12
    Else
        STRING_BAIRRO = SistGlob_STRING_BAIRRO
    End If
    
End Property

Public Property Let STRING_CIDADE(vData As Integer)
    SistGlob_STRING_CIDADE = vData
End Property

Public Property Get STRING_CIDADE() As Integer

    If SistGlob_STRING_CIDADE = 0 Then
        STRING_CIDADE = 15
    Else
        STRING_CIDADE = SistGlob_STRING_CIDADE
    End If
    
End Property

Public Property Let STRING_CLIENTE_RAZAO_SOCIAL(vData As Integer)
    SistGlob_STRING_CLIENTE_RAZAO_SOCIAL = vData
End Property

Public Property Get STRING_CLIENTE_RAZAO_SOCIAL() As Integer

    If SistGlob_STRING_CLIENTE_RAZAO_SOCIAL = 0 Then
        STRING_CLIENTE_RAZAO_SOCIAL = 40
    Else
        STRING_CLIENTE_RAZAO_SOCIAL = SistGlob_STRING_CLIENTE_RAZAO_SOCIAL
    End If
    
End Property

Public Property Let STRING_CLIENTE_NOME_REDUZIDO(vData As Integer)
    SistGlob_STRING_CLIENTE_NOME_REDUZIDO = vData
End Property

Public Property Get STRING_CLIENTE_NOME_REDUZIDO() As Integer

    If SistGlob_STRING_CLIENTE_NOME_REDUZIDO = 0 Then
        STRING_CLIENTE_NOME_REDUZIDO = 20
    Else
        STRING_CLIENTE_NOME_REDUZIDO = SistGlob_STRING_CLIENTE_NOME_REDUZIDO
    End If
    
End Property

Public Property Let STRING_CLIENTE_OBSERVACAO(vData As Integer)
    SistGlob_STRING_CLIENTE_OBSERVACAO = vData
End Property

Public Property Get STRING_CLIENTE_OBSERVACAO() As Integer

    If SistGlob_STRING_CLIENTE_OBSERVACAO = 0 Then
        STRING_CLIENTE_OBSERVACAO = 100
    Else
        STRING_CLIENTE_OBSERVACAO = SistGlob_STRING_CLIENTE_OBSERVACAO
    End If
    
End Property

Public Property Let NUM_MAX_ITENS_DISTRIBUICAO(vData As Integer)
    SistGlob_NUM_MAX_ITENS_DISTRIBUICAO = vData
End Property

Public Property Get NUM_MAX_ITENS_DISTRIBUICAO() As Integer

    If SistGlob_NUM_MAX_ITENS_DISTRIBUICAO = 0 Then
        NUM_MAX_ITENS_DISTRIBUICAO = 100
    Else
        NUM_MAX_ITENS_DISTRIBUICAO = SistGlob_NUM_MAX_ITENS_DISTRIBUICAO
    End If
    
End Property

Public Property Let NUM_MAX_ITENS_PEDIDO_COMPRAS(vData As Integer)
    SistGlob_NUM_MAX_ITENS_PEDIDO_COMPRAS = vData
End Property

Public Property Get NUM_MAX_ITENS_PEDIDO_COMPRAS() As Integer

    If SistGlob_NUM_MAX_ITENS_PEDIDO_COMPRAS = 0 Then
        NUM_MAX_ITENS_PEDIDO_COMPRAS = 100
    Else
        NUM_MAX_ITENS_PEDIDO_COMPRAS = SistGlob_NUM_MAX_ITENS_PEDIDO_COMPRAS
    End If
    
End Property

Public Property Let NUM_MAX_ITENS_PEDIDO_COTACAO(vData As Integer)
    SistGlob_NUM_MAX_ITENS_PEDIDO_COTACAO = vData
End Property

Public Property Get NUM_MAX_ITENS_PEDIDO_COTACAO() As Integer

    If SistGlob_NUM_MAX_ITENS_PEDIDO_COTACAO = 0 Then
        NUM_MAX_ITENS_PEDIDO_COTACAO = 100
    Else
        NUM_MAX_ITENS_PEDIDO_COTACAO = SistGlob_NUM_MAX_ITENS_PEDIDO_COTACAO
    End If
    
End Property

Public Property Let NUM_MAX_ITENS_REQUISICAO(vData As Integer)
    SistGlob_NUM_MAX_ITENS_REQUISICAO = vData
End Property

Public Property Get NUM_MAX_ITENS_REQUISICAO() As Integer

    If SistGlob_NUM_MAX_ITENS_REQUISICAO = 0 Then
        NUM_MAX_ITENS_REQUISICAO = 100
    Else
        NUM_MAX_ITENS_REQUISICAO = SistGlob_NUM_MAX_ITENS_REQUISICAO
    End If
    
End Property

Public Property Let NUM_MAX_ITENS_GERACAO(vData As Integer)
    SistGlob_NUM_MAX_ITENS_GERACAO = vData
End Property

Public Property Get NUM_MAX_ITENS_GERACAO() As Integer

    If SistGlob_NUM_MAX_ITENS_GERACAO = 0 Then
        NUM_MAX_ITENS_GERACAO = 100
    Else
        NUM_MAX_ITENS_GERACAO = SistGlob_NUM_MAX_ITENS_GERACAO
    End If
    
End Property

Public Property Let NUM_MAX_PRODUTOS_COTACAO(vData As Integer)
    SistGlob_NUM_MAX_PRODUTOS_COTACAO = vData
End Property

Public Property Get NUM_MAX_PRODUTOS_COTACAO() As Integer

    If SistGlob_NUM_MAX_PRODUTOS_COTACAO = 0 Then
        NUM_MAX_PRODUTOS_COTACAO = 100
    Else
        NUM_MAX_PRODUTOS_COTACAO = SistGlob_NUM_MAX_PRODUTOS_COTACAO
    End If
    
End Property

Public Property Let NUM_MAX_FORNECEDORES_COTACAO(vData As Integer)
    SistGlob_NUM_MAX_FORNECEDORES_COTACAO = vData
End Property

Public Property Get NUM_MAX_FORNECEDORES_COTACAO() As Integer

    If SistGlob_NUM_MAX_FORNECEDORES_COTACAO = 0 Then
        NUM_MAX_FORNECEDORES_COTACAO = 100
    Else
        NUM_MAX_FORNECEDORES_COTACAO = SistGlob_NUM_MAX_FORNECEDORES_COTACAO
    End If
    
End Property

Public Property Let NUM_MAX_NFS_ITEMREQ(vData As Integer)
    SistGlob_NUM_MAX_NFS_ITEMREQ = vData
End Property

Public Property Get NUM_MAX_NFS_ITEMREQ() As Integer

    If SistGlob_NUM_MAX_NFS_ITEMREQ = 0 Then
        NUM_MAX_NFS_ITEMREQ = 100
    Else
        NUM_MAX_NFS_ITEMREQ = SistGlob_NUM_MAX_NFS_ITEMREQ
    End If
    
End Property

Public Property Let NUM_MAX_PEDIDOS_ITEMREQ(vData As Integer)
    SistGlob_NUM_MAX_PEDIDOS_ITEMREQ = vData
End Property

Public Property Get NUM_MAX_PEDIDOS_ITEMREQ() As Integer

    If SistGlob_NUM_MAX_PEDIDOS_ITEMREQ = 0 Then
        NUM_MAX_PEDIDOS_ITEMREQ = 100
    Else
        NUM_MAX_PEDIDOS_ITEMREQ = SistGlob_NUM_MAX_PEDIDOS_ITEMREQ
    End If
    
End Property

Public Property Let NUM_MAX_COTACOES(vData As Integer)
    SistGlob_NUM_MAX_COTACOES = vData
End Property

Public Property Get NUM_MAX_COTACOES() As Integer

    If SistGlob_NUM_MAX_COTACOES = 0 Then
        NUM_MAX_COTACOES = 100
    Else
        NUM_MAX_COTACOES = SistGlob_NUM_MAX_COTACOES
    End If
    
End Property

Public Property Let NUM_MAX_NFS_ITEMPED(vData As Integer)
    SistGlob_NUM_MAX_NFS_ITEMPED = vData
End Property

Public Property Get NUM_MAX_NFS_ITEMPED() As Integer

    If SistGlob_NUM_MAX_NFS_ITEMPED = 0 Then
        NUM_MAX_NFS_ITEMPED = 100
    Else
        NUM_MAX_NFS_ITEMPED = SistGlob_NUM_MAX_NFS_ITEMPED
    End If
    
End Property

Public Property Let NUM_MAX_ITENS_MOV_ESTOQUE(vData As Integer)
    SistGlob_NUM_MAX_ITENS_MOV_ESTOQUE = vData
End Property

Public Property Get NUM_MAX_ITENS_MOV_ESTOQUE() As Integer

    If SistGlob_NUM_MAX_ITENS_MOV_ESTOQUE = 0 Then
        NUM_MAX_ITENS_MOV_ESTOQUE = 100
    Else
        NUM_MAX_ITENS_MOV_ESTOQUE = SistGlob_NUM_MAX_ITENS_MOV_ESTOQUE
    End If
    
End Property


Public Sub Inicializa_Tamanhos_String()

'    STRING_ENDERECO = GetPrivateProfileInt("Tamanhos", "STRING_ENDERECO", STRING_ENDERECO, NOME_ARQUIVO_ADM)
'    STRING_BAIRRO = GetPrivateProfileInt("Tamanhos", "STRING_BAIRRO", STRING_BAIRRO, NOME_ARQUIVO_ADM)
'    STRING_CIDADE = GetPrivateProfileInt("Tamanhos", "STRING_CIDADE", STRING_CIDADE, NOME_ARQUIVO_ADM)
'
'    STRING_TELEFONE = GetPrivateProfileInt("Tamanhos", "STRING_TELEFONE", STRING_TELEFONE, NOME_ARQUIVO_ADM)
'    STRING_FAX = GetPrivateProfileInt("Tamanhos", "STRING_FAX", STRING_FAX, NOME_ARQUIVO_ADM)
'    STRING_EMAIL = GetPrivateProfileInt("Tamanhos", "STRING_EMAIL", STRING_EMAIL, NOME_ARQUIVO_ADM)
'    STRING_CONTATO = GetPrivateProfileInt("Tamanhos", "STRING_CONTATO", STRING_CONTATO, NOME_ARQUIVO_ADM)
'
'    STRING_CLIENTE_RAZAO_SOCIAL = GetPrivateProfileInt("Tamanhos", "STRING_CLIENTE_RAZAO_SOCIAL", STRING_CLIENTE_RAZAO_SOCIAL, NOME_ARQUIVO_ADM)
'    STRING_CLIENTE_NOME_REDUZIDO = GetPrivateProfileInt("Tamanhos", "STRING_CLIENTE_NOME_REDUZIDO", STRING_CLIENTE_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
'    STRING_CLIENTE_OBSERVACAO = GetPrivateProfileInt("Tamanhos", "STRING_CLIENTE_OBSERVACAO", STRING_CLIENTE_OBSERVACAO, NOME_ARQUIVO_ADM)
'    STRING_FORNECEDOR_RAZAO_SOC = GetPrivateProfileInt("Tamanhos", "STRING_FORNECEDOR_RAZAO_SOC", STRING_FORNECEDOR_RAZAO_SOC, NOME_ARQUIVO_ADM)
'    STRING_FORNECEDOR_NOME_REDUZIDO = GetPrivateProfileInt("Tamanhos", "STRING_FORNECEDOR_NOME_REDUZIDO", STRING_FORNECEDOR_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
'    STRING_TRANSPORTADORA_NOME = GetPrivateProfileInt("Tamanhos", "STRING_TRANSPORTADORA_NOME", STRING_TRANSPORTADORA_NOME, NOME_ARQUIVO_ADM)
'    STRING_TRANSPORTADORA_NOME_REDUZIDO = GetPrivateProfileInt("Tamanhos", "STRING_TRANSPORTADORA_NOME_REDUZIDO", STRING_TRANSPORTADORA_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
'
'    '##########################################################################
'    'Inserido por Wagner
'    NUM_MAX_ITENS_REQUISICAO = GetPrivateProfileInt("Tamanhos", "NUM_MAX_ITENS_REQUISICAO", NUM_MAX_ITENS_REQUISICAO, NOME_ARQUIVO_ADM)
'    NUM_MAX_ITENS_PEDIDO_COTACAO = GetPrivateProfileInt("Tamanhos", "NUM_MAX_ITENS_PEDIDO_COTACAO", NUM_MAX_ITENS_PEDIDO_COTACAO, NOME_ARQUIVO_ADM)
'    NUM_MAX_ITENS_PEDIDO_COMPRAS = GetPrivateProfileInt("Tamanhos", "NUM_MAX_ITENS_PEDIDO_COMPRAS", NUM_MAX_ITENS_PEDIDO_COMPRAS, NOME_ARQUIVO_ADM)
'    NUM_MAX_ITENS_DISTRIBUICAO = GetPrivateProfileInt("Tamanhos", "NUM_MAX_ITENS_DISTRIBUICAO", NUM_MAX_ITENS_DISTRIBUICAO, NOME_ARQUIVO_ADM)
'    NUM_MAX_ITENS_GERACAO = GetPrivateProfileInt("Tamanhos", "NUM_MAX_ITENS_GERACAO", NUM_MAX_ITENS_GERACAO, NOME_ARQUIVO_ADM)
'    NUM_MAX_PRODUTOS_COTACAO = GetPrivateProfileInt("Tamanhos", "NUM_MAX_PRODUTOS_COTACAO", NUM_MAX_PRODUTOS_COTACAO, NOME_ARQUIVO_ADM)
'    NUM_MAX_FORNECEDORES_COTACAO = GetPrivateProfileInt("Tamanhos", "NUM_MAX_FORNECEDORES_COTACAO", NUM_MAX_FORNECEDORES_COTACAO, NOME_ARQUIVO_ADM)
'    NUM_MAX_NFS_ITEMREQ = GetPrivateProfileInt("Tamanhos", "NUM_MAX_NFS_ITEMREQ", NUM_MAX_NFS_ITEMREQ, NOME_ARQUIVO_ADM)
'    NUM_MAX_PEDIDOS_ITEMREQ = GetPrivateProfileInt("Tamanhos", "NUM_MAX_PEDIDOS_ITEMREQ", NUM_MAX_PEDIDOS_ITEMREQ, NOME_ARQUIVO_ADM)
'    NUM_MAX_COTACOES = GetPrivateProfileInt("Tamanhos", "NUM_MAX_COTACOES", NUM_MAX_COTACOES, NOME_ARQUIVO_ADM)
'    NUM_MAX_NFS_ITEMPED = GetPrivateProfileInt("Tamanhos", "NUM_MAX_NFS_ITEMPED", NUM_MAX_NFS_ITEMPED, NOME_ARQUIVO_ADM)
'    STRING_PRODUTO_REFERENCIA = GetPrivateProfileInt("Tamanhos", "STRING_PRODUTO_REFERENCIA", STRING_PRODUTO_REFERENCIA, "ADM100.INI")
'    STRING_PRODUTO_NOME_REDUZIDO = GetPrivateProfileInt("Tamanhos", "STRING_PRODUTO_NOME_REDUZIDO", STRING_PRODUTO_NOME_REDUZIDO, "ADM100.INI")
'    '##########################################################################
'
'    '##########################################################################
'    'Inserido por Wagner 24/03/2005
'    NUM_MAX_ITENS_MOV_ESTOQUE = GetPrivateProfileInt("Tamanhos", "NUM_MAX_ITENS_MOV_ESTOQUE", NUM_MAX_ITENS_MOV_ESTOQUE, NOME_ARQUIVO_ADM)
'    '##########################################################################
'
'    'Incluido por Jorge Specian - 11/07/2005
'    '---------------------------------------
'    STRING_ORDEM_DE_PRODUCAO = GetPrivateProfileInt("Tamanhos", "STRING_ORDEM_DE_PRODUCAO", STRING_ORDEM_DE_PRODUCAO, "ADM100.INI")
'    STRING_OPCODIGO = GetPrivateProfileInt("Tamanhos", "STRING_ORDEM_DE_PRODUCAO", STRING_OPCODIGO, "ADM100.INI")
'    STRING_LOTE_RASTREAMENTO = GetPrivateProfileInt("Tamanhos", "STRING_LOTE_RASTREAMENTO", STRING_LOTE_RASTREAMENTO, "ADM100.INI")
'    '---------------------------------------
    
    STRING_ENDERECO = Obtem_Tamanho("Tamanhos", "STRING_ENDERECO", STRING_ENDERECO, NOME_ARQUIVO_ADM)
    STRING_BAIRRO = Obtem_Tamanho("Tamanhos", "STRING_BAIRRO", STRING_BAIRRO, NOME_ARQUIVO_ADM)
    STRING_CIDADE = Obtem_Tamanho("Tamanhos", "STRING_CIDADE", STRING_CIDADE, NOME_ARQUIVO_ADM)
    STRING_TELEFONE = Obtem_Tamanho("Tamanhos", "STRING_TELEFONE", STRING_TELEFONE, NOME_ARQUIVO_ADM)
    STRING_FAX = Obtem_Tamanho("Tamanhos", "STRING_FAX", STRING_FAX, NOME_ARQUIVO_ADM)
    STRING_EMAIL = Obtem_Tamanho("Tamanhos", "STRING_EMAIL", STRING_EMAIL, NOME_ARQUIVO_ADM)
    STRING_CONTATO = Obtem_Tamanho("Tamanhos", "STRING_CONTATO", STRING_CONTATO, NOME_ARQUIVO_ADM)
    STRING_CLIENTE_RAZAO_SOCIAL = Obtem_Tamanho("Tamanhos", "STRING_CLIENTE_RAZAO_SOCIAL", STRING_CLIENTE_RAZAO_SOCIAL, NOME_ARQUIVO_ADM)
    STRING_CLIENTE_NOME_REDUZIDO = Obtem_Tamanho("Tamanhos", "STRING_CLIENTE_NOME_REDUZIDO", STRING_CLIENTE_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
    STRING_CLIENTE_OBSERVACAO = Obtem_Tamanho("Tamanhos", "STRING_CLIENTE_OBSERVACAO", STRING_CLIENTE_OBSERVACAO, NOME_ARQUIVO_ADM)
    STRING_FORNECEDOR_RAZAO_SOC = Obtem_Tamanho("Tamanhos", "STRING_FORNECEDOR_RAZAO_SOC", STRING_FORNECEDOR_RAZAO_SOC, NOME_ARQUIVO_ADM)
    STRING_FORNECEDOR_NOME_REDUZIDO = Obtem_Tamanho("Tamanhos", "STRING_FORNECEDOR_NOME_REDUZIDO", STRING_FORNECEDOR_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
    STRING_TRANSPORTADORA_NOME = Obtem_Tamanho("Tamanhos", "STRING_TRANSPORTADORA_NOME", STRING_TRANSPORTADORA_NOME, NOME_ARQUIVO_ADM)
    STRING_TRANSPORTADORA_NOME_REDUZIDO = Obtem_Tamanho("Tamanhos", "STRING_TRANSPORTADORA_NOME_REDUZIDO", STRING_TRANSPORTADORA_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
    NUM_MAX_ITENS_REQUISICAO = Obtem_Tamanho("Tamanhos", "NUM_MAX_ITENS_REQUISICAO", NUM_MAX_ITENS_REQUISICAO, NOME_ARQUIVO_ADM)
    NUM_MAX_ITENS_PEDIDO_COTACAO = Obtem_Tamanho("Tamanhos", "NUM_MAX_ITENS_PEDIDO_COTACAO", NUM_MAX_ITENS_PEDIDO_COTACAO, NOME_ARQUIVO_ADM)
    NUM_MAX_ITENS_PEDIDO_COMPRAS = Obtem_Tamanho("Tamanhos", "NUM_MAX_ITENS_PEDIDO_COMPRAS", NUM_MAX_ITENS_PEDIDO_COMPRAS, NOME_ARQUIVO_ADM)
    NUM_MAX_ITENS_DISTRIBUICAO = Obtem_Tamanho("Tamanhos", "NUM_MAX_ITENS_DISTRIBUICAO", NUM_MAX_ITENS_DISTRIBUICAO, NOME_ARQUIVO_ADM)
    NUM_MAX_ITENS_GERACAO = Obtem_Tamanho("Tamanhos", "NUM_MAX_ITENS_GERACAO", NUM_MAX_ITENS_GERACAO, NOME_ARQUIVO_ADM)
    NUM_MAX_PRODUTOS_COTACAO = Obtem_Tamanho("Tamanhos", "NUM_MAX_PRODUTOS_COTACAO", NUM_MAX_PRODUTOS_COTACAO, NOME_ARQUIVO_ADM)
    NUM_MAX_FORNECEDORES_COTACAO = Obtem_Tamanho("Tamanhos", "NUM_MAX_FORNECEDORES_COTACAO", NUM_MAX_FORNECEDORES_COTACAO, NOME_ARQUIVO_ADM)
    NUM_MAX_NFS_ITEMREQ = Obtem_Tamanho("Tamanhos", "NUM_MAX_NFS_ITEMREQ", NUM_MAX_NFS_ITEMREQ, NOME_ARQUIVO_ADM)
    NUM_MAX_PEDIDOS_ITEMREQ = Obtem_Tamanho("Tamanhos", "NUM_MAX_PEDIDOS_ITEMREQ", NUM_MAX_PEDIDOS_ITEMREQ, NOME_ARQUIVO_ADM)
    NUM_MAX_COTACOES = Obtem_Tamanho("Tamanhos", "NUM_MAX_COTACOES", NUM_MAX_COTACOES, NOME_ARQUIVO_ADM)
    NUM_MAX_NFS_ITEMPED = Obtem_Tamanho("Tamanhos", "NUM_MAX_NFS_ITEMPED", NUM_MAX_NFS_ITEMPED, NOME_ARQUIVO_ADM)
    STRING_PRODUTO_REFERENCIA = Obtem_Tamanho("Tamanhos", "STRING_PRODUTO_REFERENCIA", STRING_PRODUTO_REFERENCIA, "ADM100.INI")
    STRING_PRODUTO_NOME_REDUZIDO = Obtem_Tamanho("Tamanhos", "STRING_PRODUTO_NOME_REDUZIDO", STRING_PRODUTO_NOME_REDUZIDO, "ADM100.INI")
    STRING_PRODUTO_MODELO = Obtem_Tamanho("Tamanhos", "STRING_PRODUTO_MODELO", STRING_PRODUTO_MODELO, "ADM100.INI")
    NUM_MAX_ITENS_MOV_ESTOQUE = Obtem_Tamanho("Tamanhos", "NUM_MAX_ITENS_MOV_ESTOQUE", NUM_MAX_ITENS_MOV_ESTOQUE, NOME_ARQUIVO_ADM)
    STRING_ORDEM_DE_PRODUCAO = Obtem_Tamanho("Tamanhos", "STRING_ORDEM_DE_PRODUCAO", STRING_ORDEM_DE_PRODUCAO, "ADM100.INI")
    STRING_OPCODIGO = Obtem_Tamanho("Tamanhos", "STRING_ORDEM_DE_PRODUCAO", STRING_OPCODIGO, "ADM100.INI")
    STRING_LOTE_RASTREAMENTO = Obtem_Tamanho("Tamanhos", "STRING_LOTE_RASTREAMENTO", STRING_LOTE_RASTREAMENTO, "ADM100.INI")
    STRING_PRODUTO_DESCRICAO_TELA = Obtem_Tamanho("Tamanhos", "STRING_PRODUTO_DESCRICAO_TELA", STRING_PRODUTO_DESCRICAO_TELA, "ADM100.INI")
    
End Sub

Private Function Obtem_Tamanho(ByVal sTabela As String, ByVal sStringCampo As String, ByVal iValorPadrao As Integer, ByVal sNomeArqIni As String) As Integer

Dim lErro As Long
Dim iValor As Integer
Dim Y As New ADM

On Error GoTo Erro_Obtem_Tamanho

    If Y.gsNomePrinc = "SGEECF" Then
    
        iValor = GetPrivateProfileInt(sTabela, sStringCampo, iValorPadrao, sNomeArqIni)
        
    Else
    
        lErro = Y.CF("Tamanhos_Le", sStringCampo, iValor)
        If lErro <> SUCESSO Then
            iValor = GetPrivateProfileInt(sTabela, sStringCampo, iValorPadrao, sNomeArqIni)
            If iValor <> 0 Then
                Call Y.CF("Tamanhos_Grava", sStringCampo, iValor)
            End If
        End If
    
    End If
    
    Obtem_Tamanho = iValor
    
    Exit Function
    
Erro_Obtem_Tamanho:

    Obtem_Tamanho = GetPrivateProfileInt(sTabela, sStringCampo, iValorPadrao, sNomeArqIni)
    
    Exit Function

End Function

Public Property Let STRING_TELEFONE(vData As Integer)
    SistGlob_STRING_TELEFONE = vData
End Property

Public Property Get STRING_TELEFONE() As Integer

    If SistGlob_STRING_TELEFONE = 0 Then
        STRING_TELEFONE = 18
    Else
        STRING_TELEFONE = SistGlob_STRING_TELEFONE
    End If
    
End Property

Public Property Let STRING_FAX(vData As Integer)
    SistGlob_STRING_FAX = vData
End Property

Public Property Get STRING_FAX() As Integer

    If SistGlob_STRING_FAX = 0 Then
        STRING_FAX = 18
    Else
        STRING_FAX = SistGlob_STRING_FAX
    End If
    
End Property

Public Property Let STRING_EMAIL(vData As Integer)
    SistGlob_STRING_EMAIL = vData
End Property

Public Property Get STRING_EMAIL() As Integer

    If SistGlob_STRING_EMAIL = 0 Then
        STRING_EMAIL = 50
    Else
        STRING_EMAIL = SistGlob_STRING_EMAIL
    End If
    
End Property

Public Property Let STRING_CONTATO(vData As Integer)
    SistGlob_STRING_CONTATO = vData
End Property

Public Property Get STRING_CONTATO() As Integer

    If SistGlob_STRING_CONTATO = 0 Then
        STRING_CONTATO = 50
    Else
        STRING_CONTATO = SistGlob_STRING_CONTATO
    End If
    
End Property

Public Property Let STRING_FORNECEDOR_RAZAO_SOC(vData As Integer)
    SistGlob_STRING_FORNECEDOR_RAZAO_SOC = vData
End Property

Public Property Get STRING_FORNECEDOR_RAZAO_SOC() As Integer

    If SistGlob_STRING_FORNECEDOR_RAZAO_SOC = 0 Then
        STRING_FORNECEDOR_RAZAO_SOC = 40
    Else
        STRING_FORNECEDOR_RAZAO_SOC = SistGlob_STRING_FORNECEDOR_RAZAO_SOC
    End If
    
End Property

Public Property Let STRING_FORNECEDOR_NOME_REDUZIDO(vData As Integer)
    SistGlob_STRING_FORNECEDOR_NOME_REDUZIDO = vData
End Property

Public Property Get STRING_FORNECEDOR_NOME_REDUZIDO() As Integer

    If SistGlob_STRING_FORNECEDOR_NOME_REDUZIDO = 0 Then
        STRING_FORNECEDOR_NOME_REDUZIDO = 20
    Else
        STRING_FORNECEDOR_NOME_REDUZIDO = SistGlob_STRING_FORNECEDOR_NOME_REDUZIDO
    End If
    
End Property

Public Property Let STRING_ORDEM_DE_PRODUCAO(vData As Integer)
    SistGlob_STRING_ORDEM_DE_PRODUCAO = vData
    End Property

Public Property Get STRING_ORDEM_DE_PRODUCAO() As Integer

    If SistGlob_STRING_ORDEM_DE_PRODUCAO = 0 Then
        STRING_ORDEM_DE_PRODUCAO = 9
    Else
        STRING_ORDEM_DE_PRODUCAO = SistGlob_STRING_ORDEM_DE_PRODUCAO
    End If
    
End Property

Public Property Let STRING_OPCODIGO(vData As Integer)
    SistGlob_STRING_ORDEM_DE_PRODUCAO = vData
End Property

Public Property Get STRING_OPCODIGO() As Integer

    If SistGlob_STRING_ORDEM_DE_PRODUCAO = 0 Then
        STRING_OPCODIGO = 9
    Else
        STRING_OPCODIGO = SistGlob_STRING_ORDEM_DE_PRODUCAO
    End If
    
End Property

Public Property Let STRING_LOTE_RASTREAMENTO(vData As Integer)
    SistGlob_STRING_LOTE_RASTREAMENTO = vData
End Property

Public Property Get STRING_LOTE_RASTREAMENTO() As Integer

    If SistGlob_STRING_LOTE_RASTREAMENTO = 0 Then
        STRING_LOTE_RASTREAMENTO = 9
    Else
        STRING_LOTE_RASTREAMENTO = SistGlob_STRING_LOTE_RASTREAMENTO
    End If
    
End Property

Public Property Let STRING_PRODUTO_REFERENCIA(vData As Integer)
    SistGlob_STRING_PRODUTO_REFERENCIA = vData
End Property

Public Property Get STRING_PRODUTO_REFERENCIA() As Integer

    If SistGlob_STRING_PRODUTO_REFERENCIA = 0 Then
        STRING_PRODUTO_REFERENCIA = 20
    Else
        STRING_PRODUTO_REFERENCIA = SistGlob_STRING_PRODUTO_REFERENCIA
    End If
    
End Property

Public Property Let STRING_PRODUTO_NOME_REDUZIDO(vData As Integer)
    SistGlob_STRING_PRODUTO_NOME_REDUZIDO = vData
End Property

Public Property Get STRING_PRODUTO_NOME_REDUZIDO() As Integer

    If SistGlob_STRING_PRODUTO_NOME_REDUZIDO = 0 Then
        STRING_PRODUTO_NOME_REDUZIDO = 20
    Else
        STRING_PRODUTO_NOME_REDUZIDO = SistGlob_STRING_PRODUTO_NOME_REDUZIDO
    End If
    
End Property

Public Property Let STRING_PRODUTO_MODELO(vData As Integer)
    SistGlob_STRING_PRODUTO_MODELO = vData
End Property

Public Property Get STRING_PRODUTO_MODELO() As Integer

    If SistGlob_STRING_PRODUTO_MODELO = 0 Then
        STRING_PRODUTO_MODELO = 40
    Else
        STRING_PRODUTO_MODELO = SistGlob_STRING_PRODUTO_MODELO
    End If
    
End Property

Public Property Let iForcaSistemaIntegrado(vData As Integer)
    SistGlob_iForcaSistemaIntegrado = vData
End Property

Public Property Get iForcaSistemaIntegrado() As Integer
    iForcaSistemaIntegrado = SistGlob_iForcaSistemaIntegrado
End Property

Public Property Let iSistemaIntegradoForcado(vData As Integer)
    SistGlob_iSistemaIntegradoForcado = vData
End Property

Public Property Get iSistemaIntegradoForcado() As Integer
    iSistemaIntegradoForcado = SistGlob_iSistemaIntegradoForcado
End Property
