VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoDocInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'**** ATENCAO: QUANDO INCLUIR NOVO CAMPO ATUALIZAR A SUB COPIAR ******'

Private mvarsSigla As String
Private mvarsDescricao As String
Private mvariTipoMovtoEstoque As Integer
Private mvarsNaturezaOperacaoPadrao As String
Private mvariInfoContabilizacao As Integer
Private mvarsTipoDocCPR As String
Private mvariCodigo As Integer
Private mvarsNomeReduzido As String
Private mvarsNomeTelaNFiscal As String
Private mvarsTituloTelaNFiscal As String
Private mvariFaturavel As Integer
Private mvariComissao As Integer
Private mvariEmitente As Integer
Private mvariDestinatario As Integer
Private mvariComplementar As Integer
Private mvariTipo As Integer
Private mvariOrigem As Integer
Private mvariTipoMovtoEstoque2 As Integer
Private mvariPadrao As Integer
Private mvariFaturamento As Integer
Private mvariTipoOperacaoTrib As Integer
Private mvarsNatOpExtPadrao As String
Private mvariTipoMovtoEstoqueBenef As Integer
Private mvarsSiglaNFOriginal As String
Private mvariModeloArqICMS As Integer
Private mvariNFFatura As Integer
Private mvariTipoMovtoEstoqueBenef2 As Integer
Private mvariTipoMovtoEstoqueConsig As Integer
Private mvariTipoMovtoEstoqueConsig2 As Integer
Private mvariSubTipoContabil As Integer
Private mvariRastreavel As Integer
Private mvariCompras As Integer
Private mvariEscaninhoRastro As Integer
Private mvarbExibeGridAlocacao As Boolean

'nfe 3.10
Public iModDocFis As Integer
Public iModDocFisE As Integer
Public iFinalidadeNFe As Integer
Public iIndConsumidorFinal As Integer
Public iIndPresenca As Integer

Public Property Get bExibeGridAlocacao() As Boolean

    Select Case mvarsNomeTelaNFiscal
    
        Case NOME_TELA_GERACAONFISCAL, NOME_TELA_PEDIDOVENDA, NOME_TELA_NFISCAL, NOME_TELA_NFISCALFATURA, NOME_TELA_NFISCALPEDIDO, NOME_TELA_NFISCALFATURAPEDIDO, NOME_TELA_NFISCALREM, "NFiscalRemPedido"
            bExibeGridAlocacao = True
        
        Case NOME_TELA_CONHECIMENTOFRETEFATURA
            bExibeGridAlocacao = False
        
        Case Else
            bExibeGridAlocacao = False
    
    End Select

End Property

Public Property Let iCompras(ByVal vData As Integer)
    mvariCompras = vData
End Property

Public Property Get iCompras() As Integer
    iCompras = mvariCompras
End Property

Public Property Let iRastreavel(ByVal vData As Integer)
    mvariRastreavel = vData
End Property

Public Property Get iRastreavel() As Integer
    iRastreavel = mvariRastreavel
End Property

Public Property Let iSubTipoContabil(ByVal vData As Integer)
    mvariSubTipoContabil = vData
End Property

Public Property Get iSubTipoContabil() As Integer
    iSubTipoContabil = mvariSubTipoContabil
End Property

Public Property Let iTipoMovtoEstoqueConsig2(ByVal vData As Integer)
    mvariTipoMovtoEstoqueConsig2 = vData
End Property

Public Property Get iTipoMovtoEstoqueConsig2() As Integer
    iTipoMovtoEstoqueConsig2 = mvariTipoMovtoEstoqueConsig2
End Property

Public Property Let iTipoMovtoEstoqueConsig(ByVal vData As Integer)
    mvariTipoMovtoEstoqueConsig = vData
End Property

Public Property Get iTipoMovtoEstoqueConsig() As Integer
    iTipoMovtoEstoqueConsig = mvariTipoMovtoEstoqueConsig
End Property

Public Property Let iTipoMovtoEstoqueBenef2(ByVal vData As Integer)
    mvariTipoMovtoEstoqueBenef2 = vData
End Property


Public Property Get iTipoMovtoEstoqueBenef2() As Integer
    iTipoMovtoEstoqueBenef2 = mvariTipoMovtoEstoqueBenef2
End Property




Public Property Let iNFFatura(ByVal vData As Integer)
    mvariNFFatura = vData
End Property

Public Property Get iNFFatura() As Integer
    iNFFatura = mvariNFFatura
End Property

Public Property Let iModeloArqICMS(ByVal vData As Integer)
    mvariModeloArqICMS = vData
End Property

Public Property Get iModeloArqICMS() As Integer
    iModeloArqICMS = mvariModeloArqICMS
End Property

Public Property Let sSiglaNFOriginal(ByVal vData As String)
    mvarsSiglaNFOriginal = vData
End Property

Public Property Get sSiglaNFOriginal() As String
    sSiglaNFOriginal = mvarsSiglaNFOriginal
End Property

Public Property Let iTipoMovtoEstoqueBenef(ByVal vData As Integer)
    mvariTipoMovtoEstoqueBenef = vData
End Property

Public Property Get iTipoMovtoEstoqueBenef() As Integer
    iTipoMovtoEstoqueBenef = mvariTipoMovtoEstoqueBenef
End Property

Public Property Let sNatOpExtPadrao(ByVal vData As String)
    mvarsNatOpExtPadrao = vData
End Property

Public Property Get sNatOpExtPadrao() As String
    sNatOpExtPadrao = mvarsNatOpExtPadrao
End Property

Public Property Let iTipoOperacaoTrib(ByVal vData As Integer)
    mvariTipoOperacaoTrib = vData
End Property

Public Property Get iTipoOperacaoTrib() As Integer
    iTipoOperacaoTrib = mvariTipoOperacaoTrib
End Property

Public Property Let iFaturamento(ByVal vData As Integer)
    mvariFaturamento = vData
End Property

Public Property Get iFaturamento() As Integer
    iFaturamento = mvariFaturamento
End Property

Public Property Let iPadrao(ByVal vData As Integer)
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
End Property

Public Property Let iTipoMovtoEstoque2(ByVal vData As Integer)
    mvariTipoMovtoEstoque2 = vData
End Property

Public Property Get iTipoMovtoEstoque2() As Integer
    iTipoMovtoEstoque2 = mvariTipoMovtoEstoque2
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iComplementar(ByVal vData As Integer)
    mvariComplementar = vData
End Property

Public Property Get iComplementar() As Integer
    iComplementar = mvariComplementar
End Property

Public Property Let iDestinatario(ByVal vData As Integer)
    mvariDestinatario = vData
End Property

Public Property Get iDestinatario() As Integer
    iDestinatario = mvariDestinatario
End Property

Public Property Let iEmitente(ByVal vData As Integer)
    mvariEmitente = vData
End Property

Public Property Get iEmitente() As Integer
    iEmitente = mvariEmitente
End Property

Public Property Let iComissao(ByVal vData As Integer)
    mvariComissao = vData
End Property

Public Property Get iComissao() As Integer
    iComissao = mvariComissao
End Property

Public Property Let iFaturavel(ByVal vData As Integer)
    mvariFaturavel = vData
End Property

Public Property Get iFaturavel() As Integer
    iFaturavel = mvariFaturavel
End Property

Public Property Let sTituloTelaNFiscal(ByVal vData As String)
    mvarsTituloTelaNFiscal = vData
End Property

Public Property Get sTituloTelaNFiscal() As String
    sTituloTelaNFiscal = mvarsTituloTelaNFiscal
End Property

Public Property Let sNomeTelaNFiscal(ByVal vData As String)
    mvarsNomeTelaNFiscal = vData
End Property

Public Property Get sNomeTelaNFiscal() As String
    sNomeTelaNFiscal = mvarsNomeTelaNFiscal
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sTipoDocCPR(ByVal vData As String)
    mvarsTipoDocCPR = vData
End Property

Public Property Get sTipoDocCPR() As String
    sTipoDocCPR = mvarsTipoDocCPR
End Property

Public Property Let iInfoContabilizacao(ByVal vData As Integer)
    mvariInfoContabilizacao = vData
End Property

Public Property Get iInfoContabilizacao() As Integer
    iInfoContabilizacao = mvariInfoContabilizacao
End Property

Public Property Let sNaturezaOperacaoPadrao(ByVal vData As String)
    mvarsNaturezaOperacaoPadrao = vData
End Property

Public Property Get sNaturezaOperacaoPadrao() As String
    sNaturezaOperacaoPadrao = mvarsNaturezaOperacaoPadrao
End Property

Public Property Let iTipoMovtoEstoque(ByVal vData As Integer)
    mvariTipoMovtoEstoque = vData
End Property

Public Property Get iTipoMovtoEstoque() As Integer
    iTipoMovtoEstoque = mvariTipoMovtoEstoque
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let iEscaninhoRastro(ByVal vData As Integer)
    mvariEscaninhoRastro = vData
End Property

Public Property Get iEscaninhoRastro() As Integer
    iEscaninhoRastro = mvariEscaninhoRastro
End Property

Public Sub Copiar(objTipoDocInfo As ClassTipoDocInfo)
'copia os dados de objTributacaoTipo p/"Me"

    With objTipoDocInfo
        mvarbExibeGridAlocacao = .bExibeGridAlocacao
        mvarsSigla = .sSigla
        mvarsDescricao = .sDescricao
        mvariTipoMovtoEstoque = .iTipoMovtoEstoque
        mvarsNaturezaOperacaoPadrao = .sNaturezaOperacaoPadrao
        mvariInfoContabilizacao = .iInfoContabilizacao
        mvarsTipoDocCPR = .sTipoDocCPR
        mvariCodigo = .iCodigo
        mvarsNomeReduzido = .sNomeReduzido
        mvarsNomeTelaNFiscal = .sNomeTelaNFiscal
        mvarsTituloTelaNFiscal = .sTituloTelaNFiscal
        mvariFaturavel = .iFaturavel
        mvariComissao = .iComissao
        mvariEmitente = .iEmitente
        mvariDestinatario = .iDestinatario
        mvariComplementar = .iComplementar
        mvariTipo = .iTipo
        mvariOrigem = .iOrigem
        mvariTipoMovtoEstoque2 = .iTipoMovtoEstoque2
        mvariPadrao = .iPadrao
        mvariFaturamento = .iFaturamento
        mvariTipoOperacaoTrib = .iTipoOperacaoTrib
        mvarsNatOpExtPadrao = .sNatOpExtPadrao
        mvariTipoMovtoEstoqueBenef = .iTipoMovtoEstoqueBenef
        mvarsSiglaNFOriginal = .sSiglaNFOriginal
        mvariModeloArqICMS = .iModeloArqICMS
        mvariNFFatura = .iNFFatura
        mvariTipoMovtoEstoqueBenef2 = .iTipoMovtoEstoqueBenef2
        mvariTipoMovtoEstoqueConsig = .iTipoMovtoEstoqueConsig
        mvariTipoMovtoEstoqueConsig2 = .iTipoMovtoEstoqueConsig2
        mvariSubTipoContabil = .iSubTipoContabil
        mvariRastreavel = .iRastreavel
        mvariCompras = .iCompras
        mvariEscaninhoRastro = .iEscaninhoRastro
        
        'nfe 3.10
        iModDocFis = .iModDocFis
        iModDocFisE = .iModDocFisE
        iFinalidadeNFe = .iFinalidadeNFe
        iIndConsumidorFinal = .iIndConsumidorFinal
        iIndPresenca = .iIndPresenca

    End With
    
End Sub

Public Property Get iTipoTrib() As Integer
    
    'é uma entrada
    If mvariOrigem <> DOCINFO_ORIGEM_EMPRESA Then
    
        If mvariOrigem = DOCINFO_ORIGEM_CLIENTE Then
            iTipoTrib = TRIB_ENTRADA_CLI
        Else
            iTipoTrib = TRIB_ENTRADA_FORN
        End If
    
    Else 'é uma saida
    
        If mvariDestinatario = DESTINATARIO_CLIENTE Then
            iTipoTrib = TRIB_SAIDA_CLI
        Else
            iTipoTrib = TRIB_SAIDA_FORN
        End If
    
    End If
    
End Property
