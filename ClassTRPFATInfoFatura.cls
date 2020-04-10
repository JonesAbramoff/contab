VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPFATInfoFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolInfoVouchers As New Collection
Private mvarcolInfoVouchersAG As New Collection
Private mvarcolFatAporte As New Collection
Private mvarlCliente As Long
Private mvarlFatura As Long
Private mvardtDataEmissao As Date
Private mvardtDataVencimento As Date
Private mvardValor As Double
Private mvariFilialEmpresa As Integer
Private mvariCondPagto As Integer
Private mvariMarcado As Integer
Private mvardValorComissao As Double
Private mvardValorAporte As Double
Private mvardValorCreditos As Double
Private mvardValorDebitos As Double
Private mvariLinha As Integer
Private mvariLinhaFatFilial As Integer
Private mvariLinhaFatFilialCliente As Integer
Private mvariLinhaFatCliente As Integer
Private mvariLinhaFatEnderecoCli As Integer
Private mvariLinhaFatEnderecoEmp As Integer
Private mvarlNumeroFaturaReal As Long
Private mvariParcelaAporte As Integer
Private mvariCartao As Integer
Private mvardValorTarifa As Double
Private mvardValorDeducoes As Double
Private mvarsSiglaDocumento As String
Private mvariTipoDoc As Integer
Private mvarcolFatAporteCred As New Collection
Private mvardValorAporteCred As Double

Public Property Let dValorAporteCred(ByVal vData As Double)
    mvardValorAporteCred = vData
End Property

Public Property Get dValorAporteCred() As Double
    dValorAporteCred = mvardValorAporteCred
End Property

Public Property Set colFatAporteCred(ByVal vData As Collection)
    Set mvarcolFatAporteCred = vData
End Property

Public Property Get colFatAporteCred() As Collection
    Set colFatAporteCred = mvarcolFatAporteCred
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let sSiglaDocumento(ByVal vData As String)
    mvarsSiglaDocumento = vData
End Property

Public Property Get sSiglaDocumento() As String
    sSiglaDocumento = mvarsSiglaDocumento
End Property

Public Property Let dValorDebitos(ByVal vData As Double)
    mvardValorDebitos = vData
End Property

Public Property Get dValorDebitos() As Double
    dValorDebitos = mvardValorDebitos
End Property

Public Property Let dValorCreditos(ByVal vData As Double)
    mvardValorCreditos = vData
End Property

Public Property Get dValorCreditos() As Double
    dValorCreditos = mvardValorCreditos
End Property

Public Property Let dValorDeducoes(ByVal vData As Double)
    mvardValorDeducoes = vData
End Property

Public Property Get dValorDeducoes() As Double
    dValorDeducoes = mvardValorDeducoes
End Property

Public Property Let dValorTarifa(ByVal vData As Double)
    mvardValorTarifa = vData
End Property

Public Property Get dValorTarifa() As Double
    dValorTarifa = mvardValorTarifa
End Property

Public Property Let iCartao(ByVal vData As Integer)
    mvariCartao = vData
End Property

Public Property Get iCartao() As Integer
    iCartao = mvariCartao
End Property

Public Property Let iParcelaAporte(ByVal vData As Integer)
    mvariParcelaAporte = vData
End Property

Public Property Get iParcelaAporte() As Integer
    iParcelaAporte = mvariParcelaAporte
End Property

Public Property Let lNumeroFaturaReal(ByVal vData As Long)
    mvarlNumeroFaturaReal = vData
End Property

Public Property Get lNumeroFaturaReal() As Long
    lNumeroFaturaReal = mvarlNumeroFaturaReal
End Property

Public Property Let iLinhaFatEnderecoEmp(ByVal vData As Integer)
    mvariLinhaFatEnderecoEmp = vData
End Property

Public Property Get iLinhaFatEnderecoEmp() As Integer
    iLinhaFatEnderecoEmp = mvariLinhaFatEnderecoEmp
End Property

Public Property Let iLinhaFatEnderecoCli(ByVal vData As Integer)
    mvariLinhaFatEnderecoCli = vData
End Property

Public Property Get iLinhaFatEnderecoCli() As Integer
    iLinhaFatEnderecoCli = mvariLinhaFatEnderecoCli
End Property

Public Property Let iLinhaFatCliente(ByVal vData As Integer)
    mvariLinhaFatCliente = vData
End Property

Public Property Get iLinhaFatCliente() As Integer
    iLinhaFatCliente = mvariLinhaFatCliente
End Property

Public Property Let iLinhaFatFilialCliente(ByVal vData As Integer)
    mvariLinhaFatFilialCliente = vData
End Property

Public Property Get iLinhaFatFilialCliente() As Integer
    iLinhaFatFilialCliente = mvariLinhaFatFilialCliente
End Property

Public Property Let iLinhaFatFilial(ByVal vData As Integer)
    mvariLinhaFatFilial = vData
End Property

Public Property Get iLinhaFatFilial() As Integer
    iLinhaFatFilial = mvariLinhaFatFilial
End Property

Public Property Let iLinha(ByVal vData As Integer)
    mvariLinha = vData
End Property

Public Property Get iLinha() As Integer
    iLinha = mvariLinha
End Property

Public Property Let dValorAporte(ByVal vData As Double)
    mvardValorAporte = vData
End Property

Public Property Get dValorAporte() As Double
    dValorAporte = mvardValorAporte
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let iMarcado(ByVal vData As Integer)
    mvariMarcado = vData
End Property

Public Property Get iMarcado() As Integer
    iMarcado = mvariMarcado
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lFatura(ByVal vData As Long)
    mvarlFatura = vData
End Property

Public Property Get lFatura() As Long
    lFatura = mvarlFatura
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Set colFatAporte(ByVal vData As Collection)
    Set mvarcolFatAporte = vData
End Property

Public Property Get colFatAporte() As Collection
    Set colFatAporte = mvarcolFatAporte
End Property

Public Property Set colInfoVouchers(ByVal vData As Collection)
    Set mvarcolInfoVouchers = vData
End Property

Public Property Get colInfoVouchers() As Collection
    Set colInfoVouchers = mvarcolInfoVouchers
End Property


Public Property Set colInfoVouchersAG(ByVal vData As Collection)
    Set mvarcolInfoVouchersAG = vData
End Property

Public Property Get colInfoVouchersAG() As Collection
    Set colInfoVouchersAG = mvarcolInfoVouchersAG
End Property

Public Sub Agrupa_Itens()

Dim colAux As New Collection
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoVoucherAux As ClassTRPFATInfoVoucher
Dim bAchou As Boolean

    'Agrupa dados do mesmo voucher\Tipo
    For Each objInfoVoucherAux In mvarcolInfoVouchers
        If objInfoVoucherAux.iMarcado = MARCADO Then
            If objInfoVoucherAux.sTipoDoc <> "VOU" And objInfoVoucherAux.sTipoDoc <> "NVL" And objInfoVoucherAux.sTipoDoc <> "OCR" Then
                bAchou = False
                For Each objInfoVoucher In colAux
                    'Se é o mesmo documento do mesmo voucher agrupa os valores
                    If objInfoVoucherAux.sTipoDoc = objInfoVoucher.sTipoDoc And objInfoVoucherAux.sTipoVou = objInfoVoucher.sTipoVou And objInfoVoucherAux.sSerie = objInfoVoucher.sSerie And objInfoVoucherAux.lNumVou = objInfoVoucher.lNumVou Then
                        objInfoVoucher.dValor = objInfoVoucher.dValor + objInfoVoucherAux.dValor
                        objInfoVoucher.dValorBruto = objInfoVoucher.dValorBruto + objInfoVoucherAux.dValorBruto
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    Set objInfoVoucher = New ClassTRPFATInfoVoucher
                    objInfoVoucher.iMarcado = objInfoVoucherAux.iMarcado
                    objInfoVoucher.dCambio = objInfoVoucherAux.dCambio
                    objInfoVoucher.dtDataEmissao = objInfoVoucherAux.dtDataEmissao
                    objInfoVoucher.dValor = objInfoVoucherAux.dValor
                    objInfoVoucher.dValorBruto = objInfoVoucherAux.dValorBruto
                    objInfoVoucher.dValorCambio = objInfoVoucherAux.dValorCambio
                    objInfoVoucher.dValorComissao = objInfoVoucherAux.dValorComissao
                    objInfoVoucher.dValorVou = objInfoVoucherAux.dValorVou
                    objInfoVoucher.iCartao = objInfoVoucherAux.iCartao
                    objInfoVoucher.iCondPagto = objInfoVoucherAux.iCondPagto
                    objInfoVoucher.iFilialEmpresa = objInfoVoucherAux.iFilialEmpresa
                    objInfoVoucher.iFilialFat = objInfoVoucherAux.iFilialFat
                    objInfoVoucher.iLinha = objInfoVoucherAux.iLinha
                    objInfoVoucher.iPax = objInfoVoucherAux.iPax
                    objInfoVoucher.iTipoDoc = objInfoVoucherAux.iTipoDoc
                    objInfoVoucher.iTipoGeracao = objInfoVoucherAux.iTipoGeracao
                    objInfoVoucher.lCliente = objInfoVoucherAux.lCliente
                    objInfoVoucher.lCliVou = objInfoVoucherAux.lCliVou
                    objInfoVoucher.lEmissor = objInfoVoucherAux.lEmissor
                    objInfoVoucher.lFatura = objInfoVoucherAux.lFatura
                    objInfoVoucher.lNumero = objInfoVoucherAux.lNumero
                    objInfoVoucher.lNumIntDoc = objInfoVoucherAux.lNumIntDoc
                    objInfoVoucher.lNumVou = objInfoVoucherAux.lNumVou
                    objInfoVoucher.sControle = objInfoVoucherAux.sControle
                    objInfoVoucher.sEmissor = objInfoVoucherAux.sEmissor
                    objInfoVoucher.iMoeda = objInfoVoucherAux.iMoeda
                    objInfoVoucher.sNomeCliVou = objInfoVoucherAux.sNomeCliVou
                    objInfoVoucher.sNome = objInfoVoucherAux.sNome
                    objInfoVoucher.sPassageiroNome = objInfoVoucherAux.sPassageiroNome
                    objInfoVoucher.sPassageiroSobreNome = objInfoVoucherAux.sPassageiroSobreNome
                    objInfoVoucher.sProduto = objInfoVoucherAux.sProduto
                    objInfoVoucher.sSerie = objInfoVoucherAux.sSerie
                    objInfoVoucher.sTipoDoc = objInfoVoucherAux.sTipoDoc
                    objInfoVoucher.sTipoVou = objInfoVoucherAux.sTipoVou
                    objInfoVoucher.sTitular = objInfoVoucherAux.sTitular
                    objInfoVoucher.sTitularVou = objInfoVoucherAux.sTitularVou
                    
                    colAux.Add objInfoVoucher
                End If
            Else
                colAux.Add objInfoVoucherAux
            End If
        End If
    Next

    Set mvarcolInfoVouchersAG = colAux
End Sub
