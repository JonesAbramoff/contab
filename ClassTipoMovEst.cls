VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoMovEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvarsEntradaOuSaida As String
Private mvariInativo As Integer
Private mvariAtualizaConsumo As Integer
Private mvariAtualizaVendaConsig3 As Integer
Private mvariAtualizaVenda As Integer
Private mvariEditavel As Integer
Private mvariValidoMovInt As Integer
Private mvariAtualizaCusto As Integer
Private mvariAtualizaConsig As Integer
Private mvariAtualizaDemo As Integer
Private mvariAtualizaConserto As Integer
Private mvariProdutoDeTerc As Integer
Private mvariAtualizaOutras As Integer
Private mvariAtualizaIndOutras As Integer
Private mvariAtualizaNossaDisp As Integer
Private mvariAtualizaDefeituosa As Integer
Private mvariAtualizaRecebIndisp As Integer
Private mvariAtualizaMovEstoque As Integer
Private mvariAtualizaReserva As Integer
Private mvariTransferencia As Integer
Private mvariAtualizaCMRProd As Integer
Private mvariCustoInformado As Integer
Private mvariValidaDataUltMov As Integer
Private mvariAtualizaCompra As Integer
Private mvariAtualizaBenef As Integer
Private mvariAtualizaOP As Integer
Private mvarsDescrNumIntDocOrigem As String
Private mvariTipoNumIntDocOrigem As Integer
Private mvariKardex As Integer
Private mvarsSigla As String
Private mvarsNomeTela As String
Private mvariLivroMod3 As Integer
Private mvariAtualizaSaldoCusto As Integer
Private mvariCustoMedio As Integer
Private mvarsEntradaSaidaCMP As String
Private mvariAtualizaCustoConsig As Integer
Private mvariAtualizaCustoDemo As Integer
Private mvariAtualizaCustoConserto As Integer
Private mvariAtualizaCustoOutros As Integer
Private mvariAtualizaCustoBenef As Integer
Private mvariAtualizaCustoConsig3 As Integer
Private mvariAtualizaCustoDemo3 As Integer
Private mvariAtualizaCustoConserto3 As Integer
Private mvariAtualizaCustoOutros3 As Integer
Private mvariAtualizaCustoBenef3 As Integer
Private mvariCodigoOrig As Integer
Private mvariAtualizaSoLote As Integer
Private mvariInventario As Integer
Private mvariNFDevolucao As Integer

Public Property Let iNFDevolucao(ByVal vData As Integer)
    mvariNFDevolucao = vData
End Property

Public Property Get iNFDevolucao() As Integer
    iNFDevolucao = mvariNFDevolucao
End Property

Public Property Let iInventario(ByVal vData As Integer)
    mvariInventario = vData
End Property

Public Property Get iInventario() As Integer
    iInventario = mvariInventario
End Property

Public Property Let iAtualizaSoLote(ByVal vData As Integer)
    mvariAtualizaSoLote = vData
End Property

Public Property Get iAtualizaSoLote() As Integer
    iAtualizaSoLote = mvariAtualizaSoLote
End Property

Public Property Let iCodigoOrig(ByVal vData As Integer)
    mvariCodigoOrig = vData
End Property

Public Property Get iCodigoOrig() As Integer
    iCodigoOrig = mvariCodigoOrig
End Property

Public Property Let iAtualizaCustoConsig3(ByVal vData As Integer)
    mvariAtualizaCustoConsig3 = vData
End Property

Public Property Get iAtualizaCustoConsig3() As Integer
    iAtualizaCustoConsig3 = mvariAtualizaCustoConsig3
End Property

Public Property Let iAtualizaCustoDemo3(ByVal vData As Integer)
    mvariAtualizaCustoDemo3 = vData
End Property

Public Property Get iAtualizaCustoDemo3() As Integer
    iAtualizaCustoDemo3 = mvariAtualizaCustoDemo3
End Property

Public Property Let iAtualizaCustoConserto3(ByVal vData As Integer)
    mvariAtualizaCustoConserto3 = vData
End Property

Public Property Get iAtualizaCustoConserto3() As Integer
    iAtualizaCustoConserto3 = mvariAtualizaCustoConserto3
End Property

Public Property Let iAtualizaCustoOutros3(ByVal vData As Integer)
    mvariAtualizaCustoOutros3 = vData
End Property

Public Property Get iAtualizaCustoOutros3() As Integer
    iAtualizaCustoOutros3 = mvariAtualizaCustoOutros3
End Property

Public Property Let iAtualizaCustoBenef3(ByVal vData As Integer)
    mvariAtualizaCustoBenef3 = vData
End Property

Public Property Get iAtualizaCustoBenef3() As Integer
    iAtualizaCustoBenef3 = mvariAtualizaCustoBenef3
End Property

Public Property Let iAtualizaCustoConsig(ByVal vData As Integer)
    mvariAtualizaCustoConsig = vData
End Property

Public Property Get iAtualizaCustoConsig() As Integer
    iAtualizaCustoConsig = mvariAtualizaCustoConsig
End Property

Public Property Let iAtualizaCustoDemo(ByVal vData As Integer)
    mvariAtualizaCustoDemo = vData
End Property

Public Property Get iAtualizaCustoDemo() As Integer
    iAtualizaCustoDemo = mvariAtualizaCustoDemo
End Property

Public Property Let iAtualizaCustoConserto(ByVal vData As Integer)
    mvariAtualizaCustoConserto = vData
End Property

Public Property Get iAtualizaCustoConserto() As Integer
    iAtualizaCustoConserto = mvariAtualizaCustoConserto
End Property

Public Property Let iAtualizaCustoOutros(ByVal vData As Integer)
    mvariAtualizaCustoOutros = vData
End Property

Public Property Get iAtualizaCustoOutros() As Integer
    iAtualizaCustoOutros = mvariAtualizaCustoOutros
End Property

Public Property Let iAtualizaCustoBenef(ByVal vData As Integer)
    mvariAtualizaCustoBenef = vData
End Property

Public Property Get iAtualizaCustoBenef() As Integer
    iAtualizaCustoBenef = mvariAtualizaCustoBenef
End Property

Public Property Let sEntradaSaidaCMP(ByVal vData As String)
    mvarsEntradaSaidaCMP = vData
End Property

Public Property Get sEntradaSaidaCMP() As String
    sEntradaSaidaCMP = mvarsEntradaSaidaCMP
End Property

Public Property Let iCustoMedio(ByVal vData As Integer)
    mvariCustoMedio = vData
End Property

Public Property Get iCustoMedio() As Integer
    iCustoMedio = mvariCustoMedio
End Property

Public Property Let iAtualizaSaldoCusto(ByVal vData As Integer)
    mvariAtualizaSaldoCusto = vData
End Property

Public Property Get iAtualizaSaldoCusto() As Integer
    iAtualizaSaldoCusto = mvariAtualizaSaldoCusto
End Property

Public Property Let iLivroMod3(ByVal vData As Integer)
    mvariLivroMod3 = vData
End Property

Public Property Get iLivroMod3() As Integer
    iLivroMod3 = mvariLivroMod3
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

Public Property Let iKardex(ByVal vData As Integer)
    mvariKardex = vData
End Property

Public Property Get iKardex() As Integer
    iKardex = mvariKardex
End Property

Public Property Let iTipoNumIntDocOrigem(ByVal vData As Integer)
    mvariTipoNumIntDocOrigem = vData
End Property

Public Property Get iTipoNumIntDocOrigem() As Integer
    iTipoNumIntDocOrigem = mvariTipoNumIntDocOrigem
End Property

Public Property Let sDescrNumIntDocOrigem(ByVal vData As String)
    mvarsDescrNumIntDocOrigem = vData
End Property

Public Property Get sDescrNumIntDocOrigem() As String
    sDescrNumIntDocOrigem = mvarsDescrNumIntDocOrigem
End Property

Public Property Let iAtualizaOP(ByVal vData As Integer)
    mvariAtualizaOP = vData
End Property

Public Property Get iAtualizaOP() As Integer
    iAtualizaOP = mvariAtualizaOP
End Property

Public Property Let iAtualizaBenef(ByVal vData As Integer)
    mvariAtualizaBenef = vData
End Property

Public Property Get iAtualizaBenef() As Integer
    iAtualizaBenef = mvariAtualizaBenef
End Property

Public Property Let iAtualizaCompra(ByVal vData As Integer)
    mvariAtualizaCompra = vData
End Property

Public Property Get iAtualizaCompra() As Integer
    iAtualizaCompra = mvariAtualizaCompra
End Property

Public Property Let iValidaDataUltMov(ByVal vData As Integer)
Attribute iValidaDataUltMov.VB_Description = "0=Não valida 1= Valida data do movimento para que seja maior ou igual a data do ultimo movimento de estoque"
    mvariValidaDataUltMov = vData
End Property

Public Property Get iValidaDataUltMov() As Integer
    iValidaDataUltMov = mvariValidaDataUltMov
End Property

Public Property Let iCustoInformado(ByVal vData As Integer)
    mvariCustoInformado = vData
End Property

Public Property Get iCustoInformado() As Integer
    iCustoInformado = mvariCustoInformado
End Property

Public Property Let iAtualizaCMRProd(ByVal vData As Integer)
    mvariAtualizaCMRProd = vData
End Property

Public Property Get iAtualizaCMRProd() As Integer
    iAtualizaCMRProd = mvariAtualizaCMRProd
End Property

Public Property Let iTransferencia(ByVal vData As Integer)
    mvariTransferencia = vData
End Property

Public Property Get iTransferencia() As Integer
    iTransferencia = mvariTransferencia
End Property

Public Property Let iAtualizaMovEstoque(ByVal vData As Integer)
    mvariAtualizaMovEstoque = vData
End Property

Public Property Get iAtualizaMovEstoque() As Integer
    iAtualizaMovEstoque = mvariAtualizaMovEstoque
End Property

Public Property Let iAtualizaReserva(ByVal vData As Integer)
    mvariAtualizaReserva = vData
End Property

Public Property Get iAtualizaReserva() As Integer
    iAtualizaReserva = mvariAtualizaReserva
End Property

Public Property Let iAtualizaRecebIndisp(ByVal vData As Integer)
    mvariAtualizaRecebIndisp = vData
End Property

Public Property Get iAtualizaRecebIndisp() As Integer
    iAtualizaRecebIndisp = mvariAtualizaRecebIndisp
End Property

Public Property Let iAtualizaDefeituosa(ByVal vData As Integer)
    mvariAtualizaDefeituosa = vData
End Property

Public Property Get iAtualizaDefeituosa() As Integer
    iAtualizaDefeituosa = mvariAtualizaDefeituosa
End Property

Public Property Let iAtualizaNossaDisp(ByVal vData As Integer)
    mvariAtualizaNossaDisp = vData
End Property

Public Property Get iAtualizaNossaDisp() As Integer
    iAtualizaNossaDisp = mvariAtualizaNossaDisp
End Property

Public Property Let iAtualizaIndOutras(ByVal vData As Integer)
    mvariAtualizaIndOutras = vData
End Property

Public Property Get iAtualizaIndOutras() As Integer
    iAtualizaIndOutras = mvariAtualizaIndOutras
End Property

Public Property Let iAtualizaOutras(ByVal vData As Integer)
    mvariAtualizaOutras = vData
End Property

Public Property Get iAtualizaOutras() As Integer
    iAtualizaOutras = mvariAtualizaOutras
End Property

Public Property Let iProdutoDeTerc(ByVal vData As Integer)
    mvariProdutoDeTerc = vData
End Property

Public Property Get iProdutoDeTerc() As Integer
    iProdutoDeTerc = mvariProdutoDeTerc
End Property

Public Property Let iAtualizaConserto(ByVal vData As Integer)
    mvariAtualizaConserto = vData
End Property

Public Property Get iAtualizaConserto() As Integer
    iAtualizaConserto = mvariAtualizaConserto
End Property

Public Property Let iAtualizaDemo(ByVal vData As Integer)
    mvariAtualizaDemo = vData
End Property

Public Property Get iAtualizaDemo() As Integer
    iAtualizaDemo = mvariAtualizaDemo
End Property

Public Property Let iAtualizaConsig(ByVal vData As Integer)
    mvariAtualizaConsig = vData
End Property

Public Property Get iAtualizaConsig() As Integer
    iAtualizaConsig = mvariAtualizaConsig
End Property

Public Property Let iAtualizaCusto(ByVal vData As Integer)
    mvariAtualizaCusto = vData
End Property

Public Property Get iAtualizaCusto() As Integer
    iAtualizaCusto = mvariAtualizaCusto
End Property

Public Property Let iValidoMovInt(ByVal vData As Integer)
    mvariValidoMovInt = vData
End Property

Public Property Get iValidoMovInt() As Integer
    iValidoMovInt = mvariValidoMovInt
End Property

Public Property Let iEditavel(ByVal vData As Integer)
    mvariEditavel = vData
End Property

Public Property Get iEditavel() As Integer
    iEditavel = mvariEditavel
End Property

Public Property Let iAtualizaVendaConsig3(ByVal vData As Integer)
    mvariAtualizaVendaConsig3 = vData
End Property

Public Property Get iAtualizaVendaConsig3() As Integer
    iAtualizaVendaConsig3 = mvariAtualizaVendaConsig3
End Property

Public Property Let iAtualizaVenda(ByVal vData As Integer)
    mvariAtualizaVenda = vData
End Property

Public Property Get iAtualizaVenda() As Integer
    iAtualizaVenda = mvariAtualizaVenda
End Property

Public Property Let iAtualizaConsumo(ByVal vData As Integer)
    mvariAtualizaConsumo = vData
End Property

Public Property Get iAtualizaConsumo() As Integer
    iAtualizaConsumo = mvariAtualizaConsumo
End Property

Public Property Let iInativo(ByVal vData As Integer)
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let sEntradaOuSaida(ByVal vData As String)
    mvarsEntradaOuSaida = vData
End Property

Public Property Get sEntradaOuSaida() As String
    sEntradaOuSaida = mvarsEntradaOuSaida
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Sub Copiar(objTipoMovEstOrig As ClassTipoMovEst)

    With objTipoMovEstOrig
        mvariCodigo = .iCodigo
        mvarsDescricao = .sDescricao
        mvarsEntradaOuSaida = .sEntradaOuSaida
        mvariInativo = .iInativo
        mvariAtualizaConsumo = .iAtualizaConsumo
        mvariAtualizaVendaConsig3 = .iAtualizaVendaConsig3
        mvariAtualizaVenda = .iAtualizaVenda
        mvariEditavel = .iEditavel
        mvariValidoMovInt = .iValidoMovInt
        mvariAtualizaCusto = .iAtualizaCusto
        mvariAtualizaConsig = .iAtualizaConsig
        mvariAtualizaDemo = .iAtualizaDemo
        mvariAtualizaConserto = .iAtualizaConserto
        mvariProdutoDeTerc = .iProdutoDeTerc
        mvariAtualizaOutras = .iAtualizaOutras
        mvariAtualizaIndOutras = .iAtualizaIndOutras
        mvariAtualizaNossaDisp = .iAtualizaNossaDisp
        mvariAtualizaDefeituosa = .iAtualizaDefeituosa
        mvariAtualizaRecebIndisp = .iAtualizaRecebIndisp
        mvariAtualizaMovEstoque = .iAtualizaMovEstoque
        mvariAtualizaReserva = .iAtualizaReserva
        mvariTransferencia = .iTransferencia
        mvariAtualizaCMRProd = .iAtualizaCMRProd
        mvariCustoInformado = .iCustoInformado
        mvariValidaDataUltMov = .iValidaDataUltMov
        mvariAtualizaCompra = .iAtualizaCompra
        mvariAtualizaBenef = .iAtualizaBenef
        mvariAtualizaOP = .iAtualizaOP
        mvarsDescrNumIntDocOrigem = .sDescrNumIntDocOrigem
        mvariTipoNumIntDocOrigem = .iTipoNumIntDocOrigem
        mvariKardex = .iKardex
        mvarsSigla = .sSigla
        mvarsNomeTela = .sNomeTela
        mvariLivroMod3 = .iLivroMod3
        mvariAtualizaSaldoCusto = .iAtualizaSaldoCusto
        mvariCustoMedio = .iCustoMedio
        mvarsEntradaSaidaCMP = .sEntradaSaidaCMP
        mvariAtualizaCustoConsig = .iAtualizaCustoConsig
        mvariAtualizaCustoDemo = .iAtualizaCustoDemo
        mvariAtualizaCustoConserto = .iAtualizaCustoConserto
        mvariAtualizaCustoOutros = .iAtualizaCustoOutros
        mvariAtualizaCustoBenef = .iAtualizaCustoBenef
        mvariAtualizaCustoConsig3 = .iAtualizaCustoConsig3
        mvariAtualizaCustoDemo3 = .iAtualizaCustoDemo3
        mvariAtualizaCustoConserto3 = .iAtualizaCustoConserto3
        mvariAtualizaCustoOutros3 = .iAtualizaCustoOutros3
        mvariAtualizaCustoBenef3 = .iAtualizaCustoBenef3
        mvariCodigoOrig = .iCodigoOrig
        mvariAtualizaSoLote = .iAtualizaSoLote
        mvariInventario = .iInventario
        mvariNFDevolucao = .iNFDevolucao
    End With
    
End Sub

