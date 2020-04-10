VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIntegracaoSldProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntGer As Long
Private mvarlNumIntArq As Long
Private mvarlSeqRegistro As Long
Private mvariTipoInt As Integer
Private mvarsCodProduto As String
Private mvardSaldoDisp As Double
Private mvariAlmoxarifado As Integer
Private mvariComErro As Integer
Private mvardtDataAtualizacao As Date

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntGer(ByVal vData As Long)
    mvarlNumIntGer = vData
End Property

Public Property Get lNumIntGer() As Long
    lNumIntGer = mvarlNumIntGer
End Property

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let lSeqRegistro(ByVal vData As Long)
    mvarlSeqRegistro = vData
End Property

Public Property Get lSeqRegistro() As Long
    lSeqRegistro = mvarlSeqRegistro
End Property

Public Property Let iTipoInt(ByVal vData As Integer)
    mvariTipoInt = vData
End Property

Public Property Get iTipoInt() As Integer
    iTipoInt = mvariTipoInt
End Property

Public Property Let sCodProduto(ByVal vData As String)
    mvarsCodProduto = vData
End Property

Public Property Get sCodProduto() As String
    sCodProduto = mvarsCodProduto
End Property

Public Property Let dSaldoDisp(ByVal vData As Double)
    mvardSaldoDisp = vData
End Property

Public Property Get dSaldoDisp() As Double
    dSaldoDisp = mvardSaldoDisp
End Property

Public Property Let iComErro(ByVal vData As Integer)
    mvariComErro = vData
End Property

Public Property Get iComErro() As Integer
    iComErro = mvariComErro
End Property

Public Property Let dtDataAtualizacao(ByVal vData As Date)
    mvardtDataAtualizacao = vData
End Property

Public Property Get dtDataAtualizacao() As Date
    dtDataAtualizacao = mvardtDataAtualizacao
End Property

Public Function Monta_Registro_SldProd(sRegistro As String)

Dim lErro As Long

On Error GoTo Erro_Monta_Registro_SldProd

    sRegistro = "01"

    sRegistro = sRegistro & FormataCpoTexto(sCodProduto, 20) & FormataCpoValor(dSaldoDisp, 10) & FormataCpoNum(iAlmoxarifado, 4)

    Monta_Registro_SldProd = SUCESSO

    Exit Function

Erro_Monta_Registro_SldProd:

    Monta_Registro_SldProd = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190090)

    End Select

    Exit Function
    
End Function

