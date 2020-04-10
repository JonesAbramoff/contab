VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object

'??? atencao: se incluir algum novo campo alterar o metodo Copiar
Private mvarsEndereco As String
Private mvarsBairro As String
Private mvarsCidade As String
Private mvarsSiglaEstado As String
Private mvariCodigoPais As Integer
Private mvarsCEP As String
Private mvarsTelefone1 As String
Private mvarsTelefone2 As String
Private mvarsEmail As String
Private mvarsFax As String
Private mvarsContato As String
Private mvarlCodigo As Long

Private mvarsReferencia As String
Private mvarsLogradouro As String
Private mvarsComplemento As String
Private mvarsTipoLogradouro As String
Private mvarsEmail2 As String
Private mvarlNumero As Long
Private mvariTelDDD1 As Integer
Private mvariTelDDD2 As Integer
Private mvariFaxDDD As Integer
Private mvarsTelNumero1 As String
Private mvarsTelNumero2 As String
Private mvarsFaxNumero As String

Public Property Let sFaxNumero(ByVal vData As String)
    mvarsFaxNumero = vData
End Property

Public Property Get sFaxNumero() As String
    sFaxNumero = mvarsFaxNumero
End Property

Public Property Let sTelNumero2(ByVal vData As String)
    mvarsTelNumero2 = vData
End Property

Public Property Get sTelNumero2() As String
    sTelNumero2 = mvarsTelNumero2
End Property

Public Property Let sTelNumero1(ByVal vData As String)
    mvarsTelNumero1 = vData
End Property

Public Property Get sTelNumero1() As String
    sTelNumero1 = mvarsTelNumero1
End Property

Public Property Let iFaxDDD(ByVal vData As Integer)
    mvariFaxDDD = vData
End Property

Public Property Get iFaxDDD() As Integer
    iFaxDDD = mvariFaxDDD
End Property

Public Property Let iTelDDD2(ByVal vData As Integer)
    mvariTelDDD2 = vData
End Property

Public Property Get iTelDDD2() As Integer
    iTelDDD2 = mvariTelDDD2
End Property

Public Property Let iTelDDD1(ByVal vData As Integer)
    mvariTelDDD1 = vData
End Property

Public Property Get iTelDDD1() As Integer
    iTelDDD1 = mvariTelDDD1
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let sEmail2(ByVal vData As String)
    mvarsEmail2 = vData
End Property

Public Property Get sEmail2() As String
    sEmail2 = mvarsEmail2
End Property

Public Property Let sTipoLogradouro(ByVal vData As String)
    mvarsTipoLogradouro = vData
End Property

Public Property Get sTipoLogradouro() As String
    sTipoLogradouro = mvarsTipoLogradouro
End Property

Public Property Let sComplemento(ByVal vData As String)
    mvarsComplemento = vData
End Property

Public Property Get sComplemento() As String
    sComplemento = mvarsComplemento
End Property

Public Property Let sLogradouro(ByVal vData As String)
    mvarsLogradouro = vData
End Property

Public Property Get sLogradouro() As String
    sLogradouro = mvarsLogradouro
End Property

Public Property Let sReferencia(ByVal vData As String)
    mvarsReferencia = vData
End Property

Public Property Get sReferencia() As String
    sReferencia = mvarsReferencia
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let sFax(ByVal vData As String)
    mvarsFax = vData
End Property

Public Property Get sFax() As String
    sFax = mvarsFax
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let sTelefone2(ByVal vData As String)
    mvarsTelefone2 = vData
End Property

Public Property Get sTelefone2() As String
    sTelefone2 = mvarsTelefone2
End Property

Public Property Let sTelefone1(ByVal vData As String)
    mvarsTelefone1 = vData
End Property

Public Property Get sTelefone1() As String
    sTelefone1 = mvarsTelefone1
End Property

Public Property Let sCEP(ByVal vData As String)
    mvarsCEP = vData
End Property

Public Property Get sCEP() As String
    sCEP = mvarsCEP
End Property

Public Property Let iCodigoPais(ByVal vData As Integer)
    mvariCodigoPais = vData
End Property

Public Property Get iCodigoPais() As Integer
    iCodigoPais = mvariCodigoPais
End Property

Public Property Let sSiglaEstado(ByVal vData As String)
    mvarsSiglaEstado = vData
End Property

Public Property Get sSiglaEstado() As String
    sSiglaEstado = mvarsSiglaEstado
End Property

Public Property Let sCidade(ByVal vData As String)
    mvarsCidade = vData
End Property

Public Property Get sCidade() As String
    sCidade = mvarsCidade
End Property

Public Property Let sBairro(ByVal vData As String)
    mvarsBairro = vData
End Property

Public Property Get sBairro() As String
    sBairro = mvarsBairro
End Property

Public Property Let sEndereco(ByVal vData As String)
    mvarsEndereco = vData
End Property

Public Property Get sEndereco() As String
    sEndereco = mvarsEndereco
End Property

Public Sub Copiar(objEnderecoOrig As ClassEndereco)
'atualiza o proprio conteudo copiando dados de outra instancia

    mvarsEndereco = objEnderecoOrig.sEndereco
    mvarsBairro = objEnderecoOrig.sBairro
    mvarsCidade = objEnderecoOrig.sCidade
    mvarsSiglaEstado = objEnderecoOrig.sSiglaEstado
    mvariCodigoPais = objEnderecoOrig.iCodigoPais
    mvarsCEP = objEnderecoOrig.sCEP
    mvarsTelefone1 = objEnderecoOrig.sTelefone1
    mvarsTelefone2 = objEnderecoOrig.sTelefone2
    mvarsEmail = objEnderecoOrig.sEmail
    mvarsFax = objEnderecoOrig.sFax
    mvarsContato = objEnderecoOrig.sContato
    mvarlCodigo = objEnderecoOrig.lCodigo
    
    mvarsReferencia = objEnderecoOrig.sReferencia
    mvarsLogradouro = objEnderecoOrig.sLogradouro
    mvarsComplemento = objEnderecoOrig.sComplemento
    mvarsTipoLogradouro = objEnderecoOrig.sTipoLogradouro
    mvarsEmail2 = objEnderecoOrig.sEmail2
    mvarlNumero = objEnderecoOrig.lNumero
    mvariTelDDD1 = objEnderecoOrig.iTelDDD1
    mvariTelDDD2 = objEnderecoOrig.iTelDDD2
    mvariFaxDDD = objEnderecoOrig.iFaxDDD
    mvarsTelNumero1 = objEnderecoOrig.sTelNumero1
    mvarsTelNumero2 = objEnderecoOrig.sTelNumero2
    mvarsFaxNumero = objEnderecoOrig.sFaxNumero

End Sub

