VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOPFichaProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'#########################################
'Inserido por Wagner - CROMATON 03/11/04
Private mvariFilialEmpresa As Integer
Private mvariMistura As Integer
Private mvariTempoMistura As Integer
Private mvariCargaDrays As Integer
Private mvariEmbalagem As Integer
Private mvarsCodigoOP As String

'################
'03/06/05
Private mvardProcessoAgua As Double
Private mvardProcessoOleo As Double
Private mvariAmperagem As Integer
Private mvardVelRotoresDe As Double
Private mvardVelRotoresAte As Double
Private mvardAberturaGate As Double
Private mvardTempCamaraZDe(1 To 3) As Double
Private mvardTempCamaraZAte(1 To 3) As Double
Private mvarsTela As String
Private mvardTempExtrusoraZDe(1 To 7) As Double
Private mvardTempExtrusoraZAte(1 To 7) As Double
Private mvardVelVariadorDe As Double
Private mvardVelVariadorAte As Double
Private mvarsObservacao As String
Private mvarsAmostra As String
'################

Public Property Let iEmbalagem(ByVal vData As Integer)
    mvariEmbalagem = vData
End Property

Public Property Get iEmbalagem() As Integer
    iEmbalagem = mvariEmbalagem
End Property

Public Property Let iCargaDrays(ByVal vData As Integer)
    mvariCargaDrays = vData
End Property

Public Property Get iCargaDrays() As Integer
    iCargaDrays = mvariCargaDrays
End Property

Public Property Let iTempoMistura(ByVal vData As Integer)
    mvariTempoMistura = vData
End Property

Public Property Get iTempoMistura() As Integer
    iTempoMistura = mvariTempoMistura
End Property

Public Property Let iMistura(ByVal vData As Integer)
    mvariMistura = vData
End Property

Public Property Get iMistura() As Integer
    iMistura = mvariMistura
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCodigoOP(ByVal vData As String)
    mvarsCodigoOP = vData
End Property

Public Property Get sCodigoOP() As String
    sCodigoOP = mvarsCodigoOP
End Property

'#######################
'03/06/05
Public Property Let dProcessoAgua(ByVal vData As Double)
    mvardProcessoAgua = vData
End Property

Public Property Get dProcessoAgua() As Double
    dProcessoAgua = mvardProcessoAgua
End Property

Public Property Let dProcessoOleo(ByVal vData As Double)
    mvardProcessoOleo = vData
End Property

Public Property Get dProcessoOleo() As Double
    dProcessoOleo = mvardProcessoOleo
End Property

Public Property Let iAmperagem(ByVal vData As Integer)
    mvariAmperagem = vData
End Property

Public Property Get iAmperagem() As Integer
    iAmperagem = mvariAmperagem
End Property

Public Property Let dVelRotoresDe(ByVal vData As Double)
    mvardVelRotoresDe = vData
End Property

Public Property Get dVelRotoresDe() As Double
    dVelRotoresDe = mvardVelRotoresDe
End Property

Public Property Let dVelRotoresAte(ByVal vData As Double)
    mvardVelRotoresAte = vData
End Property

Public Property Get dVelRotoresAte() As Double
    dVelRotoresAte = mvardVelRotoresAte
End Property

Public Property Let dAberturaGate(ByVal vData As Double)
    mvardAberturaGate = vData
End Property

Public Property Get dAberturaGate() As Double
    dAberturaGate = mvardAberturaGate
End Property

Public Property Let dTempCamaraZDe(ByVal iIndice As Integer, ByVal vData As Double)
    mvardTempCamaraZDe(iIndice) = vData
End Property

Public Property Get dTempCamaraZDe(ByVal iIndice As Integer) As Double
    dTempCamaraZDe = mvardTempCamaraZDe(iIndice)
End Property

Public Property Let dTempCamaraZAte(ByVal iIndice As Integer, ByVal vData As Double)
    mvardTempCamaraZAte(iIndice) = vData
End Property

Public Property Get dTempCamaraZAte(ByVal iIndice As Integer) As Double
    dTempCamaraZAte = mvardTempCamaraZAte(iIndice)
End Property

Public Property Let sTela(ByVal vData As String)
    mvarsTela = vData
End Property

Public Property Get sTela() As String
    sTela = mvarsTela
End Property

Public Property Let dTempExtrusoraZDe(ByVal iIndice As Integer, ByVal vData As Double)
    mvardTempExtrusoraZDe(iIndice) = vData
End Property

Public Property Get dTempExtrusoraZDe(ByVal iIndice As Integer) As Double
    dTempExtrusoraZDe = mvardTempExtrusoraZDe(iIndice)
End Property

Public Property Let dTempExtrusoraZAte(ByVal iIndice As Integer, ByVal vData As Double)
    mvardTempExtrusoraZAte(iIndice) = vData
End Property

Public Property Get dTempExtrusoraZAte(ByVal iIndice As Integer) As Double
    dTempExtrusoraZAte = mvardTempExtrusoraZAte(iIndice)
End Property

Public Property Let dVelVariadorDe(ByVal vData As Double)
    mvardVelVariadorDe = vData
End Property

Public Property Get dVelVariadorDe() As Double
    dVelVariadorDe = mvardVelVariadorDe
End Property

Public Property Let dVelVariadorAte(ByVal vData As Double)
    mvardVelVariadorAte = vData
End Property

Public Property Get dVelVariadorAte() As Double
    dVelVariadorAte = mvardVelVariadorAte
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sAmostra(ByVal vData As String)
    mvarsAmostra = vData
End Property

Public Property Get sAmostra() As String
    sAmostra = mvarsAmostra
End Property


'#########################################

