VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelDRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvariTipo As Integer
Private mvariNivel As Integer
Private mvarsTitulo As String
Private mvardValor As Double
Private mvarsModelo As String
Private mvariPosicao As Integer
Private mvariImprime As Integer
Private mvardValorExercAnt As Double
Private mvardValorPerAcumAnt As Double
Private mvardValorPerAcumAtual As Double
Private mvariExercicio As Integer 'indica se vai usar as contas do exercicio selecionado ou do ano anterior
Private mvarobjRelDREDetPer As New ClassRelDREDetPer

Public Property Set objRelDREDetPer(ByVal vData As ClassRelDREDetPer)
    Set mvarobjRelDREDetPer = vData
End Property

Public Property Get objRelDREDetPer() As ClassRelDREDetPer
    Set objRelDREDetPer = mvarobjRelDREDetPer
End Property

Public Property Let iExercicio(ByVal vData As Integer)
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let dValorPerAcumAtual(ByVal vData As Double)
    mvardValorPerAcumAtual = vData
End Property

Public Property Get dValorPerAcumAtual() As Double
    dValorPerAcumAtual = mvardValorPerAcumAtual
End Property

Public Property Let dValorPerAcumAnt(ByVal vData As Double)
    mvardValorPerAcumAnt = vData
End Property

Public Property Get dValorPerAcumAnt() As Double
    dValorPerAcumAnt = mvardValorPerAcumAnt
End Property

Public Property Let dValorExercAnt(ByVal vData As Double)
    mvardValorExercAnt = vData
End Property

Public Property Get dValorExercAnt() As Double
    dValorExercAnt = mvardValorExercAnt
End Property

Public Property Let iImprime(ByVal vData As Integer)
Attribute iImprime.VB_Description = "0=Nao imprime 1=Imprime"
    mvariImprime = vData
End Property

Public Property Get iImprime() As Integer
    iImprime = mvariImprime
End Property

Public Property Let iPosicao(ByVal vData As Integer)
Attribute iPosicao.VB_Description = "Posicao do Elemento na arvore"
    mvariPosicao = vData
End Property

Public Property Get iPosicao() As Integer
    iPosicao = mvariPosicao
End Property

Public Property Let sModelo(ByVal vData As String)
Attribute sModelo.VB_Description = "Identificador do Modelo"
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "Valor que será calculado antes da execução do relatório"
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sTitulo(ByVal vData As String)
Attribute sTitulo.VB_Description = "Titulo do Elemento (é o texto que será impresso)"
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property

Public Property Let iNivel(ByVal vData As Integer)
Attribute iNivel.VB_Description = "Nivel de Identacao do Elemento no Relatorio"
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "0=GrupoConta 1=Formula 2=Titulo"
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Sequencial identificador do elemento do DRE. Identifica a sequencia em que é exibido/impresso o elemento."
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



