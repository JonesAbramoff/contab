VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEdicaoTela_Controle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvarsGrupoUsuarios As String
Private mvarsNomeControle As String
Private mvarsTitulo As String
Private mvariLargura As Integer
Private mvarizOrder As Integer
Private mvariHabilitado As Integer
Private mvariHabilitadoPadrao As Integer
Private mvariVisivel As Integer
Private mvariAltura As Integer
Private mvariEsquerda As Integer
Private mvariTopo As Integer
Private mvariTabStop As Integer
Private mvarsContainer As String
'local variable(s) to hold property value(s)
Private mvariAlturaPadrao As Integer 'local copy
Private mvariEsquerdaPadrao As Integer 'local copy
Private mvariLarguraPadrao As Integer 'local copy
Private mvariTabStopPadrao As Integer 'local copy
Private mvariTopoPadrao As Integer 'local copy
Private mvariVisivelPadrao As Integer 'local copy
Private mvarsContainerPadrao As String 'local copy
Private mvarsTituloPadrao As String 'local copy
'local variable(s) to hold property value(s)
Private mvariIndiceContainer As Integer 'local copy
Private mvariIndiceContainerPadrao As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvariTabIndex As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvariTabIndexPadrao As Integer 'local copy
Public Property Let iTabIndexPadrao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTabIndexPadrao = 5
    mvariTabIndexPadrao = vData
End Property


Public Property Get iTabIndexPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTabIndexPadrao
    iTabIndexPadrao = mvariTabIndexPadrao
End Property

Public Property Let izOrder(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.izOrder = 5
    mvarizOrder = vData
End Property


Public Property Get izOrder() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.izOrder
    izOrder = mvarizOrder
End Property

Public Property Let iHabilitado(ByVal vData As Integer)
    mvariHabilitado = vData
End Property

Public Property Get iHabilitado() As Integer
    iHabilitado = mvariHabilitado
End Property

Public Property Let iHabilitadoPadrao(ByVal vData As Integer)
    mvariHabilitadoPadrao = vData
End Property

Public Property Get iHabilitadoPadrao() As Integer
    iHabilitadoPadrao = mvariHabilitadoPadrao
End Property

Public Property Let iTabIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTabIndex = 5
    mvariTabIndex = vData
End Property


Public Property Get iTabIndex() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTabIndex
    iTabIndex = mvariTabIndex
End Property



Public Property Let iIndiceContainerPadrao(ByVal vData As Integer)
Attribute iIndiceContainerPadrao.VB_Description = "Indide Padrão"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIndiceContainerPadrao = 5
    mvariIndiceContainerPadrao = vData
End Property


Public Property Get iIndiceContainerPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIndiceContainerPadrao
    iIndiceContainerPadrao = mvariIndiceContainerPadrao
End Property



Public Property Let iIndiceContainer(ByVal vData As Integer)
Attribute iIndiceContainer.VB_Description = "Indice do Container"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIndiceContainer = 5
    mvariIndiceContainer = vData
End Property


Public Property Get iIndiceContainer() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIndiceContainer
    iIndiceContainer = mvariIndiceContainer
End Property



Public Property Let sTituloPadrao(ByVal vData As String)
Attribute sTituloPadrao.VB_Description = "Titulo Padrão do Controle (no caso de SGELabel)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTituloPadrao = 5
    mvarsTituloPadrao = vData
End Property


Public Property Get sGrupoUsuarios() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sGrupoUsuarios
    sGrupoUsuarios = mvarsGrupoUsuarios
End Property

Public Property Let sGrupoUsuarios(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sGrupoUsuarios = 5
    mvarsGrupoUsuarios = vData
End Property


Public Property Get sTituloPadrao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTituloPadrao
    sTituloPadrao = mvarsTituloPadrao
End Property

Public Property Let sContainerPadrao(ByVal vData As String)
Attribute sContainerPadrao.VB_Description = "Nome Padrao do Controle que contem este controle"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContainerPadrao = 5
    mvarsContainerPadrao = vData
End Property


Public Property Get sContainerPadrao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContainerPadrao
    sContainerPadrao = mvarsContainerPadrao
End Property



Public Property Let iVisivelPadrao(ByVal vData As Integer)
Attribute iVisivelPadrao.VB_Description = "0=Invisivel 1=Visivel"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iVisivelPadrao = 5
    mvariVisivelPadrao = vData
End Property


Public Property Get iVisivelPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iVisivelPadrao
    iVisivelPadrao = mvariVisivelPadrao
End Property



Public Property Let iTopoPadrao(ByVal vData As Integer)
Attribute iTopoPadrao.VB_Description = "Posicao Padrão do Topo do controle em relação ao seu container em Twips"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTopoPadrao = 5
    mvariTopoPadrao = vData
End Property


Public Property Get iTopoPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTopoPadrao
    iTopoPadrao = mvariTopoPadrao
End Property



Public Property Let iTabStopPadrao(ByVal vData As Integer)
Attribute iTabStopPadrao.VB_Description = "0=Falso 1=Verdadeiro"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTabStopPadrao = 5
    mvariTabStopPadrao = vData
End Property


Public Property Get iTabStopPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTabStopPadrao
    iTabStopPadrao = mvariTabStopPadrao
End Property



Public Property Let iLarguraPadrao(ByVal vData As Integer)
Attribute iLarguraPadrao.VB_Description = "Largura Padrão do Controle em Twips"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLarguraPadrao = 5
    mvariLarguraPadrao = vData
End Property


Public Property Get iLarguraPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLarguraPadrao
    iLarguraPadrao = mvariLarguraPadrao
End Property



Public Property Let iEsquerdaPadrao(ByVal vData As Integer)
Attribute iEsquerdaPadrao.VB_Description = "Posicao Padrão do Lado esquerdo do controle em relação ao seu container em Twips"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iEsquerdaPadrao = 5
    mvariEsquerdaPadrao = vData
End Property


Public Property Get iEsquerdaPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iEsquerdaPadrao
    iEsquerdaPadrao = mvariEsquerdaPadrao
End Property



Public Property Let iAlturaPadrao(ByVal vData As Integer)
Attribute iAlturaPadrao.VB_Description = "Altura Padrao do Controle em Twips "
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAlturaPadrao = 5
    mvariAlturaPadrao = vData
End Property


Public Property Get iAlturaPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAlturaPadrao
    iAlturaPadrao = mvariAlturaPadrao
End Property




Public Property Let sContainer(ByVal vData As String)
Attribute sContainer.VB_Description = "Nome do Controle que contem este controle"
    mvarsContainer = vData
End Property

Public Property Get sContainer() As String
    sContainer = mvarsContainer
End Property

Public Property Let iTabStop(ByVal vData As Integer)
Attribute iTabStop.VB_Description = "0=Falso 1=Verdadeiro"
    mvariTabStop = vData
End Property

Public Property Get iTabStop() As Integer
    iTabStop = mvariTabStop
End Property

Public Property Let iTopo(ByVal vData As Integer)
Attribute iTopo.VB_Description = "Posicao do Topo do controle em relação ao seu container em Twips"
    mvariTopo = vData
End Property

Public Property Get iTopo() As Integer
    iTopo = mvariTopo
End Property

Public Property Let iEsquerda(ByVal vData As Integer)
Attribute iEsquerda.VB_Description = "Posicao do Lado esquerdo do controle em relação ao seu container em Twips"
    mvariEsquerda = vData
End Property

Public Property Get iEsquerda() As Integer
    iEsquerda = mvariEsquerda
End Property

Public Property Let iAltura(ByVal vData As Integer)
Attribute iAltura.VB_Description = "Altura do Controle em Twips"
    mvariAltura = vData
End Property

Public Property Get iAltura() As Integer
    iAltura = mvariAltura
End Property

Public Property Let iVisivel(ByVal vData As Integer)
Attribute iVisivel.VB_Description = "0=Invisivel 1=Visivel"
    mvariVisivel = vData
End Property

Public Property Get iVisivel() As Integer
    iVisivel = mvariVisivel
End Property

Public Property Let iLargura(ByVal vData As Integer)
Attribute iLargura.VB_Description = "Largura do Controle em Twips"
    mvariLargura = vData
End Property

Public Property Get iLargura() As Integer
    iLargura = mvariLargura
End Property

Public Property Let sTitulo(ByVal vData As String)
Attribute sTitulo.VB_Description = "Título do controle para caso como SGELabel"
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property

Public Property Let sNomeControle(ByVal vData As String)
Attribute sNomeControle.VB_Description = "Nome do Controle"
    mvarsNomeControle = vData
End Property

Public Property Get sNomeControle() As String
    sNomeControle = mvarsNomeControle
End Property

Public Property Let sNomeTela(ByVal vData As String)
Attribute sNomeTela.VB_Description = "Nome da Tela onde contém o controle"
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property
