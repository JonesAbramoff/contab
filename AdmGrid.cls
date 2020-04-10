VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariSaidaCelula As Integer
Private mvariTeclado As Integer
Private mvarlErroSaidaCelula As Integer
Private mvariColunaAntiga As Integer
Private mvariLinhaAntiga As Integer
Private mvarobjGrid As Object
Private mvarobjControle As Control
Private mvarcolColuna As Collection
Private mvarcolCampo As Collection
Private mvarobjForm As Object
Private mvariLinhasVisiveis As Integer
Private mvariLinhasExistentes As Integer
Private mvariGridLargAuto As Integer
Private mvariProibidoExcluir As Integer
Private mvariProibidoIncluir As Integer
Private mvarobjCheckboxChecked As Picture
Private mvarobjCheckboxUnchecked As Picture
Private mvarobjCheckboxGrayed As Picture
Private mvarobjOptionButtonChecked As Picture
Private mvarobjOptionButtonUnchecked As Picture
Private mvarobjButton As Picture
Private mvariExecutarFuncao As Integer
Private mvariIncluirHScroll As Integer
Private mvariExecutaRotinaEnable As Integer
Private mvariIncluirBotao As Integer
Private mvariExecutaSaidaCelula As Integer
Private mvariEscEnter As Integer 'Indica que foi teclado Escape ou Enter
Private mvariProibidoIncluirNoMeio As Integer 'Nao Permite incluir uma linha no meio do grid = 0 - Permite incluir uma linha no meio do grid = 1
Private mvarcolIndex As Collection 'quando o grid tiver control arrays nas suas colunas, guardar os indices do elemento no control array. As colunas que nao se referenciarem a elementos de control array devem incluir um item zerado
Private mvariNaoTrataEnter As Integer
Private mvariGerencial As Integer
Private mvariColunaMarca As Integer
Private mvariFatorTam As Integer

Public Property Let iFatorTam(ByVal vData As Integer)
    mvariFatorTam = vData
End Property

Public Property Get iFatorTam() As Integer
    iFatorTam = mvariFatorTam
End Property

Public Property Let iColunaMarca(ByVal vData As Integer)
    mvariColunaMarca = vData
End Property

Public Property Get iColunaMarca() As Integer
    iColunaMarca = mvariColunaMarca
End Property

Public Property Let iGerencial(ByVal vData As Integer)
    mvariGerencial = vData
End Property

Public Property Get iGerencial() As Integer
    iGerencial = mvariGerencial
End Property

Public Property Let iNaoTrataEnter(ByVal vData As Integer)
    mvariNaoTrataEnter = vData
End Property

Public Property Get iNaoTrataEnter() As Integer
    iNaoTrataEnter = mvariNaoTrataEnter
End Property

Public Property Get colIndex() As Collection
    Set colIndex = mvarcolIndex
End Property

Public Property Let iProibidoIncluirNoMeioGrid(ByVal vData As Integer)
    mvariProibidoIncluirNoMeio = vData
End Property

Public Property Get iProibidoIncluirNoMeioGrid() As Integer
    iProibidoIncluirNoMeioGrid = mvariProibidoIncluirNoMeio
End Property

Public Property Let iEscEnter(ByVal vData As Integer)
    mvariEscEnter = vData
End Property

Public Property Get iEscEnter() As Integer
    iEscEnter = mvariEscEnter
End Property

Public Property Let iExecutaSaidaCelula(ByVal vData As Integer)
    mvariExecutaSaidaCelula = vData
End Property

Public Property Get iExecutaSaidaCelula() As Integer
    iExecutaSaidaCelula = mvariExecutaSaidaCelula
End Property

Public Property Let iIncluirBotao(ByVal vData As Integer)
    mvariIncluirBotao = vData
End Property

Public Property Get iIncluirBotao() As Integer
    iIncluirBotao = mvariIncluirBotao
End Property

Public Property Let iExecutaRotinaEnable(ByVal vData As Integer)
    mvariExecutaRotinaEnable = vData
End Property

Public Property Get iExecutaRotinaEnable() As Integer
    iExecutaRotinaEnable = mvariExecutaRotinaEnable
End Property


Public Property Let iIncluirHScroll(ByVal vData As Integer)
Attribute iIncluirHScroll.VB_Description = "Permite visualizar mais uma linha alem das existentes. Esta linha poderá ser encoberta pela barra de scroll horizontal"
    mvariIncluirHScroll = vData
End Property

Public Property Get iIncluirHScroll() As Integer
    iIncluirHScroll = mvariIncluirHScroll
End Property

Public Property Let iExecutarFuncao(ByVal vData As Integer)
Attribute iExecutarFuncao.VB_Description = "0=executa a funcao 1=nao executa funcao"
    mvariExecutarFuncao = vData
End Property

Public Property Get iExecutarFuncao() As Integer
    iExecutarFuncao = mvariExecutarFuncao
End Property

Public Property Get objOptionButtonUnchecked() As Object
    Set objOptionButtonUnchecked = mvarobjOptionButtonUnchecked
End Property

Public Property Get objOptionButtonChecked() As Object
    Set objOptionButtonChecked = mvarobjOptionButtonChecked
End Property

Public Property Get objCheckboxUnchecked() As Object
    Set objCheckboxUnchecked = mvarobjCheckboxUnchecked
End Property

Public Property Get objCheckboxChecked() As Object
    Set objCheckboxChecked = mvarobjCheckboxChecked
End Property

Public Property Get objCheckboxGrayed() As Object
    Set objCheckboxGrayed = mvarobjCheckboxGrayed
End Property

Public Property Get objButton() As Object
    Set objButton = mvarobjButton
End Property

Public Property Let iProibidoIncluir(ByVal vData As Integer)
    mvariProibidoIncluir = vData
End Property

Public Property Get iProibidoIncluir() As Integer
    iProibidoIncluir = mvariProibidoIncluir
End Property

Public Property Let iProibidoExcluir(ByVal vData As Integer)
Attribute iProibidoExcluir.VB_Description = "Indica se algum elemento do grid pode ser excluido ou não. 0=permitido excluir 1=proibido excluir"
    mvariProibidoExcluir = vData
End Property

Public Property Get iProibidoExcluir() As Integer
    iProibidoExcluir = mvariProibidoExcluir
End Property

Public Property Let iGridLargAuto(ByVal vData As Integer)
Attribute iGridLargAuto.VB_Description = "Indica se a largura do grid será gerada automaticamente pela soma da largura dos campos ou se será fornecida pelo usuário na tela.1=auto 0=manual"
    mvariGridLargAuto = vData
End Property

Public Property Get iGridLargAuto() As Integer
    iGridLargAuto = mvariGridLargAuto
End Property

Public Property Let iLinhasExistentes(ByVal vData As Integer)
    mvariLinhasExistentes = vData
    mvarobjGrid.Tag = CStr(vData)
End Property

Public Property Get iLinhasExistentes() As Integer
    iLinhasExistentes = mvariLinhasExistentes
End Property

Public Property Let iLinhasVisiveis(ByVal vData As Integer)
Attribute iLinhasVisiveis.VB_Description = "Número de linhas visiveis descontando as linhas fixas"
    mvariLinhasVisiveis = vData
End Property

Public Property Get iLinhasVisiveis() As Integer
    iLinhasVisiveis = mvariLinhasVisiveis
End Property

Public Property Set objForm(ByVal vData As Object)
Attribute objForm.VB_Description = "Tela"
    Set mvarobjForm = vData
End Property

Public Property Get objForm() As Object
    Set objForm = mvarobjForm
End Property

Public Property Get colCampo() As Collection
    Set colCampo = mvarcolCampo
End Property

Public Property Get colColuna() As Collection
Attribute colColuna.VB_Description = "Contém os textos com os títulos das colunas"
    Set colColuna = mvarcolColuna
End Property

Public Property Let objGrid(ByVal vData As Object)
  Set mvarobjGrid = vData
End Property

Public Property Get objGrid() As Object
    Set objGrid = mvarobjGrid
End Property

Public Property Set objControle(ByVal vData As Object)
    Set mvarobjControle = vData
End Property

Public Property Get objControle() As Object
    Set objControle = mvarobjControle
End Property

Public Property Let iLinhaAntiga(ByVal vData As Integer)
Attribute iLinhaAntiga.VB_Description = "guarda a ultima linha  corrente antes da atual"
    mvariLinhaAntiga = vData
End Property

Public Property Get iLinhaAntiga() As Integer
    iLinhaAntiga = mvariLinhaAntiga
End Property

Public Property Let iColunaAntiga(ByVal vData As Integer)
Attribute iColunaAntiga.VB_Description = "guarda a ultima coluna corrente antes da atual"
    mvariColunaAntiga = vData
End Property

Public Property Get iColunaAntiga() As Integer
    iColunaAntiga = mvariColunaAntiga
End Property

Public Property Let lErroSaidaCelula(ByVal vData As Integer)
Attribute lErroSaidaCelula.VB_Description = "indica que foi identificado um erro na celula que estava sendo abandonada."
    mvarlErroSaidaCelula = vData
End Property

Public Property Get lErroSaidaCelula() As Integer
    lErroSaidaCelula = mvarlErroSaidaCelula
End Property

Public Property Let iTeclado(ByVal vData As Integer)
Attribute iTeclado.VB_Description = "indica que está trabalhando com o teclado ou não."
    mvariTeclado = vData
End Property

Public Property Get iTeclado() As Integer
    iTeclado = mvariTeclado
End Property

Public Property Let iSaidaCelula(ByVal vData As Integer)
Attribute iSaidaCelula.VB_Description = "indica se a rotina de tratamento de saida de celula já foi executada"
    mvariSaidaCelula = vData
End Property

Public Property Get iSaidaCelula() As Integer
    iSaidaCelula = mvariSaidaCelula
End Property



Private Sub Class_Initialize()

Dim objAdm As New ADM

    Set mvarcolColuna = New Collection
    Set mvarcolCampo = New Collection
    Set mvarcolIndex = New Collection
    mvariLinhasExistentes = 0
    mvariLinhasVisiveis = 0
    mvariProibidoExcluir = 0
    mvariProibidoIncluir = 0
    mvariProibidoIncluirNoMeio = GRID_PROIBIDO_INCLUIR_NOMEIO
    Set mvarobjCheckboxChecked = objAdm.gobjCheckboxChecked
    Set mvarobjCheckboxUnchecked = objAdm.gobjCheckboxUnchecked
    Set mvarobjCheckboxGrayed = objAdm.gobjCheckboxGrayed
    Set mvarobjOptionButtonChecked = objAdm.gobjOptionButtonChecked
    Set mvarobjOptionButtonUnchecked = objAdm.gobjOptionButtonUnChecked
    Set mvarobjButton = objAdm.gobjButton
    
    mvariIncluirHScroll = 0
    
End Sub




