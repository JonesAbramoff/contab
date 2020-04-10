VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ComissoesRegrasOcx 
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9615
   Begin VB.Frame FrameComissoesRegras 
      Caption         =   "Regras para cálculo de comissões"
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   9375
      Begin VB.CommandButton BotaoLimpar 
         Caption         =   "Limpar Grid"
         Height          =   540
         Left            =   4680
         Picture         =   "ComissoesRegrasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   3240
         Width           =   1275
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   975
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercComissaoEmiss 
         Height          =   285
         Left            =   3705
         TabIndex        =   6
         Top             =   2040
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Format          =   "0%"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton BotaoInserirLinhas 
         Height          =   540
         Left            =   6120
         Picture         =   "ComissoesRegrasOcx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Insere uma linha no grid, acima da linha atual (Insert)."
         Top             =   3240
         Width           =   1275
      End
      Begin VB.CommandButton BotaoConsultaCampo 
         Height          =   540
         Left            =   7560
         Picture         =   "ComissoesRegrasOcx.ctx":2674
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Chama tela de consulta correspondente ao mnemônico selecionado."
         Top             =   3240
         Width           =   1275
      End
      Begin VB.CheckBox VerificaSintaxe 
         Caption         =   "Verifica Sintaxe ao Sair da Célula (F5)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   8
         Top             =   3398
         Value           =   1  'Checked
         Width           =   3600
      End
      Begin VB.TextBox Regra 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3705
         TabIndex        =   3
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox ValorBase 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3705
         TabIndex        =   4
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox PercComissao 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3705
         TabIndex        =   5
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CheckBox Indireta 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton BotaoSubirRegra 
         Height          =   375
         Left            =   8900
         Picture         =   "ComissoesRegrasOcx.ctx":45A6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton BotaoDescerRegra 
         Height          =   375
         Left            =   8900
         Picture         =   "ComissoesRegrasOcx.ctx":4768
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1680
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid GridComissoesRegras 
         Height          =   2775
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4895
         _Version        =   393216
      End
   End
   Begin VB.Frame FrameMontaFormulas 
      Caption         =   "Monte suas fórmulas"
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   9375
      Begin VB.ComboBox Mnemonicos 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   495
         Width           =   3675
      End
      Begin VB.ComboBox Funcoes 
         Height          =   315
         ItemData        =   "ComissoesRegrasOcx.ctx":492A
         Left            =   4080
         List            =   "ComissoesRegrasOcx.ctx":492C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   495
         Width           =   3795
      End
      Begin VB.ComboBox Operadores 
         Height          =   315
         Left            =   8045
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   495
         Width           =   1150
      End
      Begin VB.TextBox Descricao 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   540
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   900
         Width           =   8955
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Operadores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8060
         TabIndex        =   22
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label LabelFuncoes 
         AutoSize        =   -1  'True
         Caption         =   "Funções:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4110
         TabIndex        =   21
         Top             =   240
         Width           =   795
      End
      Begin VB.Label LabelMnemonicos 
         AutoSize        =   -1  'True
         Caption         =   "Mnemônicos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7800
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1665
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "ComissoesRegrasOcx.ctx":492E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "ComissoesRegrasOcx.ctx":4AB8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ComissoesRegrasOcx.ctx":4C36
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "ComissoesRegrasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Iniciado por tulio em 15/04
'Supervisionado por Luiz G.F. Nogueira

'Property Variables:
Dim m_Caption As String
Event Unload()

Const NUM_MAX_REGRAS_COMISSOES = 1000

Const KEYCODE_VERIFICASINTAXE = 116

Const GRID_SUBIR_LINHA = "U"
Const GRID_DESCER_LINHA = "D"


'***************************************************
'Declaracoes Globais
'***************************************************

'Guardar o ultimo SelStart, para resolver o problema da chamada
'de browser via botao consulta
Dim iRegraSelStart As Integer
Dim iValorBaseSelStart As Integer
Dim iPercComissaoSelStart As Integer


'Propriedade iAlterado da tela
Dim iAlterado As Integer

'obj Gerenciador do Grid da tela
Public objGridRegras As AdmGrid

'Eventos browser
Private WithEvents objEventoBrowser As AdmEvento
Attribute objEventoBrowser.VB_VarHelpID = -1

'Private WithEvents objBotao As Button

'Colecao de mnemonicos
Dim colMnemonicos As Collection

'Variaveis de controle que representam as colunas do grid
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_PercComissao_Col As Integer
Dim iGrid_Regra_Col As Integer
Dim iGrid_ValorBase_Col As Integer
Dim iGrid_Indireta_Col As Integer
Dim iGrid_PercComissaoEmiss_Col As Integer
'***************************************************
'Fim Declaracoes Globais
'***************************************************

'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Regras para cálculo de comissões"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "ComissoesRegras"
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
    RaiseEvent Unload
End Sub

Public Sub Form_Activate()
   'Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    'gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub
'***************************************************
'Fim Trecho de codigo comum as telas
'***************************************************

Public Function Trata_Parametros() As Long
'Trata os parametros passados para a tela..
'No caso, so retorna sucesso....
    
    Trata_Parametros = SUCESSO

End Function

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Instancia o objeto que gerencia o Grid
    Set objGridRegras = New AdmGrid
    
    'Instancia a colecao global da tela (mnemonicos)
    Set colMnemonicos = New Collection
    
    'Inicializa os Eventos
    Set objEventoBrowser = New AdmEvento
        
    'Executa inicializacao do Grid
    lErro = Inicializa_GridComissoesRegras(objGridRegras)
    If lErro <> SUCESSO Then gError 101541

    'Carrega as combos da tela
    lErro = Carrega_Combos_Tela()
    If lErro <> SUCESSO Then gError 101540

    'Preenche o grid com as regras cadastradas
    lErro = Preenche_GridRegras()
    If lErro <> SUCESSO Then gError 101542
    
    'utilizacao das tags para guardar o selstart do campo
    'inicializacao com valor default 1...
'    Regra.Tag = 1
'    ValorBase.Tag = 1
'    PercComissao.Tag = 1
    iRegraSelStart = 1
    iValorBaseSelStart = 1
    iPercComissaoSelStart = 1
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 101540, 101541, 101542
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154429)

    End Select

    Exit Sub

End Sub

Private Function Traz_ComissoesRegras_Tela(ByVal colComissoesRegras As Collection) As Long
'Recebe a colecao carregada.
'coloca as informacoes na tela.
'a colecao representa o grid.
'colComissoesRegras eh parametro de INPUT

Dim lErro As Long
Dim objComissoesRegras As ClassComissoesRegras
Dim iLinha As Integer

On Error GoTo Erro_Traz_ComissoesRegras_Tela

    'inicializa o numero de linhas
    iLinha = 1

    'Para cada regra na colecao
    For Each objComissoesRegras In colComissoesRegras
        
        'chama a funcao responsavel por trazer
        'a regra para a linha
        lErro = Traz_Regra_Tela(objComissoesRegras, iLinha)
        If lErro <> SUCESSO Then gError 101539
        
        'passa para a proxima linha
        iLinha = iLinha + 1
    
    Next
    
    'atualiza o numero de linhas existentes
    objGridRegras.iLinhasExistentes = iLinha - 1
    
    'Chama a funcao pra atualizar o bitmap de todas as checks
    Call Grid_Refresh_Checkbox(objGridRegras)
    
    Traz_ComissoesRegras_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_ComissoesRegras_Tela:

    Traz_ComissoesRegras_Tela = gErr
    
    Select Case gErr
    
        Case 101539
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154430)
        
    End Select
        
    Exit Function

End Function

Private Function Traz_Regra_Tela(ByVal objComissoesRegras As ClassComissoesRegras, ByVal iLinha As Integer) As Long
'Traz a regra do obj para a linha do grid passada como parametro
'objComissoesRegras eh parametro de INPUT
'iLinha eh parametro de INPUT
    
Dim objVendedor As ClassVendedor
Dim lErro As Long
    
On Error GoTo Erro_Traz_Regra_Tela
    
    'Coloca o Vendedor na linha caso exista
    If objComissoesRegras.iVendedor <> 0 Then
        
        Set objVendedor = New ClassVendedor
        
        objVendedor.iCodigo = objComissoesRegras.iVendedor
        
        'le o vendedor para obter o nome reduzido
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 111791
        
        If lErro <> SUCESSO Then gError 111792
            
        GridComissoesRegras.TextMatrix(iLinha, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
    
    Else
    
        GridComissoesRegras.TextMatrix(iLinha, iGrid_Vendedor_Col) = STRING_VAZIO
        
    End If
    
    'Coloca a Regra na linha
    GridComissoesRegras.TextMatrix(iLinha, iGrid_Regra_Col) = objComissoesRegras.sRegra

    'Coloca o Valor Base na linha
    GridComissoesRegras.TextMatrix(iLinha, iGrid_ValorBase_Col) = objComissoesRegras.sValorBase
    
    'Coloca o Perc Comissao na linha
    GridComissoesRegras.TextMatrix(iLinha, iGrid_PercComissao_Col) = objComissoesRegras.sPercComissao

    'Coloca o DiretoIndireto na tela
    GridComissoesRegras.TextMatrix(iLinha, iGrid_Indireta_Col) = objComissoesRegras.iVendedorIndireto
    
    'Coloca o percemissao na tela
    GridComissoesRegras.TextMatrix(iLinha, iGrid_PercComissaoEmiss_Col) = Format(objComissoesRegras.dPercComissaoEmiss, "percent")
    
    Traz_Regra_Tela = SUCESSO

    Exit Function

Erro_Traz_Regra_Tela:

    Traz_Regra_Tela = gErr
    
    Select Case gErr
        
        Case 111791
        
        Case 111792
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", objVendedor.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154431)

    End Select
    
    Exit Function

End Function

Private Function Carrega_Combos_Tela() As Long
'Responsavel pela carga das comboboxes existentes na tela..

Dim lErro As Long

On Error GoTo Erro_Carrega_Combos_Tela

    'Carrega a combo Mnemonico
    lErro = Carrega_Mnemonicos()
    If lErro <> SUCESSO Then gError 101546
    
    'Carrega a combo Funcoes
    lErro = Carrega_Funcoes()
    If lErro <> SUCESSO Then gError 101547
    
    'Carrega a combo Operadores
    lErro = Carrega_Operacoes()
    If lErro <> SUCESSO Then gError 101548
    
    Carrega_Combos_Tela = SUCESSO

    Exit Function

Erro_Carrega_Combos_Tela:
    
    Carrega_Combos_Tela = gErr
    
    Select Case gErr
    
        Case 101546 To 101548
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154432)
            
    End Select
    
    Exit Function
    
End Function

Private Function Preenche_GridRegras() As Long
'Preenche o grid com as regras do bd

Dim lErro As Long
Dim colComissoesRegras As New Collection

On Error GoTo Erro_Preenche_GridRegras

    'Le as regras que estao cadastradas no BD
    'O erro 94916, retornado quando a tabela está vazia, nao foi tratado propositalmente,
    'pois a tela deve abrir mesmo que não existam regras
    lErro = CF("ComissoesRegras_Le_Todas", colComissoesRegras)
    If lErro <> SUCESSO And lErro <> 94916 Then gError 101543

    'Exibe na tela as regras lidas do BD
    lErro = Traz_ComissoesRegras_Tela(colComissoesRegras)
    If lErro <> SUCESSO Then gError 101545
    
    Preenche_GridRegras = SUCESSO

    Exit Function

Erro_Preenche_GridRegras:

    Preenche_GridRegras = gErr

    Select Case gErr
        
        Case 101543, 101545
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154433)
            
    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(ByVal iLinha As Integer, ByVal objControl As Object, ByVal iLocalChamada As Integer)

Dim lErro As Long
Dim iIndex As Integer
Dim bValor As Boolean

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Seleciona o controle atual
    Select Case objControl.Name

        'Se for a check Indireta => habilita / desabilita o campo Vendedor,
        'pois o vendedor só deve ser preenchido quando a regra for uma regra indireta
        Case Vendedor.Name
        
            'Se estiver marcada
            If GridComissoesRegras.TextMatrix(iLinha, iGrid_Indireta_Col) = CStr(MARCADO) Then
            
                'Habilita a coluna vendedor
                Vendedor.Enabled = True

            'Senão, ou seja, se estiver desmarcada
            Else

                'Desabilita a coluna vendedor
                Vendedor.Enabled = False
            
            End If
            
            'Indica que é para manter desabilitadas as combos Mnemonicos, Funcoes e Operadores
            bValor = False
            
        'Se for Indireta
        Case Indireta.Name
            
            'Indica que é para manter desabilitadas as combos Mnemonicos, Funcoes e Operadores
            bValor = False
            
            
        'Se for PercEmissao
        Case PercComissaoEmiss.Name
        
            'Indica que eh pra manter desabilitadas as combos Mnemonicos, Funcoes e Operadores
            bValor = False
        
        'Se for qualquer outra coluna
        Case Else
        
            'Indica que é para manter habilitadas as combos Mnemonicos, Funcoes e Operadores
            bValor = True
            
    End Select
    
    'Habilita/Desabilita a combo de mnemonicos
    Mnemonicos.Enabled = bValor
    
    'Habilita/Desabilita a combo de funcoes
    Funcoes.Enabled = bValor
    
    'Habilita/Desabilita a combo de operadores
    Operadores.Enabled = bValor
    
    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154434)
            
    End Select
    
    Exit Sub

End Sub

Private Function Inicializa_GridComissoesRegras(ByVal objGridInt As AdmGrid) As Long
'Inicializa o grid da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridComissoesRegras

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Indireta")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("Regra")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Fórmulas para Percentual Comissão")
    objGridInt.colColuna.Add ("Percentual Emissão")
    
    'campos de edição do grid
    objGridInt.colCampo.Add (Indireta.Name)
    objGridInt.colCampo.Add (Vendedor.Name)
    objGridInt.colCampo.Add (Regra.Name)
    objGridInt.colCampo.Add (ValorBase.Name)
    objGridInt.colCampo.Add (PercComissao.Name)
    objGridInt.colCampo.Add (PercComissaoEmiss.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_Indireta_Col = 1
    iGrid_Vendedor_Col = 2
    iGrid_Regra_Col = 3
    iGrid_ValorBase_Col = 4
    iGrid_PercComissao_Col = 5
    iGrid_PercComissaoEmiss_Col = 6
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridComissoesRegras

    'Numero Maximo de Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REGRAS_COMISSOES

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridComissoesRegras.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'seta execucao da rotina grid enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'permite insercao de linhas no meio do grid
    objGridInt.iProibidoIncluirNoMeioGrid = GRID_PERMITIDO_INCLUIR_NO_MEIO
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridComissoesRegras = SUCESSO

    Exit Function

Erro_Inicializa_GridComissoesRegras:

    Inicializa_GridComissoesRegras = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154435)
            
    End Select

    Exit Function

End Function

Private Function Carrega_Mnemonicos() As Long
'Preenche a combo de mnemonicos com o conteudo do bd

Dim lErro As Long
Dim objMnemonicoComissoes As ClassMnemonicoComissoes

On Error GoTo Erro_Carrega_Mnemonicos

    'Le os mnemonicos do bd
    lErro = CF("MnemonicoComissoes_Todos_Le_Todos", colMnemonicos)
    If lErro <> SUCESSO And lErro <> 101550 Then gError 101551
    
    'limpa a combo
    Mnemonicos.Clear
    
    'para cada mnemonico na colecao carregada anteriormente
    For Each objMnemonicoComissoes In colMnemonicos
    
        'adiciona o dito cujo na combo
        Mnemonicos.AddItem objMnemonicoComissoes.sMnemonico
        
    Next
        
    Carrega_Mnemonicos = SUCESSO

    Exit Function

Erro_Carrega_Mnemonicos:

    Carrega_Mnemonicos = gErr
    
    Select Case gErr
    
        Case 101551
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154436)
            
    End Select
    
    Exit Function

End Function

Private Function Carrega_Funcoes() As Long
'Preenche a combo de funcoes com o conteudo do bd

Dim lErro As Long
Dim objFormulaFuncao As ClassFormulaFuncao
Dim colFormulaFuncao As New Collection

On Error GoTo Erro_Carrega_Funcoes

    'Le as funcoes do bd
    lErro = CF("FormulaFuncao_Le_Todos", colFormulaFuncao)
    If lErro <> SUCESSO Then gError 101553
    
    'limpa a combo
    Funcoes.Clear
    
    'para cada funcao na colecao carregada anteriormente
    For Each objFormulaFuncao In colFormulaFuncao
    
        'adiciona o dito cujo na combo
        Funcoes.AddItem objFormulaFuncao.sFuncaoCombo
        
    Next
        
    Carrega_Funcoes = SUCESSO

    Exit Function

Erro_Carrega_Funcoes:

    Carrega_Funcoes = gErr
    
    Select Case gErr
    
        Case 101553
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154437)
            
    End Select
    
    Exit Function

End Function

Private Function Carrega_Operacoes() As Long
'Preenche a combo de operacoes com o conteudo do bd

Dim lErro As Long
Dim objFormulaOperador As ClassFormulaOperador
Dim colFormulaOperador As New Collection

On Error GoTo Erro_Carrega_Operacoes

    'Le as funcoes do bd
    lErro = CF("FormulaOperador_Le_Todos", colFormulaOperador)
    If lErro <> SUCESSO Then gError 101555
    
    'limpa a combo
    Operadores.Clear
    
    'para cada funcao na colecao carregada anteriormente
    For Each objFormulaOperador In colFormulaOperador
    
        'adiciona o dito cujo na combo
        Operadores.AddItem objFormulaOperador.sOperadorCombo
        
    Next
        
    Carrega_Operacoes = SUCESSO

    Exit Function

Erro_Carrega_Operacoes:

    Carrega_Operacoes = gErr
    
    Select Case gErr
    
        Case 101555
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154438)
            
    End Select
    
    Exit Function

End Function

Private Function ComissoesRegras_Critica() As Long
'Funcao que faz a critica da tela ComissoesRegras

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_ComissoesRegras_Critica

    'Para cada linha do grid
    For iLinha = 1 To objGridRegras.iLinhasExistentes
    
        'Se o campo indireta estiver marcado..
        If GridComissoesRegras.TextMatrix(iLinha, iGrid_Indireta_Col) = CStr(vbChecked) Then
    
            'se o vendedor nao estiver preenchido => erro
            If Len(Trim(GridComissoesRegras.TextMatrix(iLinha, iGrid_Vendedor_Col))) = 0 Then gError 101571
            
        End If
    
        'Se o campo regra estiver preenchido
        If Len(Trim(GridComissoesRegras.TextMatrix(iLinha, iGrid_Regra_Col))) > 0 Then
        
            'Validar as regras
            lErro = CF("Valida_Formula_Comissoes", GridComissoesRegras.TextMatrix(iLinha, iGrid_Regra_Col), colMnemonicos, TIPO_BOOLEANO)
            If lErro <> SUCESSO Then gError 101579
        
        End If
        
        'Se o campo Valor base estiver preenchido
        If Len(Trim(GridComissoesRegras.TextMatrix(iLinha, iGrid_ValorBase_Col))) > 0 Then
        
            'Validar as regras
            lErro = CF("Valida_Formula_Comissoes", GridComissoesRegras.TextMatrix(iLinha, iGrid_ValorBase_Col), colMnemonicos, TIPO_NUMERICO)
            If lErro <> SUCESSO Then gError 101580
        
        End If
        
        'Se o campo PercComissao estiver preenchido
        If Len(Trim(GridComissoesRegras.TextMatrix(iLinha, iGrid_PercComissao_Col))) > 0 Then
        
            'Validar as regras
            lErro = CF("Valida_Formula_Comissoes", GridComissoesRegras.TextMatrix(iLinha, iGrid_PercComissao_Col), colMnemonicos, TIPO_NUMERICO)
            If lErro <> SUCESSO Then gError 101581
            
        End If
    
        'se percentual de comissao na emissao nao estiver preenchido
        If Len(Trim(GridComissoesRegras.TextMatrix(iLinha, iGrid_PercComissaoEmiss_Col))) = 0 Then gError 101631
        
    
    Next
    
    ComissoesRegras_Critica = SUCESSO
    
    Exit Function
    
Erro_ComissoesRegras_Critica:

    ComissoesRegras_Critica = gErr

    Select Case gErr
    
        Case 101571
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_PREENCHIDO_INDIRETA_MARCADA", gErr, iLinha)
    
        Case 101579 To 101581
        
        Case 101631
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCCOMISSAOEMISS_NAO_PREENCHIDO", gErr, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154439)
            
    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a funcao que ira efetuar a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 101573

    iAlterado = 0

    'fecha a tela apos a gravacao (padrao nas telas de configuracao)
    Call BotaoFechar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 101573

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154440)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava o registro.. deve ser sempre publica pois sera chamada
'de fora ...

Dim lErro As Long
Dim iIndice As Integer
Dim colComissoesRegras As New Collection

On Error GoTo Erro_Gravar_Registro

     'Exibe uma ampulheta como ponteiro do mouse
     'para que o usuario tenha o feedback da gravacao
     GL_objMDIForm.MousePointer = vbHourglass

     'Critica os Dados que serao gravados
     lErro = ComissoesRegras_Critica()
     If lErro <> SUCESSO Then gError 101574
          
     'Guarda os dados presentes na tela na colecao de objetos de ClassComissoesRegras..
     lErro = Move_Tela_Memoria(colComissoesRegras)
     If lErro <> SUCESSO Then gError 101575

     'grava as regras de comissoes no BD
     lErro = CF("ComissoesRegras_Grava", colComissoesRegras)
     If lErro <> SUCESSO Then gError 101577

     'fechando comando de setas
     Call ComandoSeta_Fechar(Me.Name)

     'Exibe o ponteiro padrão do mouse
     GL_objMDIForm.MousePointer = vbDefault

     Gravar_Registro = SUCESSO
     
     Exit Function

Erro_Gravar_Registro:

    'Exibe o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 101574 To 101577
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154441)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colComissoesRegras As Collection) As Long
'Carrega em objComissoesRegras os dados da tela
'objComissoesRegras eh parametro de OUTPUT

Dim lErro As Long
Dim objComissoesRegras As ClassComissoesRegras

Dim objVendedor As ClassVendedor

Dim iLinha As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'para cada linha do grid
    For iLinha = 1 To objGridRegras.iLinhasExistentes
    
        'Instancia uma nova area de memoria a ser apontada pelo obj
        Set objComissoesRegras = New ClassComissoesRegras
    
        'se a indireta estiver marcada
        If GridComissoesRegras.TextMatrix(iLinha, iGrid_Indireta_Col) = CStr(vbChecked) Then
            
            'guarda a informacao de que eh indireto
            objComissoesRegras.iVendedorIndireto = VENDEDOR_INDIRETO
            
        Else
        'senao
        
            'guarda a informacao de que eh direto
            objComissoesRegras.iVendedorIndireto = VENDEDOR_DIRETO
        
        End If
        
        'se vendedor estiver preenchido
        If Len(Trim(GridComissoesRegras.TextMatrix(iLinha, iGrid_Vendedor_Col))) > 0 Then
            
            'coloca o codigo do vendedor no obj
            If IsNumeric(GridComissoesRegras.TextMatrix(iLinha, iGrid_Vendedor_Col)) Then
                objComissoesRegras.iVendedor = StrParaInt(GridComissoesRegras.TextMatrix(iLinha, iGrid_Vendedor_Col))
            Else
                Set objVendedor = New ClassVendedor
                
                objVendedor.sNomeReduzido = GridComissoesRegras.TextMatrix(iLinha, iGrid_Vendedor_Col)
                
                lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
                If lErro <> SUCESSO And lErro <> 25008 Then gError 111787
        
                'nesse caso, sempre vai encontrar vendedor...
                'isso eh garantido pelo saida_celula
                
                objComissoesRegras.iVendedor = objVendedor.iCodigo
            End If
        
        End If
                
        'Guarda os campos que ja foram previamente compilados..
        objComissoesRegras.sRegra = GridComissoesRegras.TextMatrix(iLinha, iGrid_Regra_Col)
        objComissoesRegras.sPercComissao = GridComissoesRegras.TextMatrix(iLinha, iGrid_PercComissao_Col)
        objComissoesRegras.sValorBase = GridComissoesRegras.TextMatrix(iLinha, iGrid_ValorBase_Col)
               
        'guarda o percentual de emissao no obj
        objComissoesRegras.dPercComissaoEmiss = PercentParaDbl(GridComissoesRegras.TextMatrix(iLinha, iGrid_PercComissaoEmiss_Col))
            
        'guarda a ordenacao no obj (a ordenacao eh a linha)
        objComissoesRegras.lOrdenacao = iLinha
        
        'adiciona o obj na colecao
        colComissoesRegras.Add objComissoesRegras
    
    Next
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154442)
            
    End Select
    
    Exit Function

End Function

Private Sub Limpa_GridRegras()
'Limpa o grid e marca a check verifica sintaxe...

    'chama a grid limpa (limpa o grid de regras)
    Call Grid_Limpa(objGridRegras)
    
    'Marca check de VerificaSintaxe
    VerificaSintaxe.Value = vbChecked
    
    iAlterado = 0

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'chama a teste_salva
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 101594

    'Limpa a tela
    Call Limpa_GridRegras

    'Fecha Comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 101594

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154443)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Exibe uma ampulheta como ponteiro do mouse
    GL_objMDIForm.MousePointer = vbHourglass

    'Pede confirmação para exclusão ao usuário
    If Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_COMISSOESREGRAS") = vbYes Then

        'Exclui as regras do bd
        lErro = CF("ComissoesRegras_Exclui")
        If lErro <> SUCESSO Then gError 101595

        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)

        'Limpa a Tela
        Call Limpa_GridRegras

    End If

    'Exibe o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'Exibe o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 101595
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154444)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

'*****************************************
'9 eventos do grid que devem ser tratados
'
'
'*****************************************

Public Sub GridComissoesRegras_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRegras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegras, iAlterado)
    End If

End Sub

Public Sub GridComissoesRegras_GotFocus()
    Call Grid_Recebe_Foco(objGridRegras)
End Sub

Public Sub GridComissoesRegras_EnterCell()
    Call Grid_Entrada_Celula(objGridRegras, iAlterado)
End Sub

Public Sub GridComissoesRegras_LeaveCell()
    Call Saida_Celula(objGridRegras)
End Sub

Public Sub GridComissoesRegras_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridRegras)
End Sub

Public Sub GridComissoesRegras_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRegras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegras, iAlterado)
    End If

End Sub

Public Sub GridComissoesRegras_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridRegras)

End Sub

Public Sub GridComissoesRegras_RowColChange()
    
    Call Grid_RowColChange(objGridRegras)
    
End Sub

Public Sub GridComissoesRegras_Scroll()
    
    Call Grid_Scroll(objGridRegras)

End Sub

'************************************************
'fim dos 9 eventos do grid que devem ser tratados
'
'
'************************************************

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'inicializa a saida
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
        
        'Verifica qual celula do grid esta deixando
        'de ser a corrente para chamar a funcao de
        'saida celula adequada...
        Select Case objGridInt.objGrid.Col

            'se for a celula de Indireta
            Case iGrid_Indireta_Col
        
'                lErro = Saida_Celula_Indireta(objGridInt)
'                If lErro <> SUCESSO Then gError 101598
            
            'se for a celula de PercComissao
            Case iGrid_PercComissao_Col
                
                lErro = Saida_Celula_PercComissao(objGridInt)
                If lErro <> SUCESSO Then gError 101599
                
            'se for a celula de PercComissaoEmiss
            Case iGrid_PercComissaoEmiss_Col
                
                lErro = Saida_Celula_PercComissaoEmiss(objGridInt)
                If lErro <> SUCESSO Then gError 101600
                
            'se for a celula de Regra
            Case iGrid_Regra_Col
                
                lErro = Saida_Celula_Regra(objGridInt)
                If lErro <> SUCESSO Then gError 101601
            
            'se for a celula de ValorBase
            Case iGrid_ValorBase_Col
                
                lErro = Saida_Celula_ValorBase(objGridInt)
                If lErro <> SUCESSO Then gError 101602
        
            'se for a celula de Vendedor
            Case iGrid_Vendedor_Col
                
                lErro = Saida_Celula_Vendedor(objGridInt)
                If lErro <> SUCESSO Then gError 101603
            
       End Select

    End If

    'finaliza a saida
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101604
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 101598 To 101603
        
        Case 101604
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154445)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Indireta(objGridInt As AdmGrid) As Long
'Faz a crítica do campo Indireta que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Indireta

    'instancia objcontrole como o controle de Indireta
    Set objGridInt.objControle = Indireta

    'se a check estiver desmarcada
    If Indireta.Value = vbUnchecked Then
        
        'limpa o campo vendedor
        GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, iGrid_Vendedor_Col) = STRING_VAZIO

    End If

    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101620

    Saida_Celula_Indireta = SUCESSO

    Exit Function

Erro_Saida_Celula_Indireta:

    Saida_Celula_Indireta = gErr

    Select Case gErr

        Case 101620
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154446)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'Faz a crítica do campo Vendedor que está deixando de ser o campo corrente

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Saida_Celula_Vendedor

    'instancia objcontrole como o controle Vendedor
    Set objGridInt.objControle = Vendedor

    'se o campo vendedor estiver preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then
    
        'se o conteudo do campo for numerico, trata-se do codigo
        If IsNumeric(Vendedor.Text) = True Then
        
            'Coloca o codigo do vendedor no objvendedor
            objVendedor.iCodigo = StrParaInt(Vendedor.Text)
        
            'Realiza a leitura do vendedor visando valida-lo
            lErro = CF("Vendedor_Le", objVendedor)
            If lErro <> SUCESSO And lErro <> 12582 Then gError 101606
        
            'se nao achou vendedor => erro
            If lErro <> SUCESSO Then gError 101608
            
            'coloca no controle o nomereduzido do vendedor lido
            Vendedor.Text = objVendedor.sNomeReduzido
        
        'senao, trata-se do nomereduzido
        Else
        
            'Coloca o nomereduzido do vendedor no objvendedor
            objVendedor.sNomeReduzido = Vendedor.Text
        
            'Realiza a leitura do vendedor visando valida-lo
            lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
            If lErro <> SUCESSO And lErro <> 25008 Then gError 111787
        
            'se nao achou vendedor => erro
            If lErro <> SUCESSO Then gError 111788
    
        End If
    
    End If
    
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101607

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = gErr

    Select Case gErr

        Case 101606, 101607, 111787
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 101608
            'CASO DO CODIGO
            'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
            If Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR1", objVendedor.iCodigo) = vbYes Then
                'Chama tela de Vendedores
                lErro = Chama_Tela("Vendedores", objVendedor)
            End If
            
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111788
            'CASO DO NOME REDUZIDO
            'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
            If Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR2", objVendedor.sNomeReduzido) = vbYes Then
                'Chama tela de Vendedores
                lErro = Chama_Tela("Vendedores", objVendedor)
            End If
            
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154447)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercComissaoEmiss(objGridInt As AdmGrid) As Long
'Faz a crítica do campo PercComissaoEmiss que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PercComissaoEmiss

    'instancia objcontrole como o controle de PercComissaoEmiss
    Set objGridInt.objControle = PercComissaoEmiss

    'Se o campo está preenchido
    If Len(Trim(PercComissaoEmiss.Text)) > 0 Then
    
        'critica o valor digitado
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 101609
            
        'aplica o formato de porcentagem
        objGridInt.objControle.Text = PercentParaDbl(objGridInt.objControle.FormattedText)
            
    End If
        
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101610

    Saida_Celula_PercComissaoEmiss = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComissaoEmiss:

    Saida_Celula_PercComissaoEmiss = gErr

    Select Case gErr

        Case 101609, 101610
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154448)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercComissao(objGridInt As AdmGrid) As Long
'Faz a crítica do campo PercComissao que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PercComissao

    'instancia objcontrole como o controle de PercComissao
    Set objGridInt.objControle = PercComissao

    'Se o campo está preenchido
    If Len(Trim(PercComissao.Text)) > 0 And VerificaSintaxe.Value = vbChecked Then
    
        'Valida o perc comissao
        lErro = CF("Valida_Formula_Comissoes", PercComissao.Text, colMnemonicos, TIPO_NUMERICO)
        If lErro <> SUCESSO Then gError 101613
            
    End If
        
    IIf PercComissao.SelStart > 0, iPercComissaoSelStart = PercComissao.SelStart, iPercComissaoSelStart = 1
        
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101612

    Saida_Celula_PercComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComissao:

    Saida_Celula_PercComissao = gErr

    Select Case gErr

        Case 101612, 101613
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154449)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorBase(objGridInt As AdmGrid) As Long
'Faz a crítica do campo ValorBase que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ValorBase

    'instancia objcontrole como o controle de valorbase
    Set objGridInt.objControle = ValorBase

    'Se o campo está preenchido
    If Len(Trim(ValorBase.Text)) > 0 And VerificaSintaxe.Value = vbChecked Then
    
        'Valida o valorbase
        lErro = CF("Valida_Formula_Comissoes", ValorBase.Text, colMnemonicos, TIPO_NUMERICO)
        If lErro <> SUCESSO Then gError 101614
            
    End If
        
    If ValorBase.SelStart > 0 Then
        iValorBaseSelStart = ValorBase.SelStart
    Else
        ValorBase.Tag = 1
    End If
        
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101615

    Saida_Celula_ValorBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBase:

    Saida_Celula_ValorBase = gErr

    Select Case gErr

        Case 101614, 101615
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154450)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Regra(objGridInt As AdmGrid) As Long
'Faz a crítica do campo ValorBase que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Regra

    'instancia objcontrole como o controle de regra
    Set objGridInt.objControle = Regra

    'Se o campo está preenchido
    If Len(Trim(Regra.Text)) > 0 And VerificaSintaxe.Value = vbChecked Then
    
        'Valida a regra
        lErro = CF("Valida_Formula_Comissoes", Regra.Text, colMnemonicos, TIPO_BOOLEANO)
        If lErro <> SUCESSO Then gError 101616
            
    End If
    
    If Regra.SelStart > 0 Then
        iRegraSelStart = Regra.SelStart
    Else
        iRegraSelStart = 1
    End If
        
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 101617

    'adiciona a proxima linha se for o caso...
    Call Adiciona_Linha_Seguinte
    
    Saida_Celula_Regra = SUCESSO

    Exit Function

Erro_Saida_Celula_Regra:

    Saida_Celula_Regra = gErr

    Select Case gErr

        Case 101616, 101617
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154451)

    End Select

    Exit Function

End Function

Private Sub Adiciona_Linha_Seguinte()
'adiciona uma linha se a linha corrente for a ultima e se a coluna corrente estiver preenchida
'deve ser chamada apos o abandona_celula

    'se for ultima linha do grid habilitada e o campo estiver preenchido
    If GridComissoesRegras.Row - GridComissoesRegras.FixedRows = objGridRegras.iLinhasExistentes And Len(Trim(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, iGrid_Regra_Col))) > 0 Then
        
        'inclui a proxima linha
        objGridRegras.iLinhasExistentes = objGridRegras.iLinhasExistentes + 1

    End If

End Sub

'********** CONTROLES DO GRID, 4 EVENTOS

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'-----------------***********************

Private Sub Regra_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Regra_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub Regra_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub Regra_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Regra
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'-----------------***********************
Private Sub ValorBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'-----------------***********************
Private Sub PercComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub PercComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub PercComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = PercComissao
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'-------------------**********************
Private Sub PercComissaoEmiss_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercComissaoEmiss_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub PercComissaoEmiss_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub PercComissaoEmiss_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = PercComissaoEmiss
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'-------------------**********************
Private Sub Indireta_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Indireta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub Indireta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub Indireta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Indireta
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'-------------------**********************

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Se a tecla pressionada foi a tecla de chamada de browser
    If KeyCode = KEYCODE_BROWSER Then
        
        'se controle ativo eh vendedor
        If Me.ActiveControl Is Vendedor Then
            
            'chama o browser de vendedor
            Call Chama_Browser_Vendedor
            
        ElseIf Me.ActiveControl Is Regra Or Me.ActiveControl Is ValorBase Or Me.ActiveControl Is PercComissao Then

            'Analisa o campo para chamar o browser adequado se for o caso
            Call Formula_Analisa(Me.ActiveControl.Text, IIf(Me.ActiveControl.SelStart > 0, Me.ActiveControl.SelStart, 1))

        End If
        
    End If
        
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    'se a tecla pressionada for a de verificar sintaxe
    If KeyCode = KEYCODE_VERIFICASINTAXE Then
    
        'troca o valor da check
        VerificaSintaxe.Value = 1 - VerificaSintaxe.Value
        
    'se for pressionada a tecla de subir linha
    ElseIf KeyCode = Asc(GRID_SUBIR_LINHA) And Shift = vbCtrlMask Then
    
        'chama o evento do botao que sobe a linha
        Call BotaoSubirRegra_Click
    
    
    'se for pressionada a tecla de descer linha
    ElseIf KeyCode = Asc(GRID_DESCER_LINHA) And Shift = vbCtrlMask Then
    
        'chama o evento do botao que desce a linha
        Call BotaoDescerRegra_Click
    
    End If
        
End Sub

Public Sub Chama_Browser_Vendedor()
'funcao responsavel por chamar o browser de vendedor...
'feita para ser chamada a partir do usercontrol_keydown
 
Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    'se o vendedor estiver preenchido guarda o codigo do vendedor no objvendedor
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        'se o conteudo do campo for numerico
        If IsNumeric(Vendedor.Text) = True Then
            'guarda o codigo
            objVendedor.iCodigo = StrParaInt(Vendedor.Text)
        'senao, guarda o nomereduzido
        Else
            objVendedor.sNomeReduzido = Vendedor.Text
        End If
    
    End If
    
    'chama o browser de vendedor
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoBrowser)

End Sub

Private Sub objEventoBrowser_evSelecao(obj1 As Object)

Dim objBrowser As Object
Dim lErro As Long
Dim sValorCampo As String
Dim sProdutoMascarado As String
Dim sProperty As String
Dim objMnemonico As ClassMnemonicoComissoes

On Error GoTo Erro_objEventoBrowser_evSelecao

    'faz com que o ponteiro objBrowser
    'aponte para obj1
    Set objBrowser = obj1

    'se o obj retornado for da classe vendedor e o controle com o foco for a textbox vendedor
    If TypeName(objBrowser) = CLASSE_VENDEDOR And (Me.ActiveControl.Name = Vendedor.Name Or Me.ActiveControl.Name = BotaoConsultaCampo.Name) Then
        
        'coloca o codigo na tela
        Vendedor.Text = objBrowser.sNomeReduzido
        GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, iGrid_Vendedor_Col) = objBrowser.sNomeReduzido
    
    Else
    
        'Para cada mnemônico na coleção
        For Each objMnemonico In colMnemonicos
            'Se a classe do mnemônico for a mesma classe do browser => guarda o nome da property que deve ser usada para obter o conteúdo a ser exibido no grid
            If objMnemonico.sClasseBrowser = TypeName(objBrowser) Then sProperty = objMnemonico.sPropertyBrowser
        Next
            
        'extrai do browser generico o que deve ir pra tela (q tbm eh generico)
        sValorCampo = CallByName(objBrowser, sProperty, VbGet)
        
        'Incluído por Luiz em 07/02/03
        'Se a classe retornada pelo browser for classe de produto
        If TypeName(objBrowser) = NOME_CLASSE_PRODUTO Then
        
            'Mascara o código do produto
            lErro = Mascara_MascararProduto(sValorCampo, sProdutoMascarado)
            If lErro <> SUCESSO Then Error 111785
            
            sValorCampo = Chr(34) & sProdutoMascarado & Chr(34)
        
        End If

        'select case para verificar em qual coluna deve jogar o valor q vai pra tela
        Select Case GridComissoesRegras.Col

            Case iGrid_Regra_Col
                'se for regra, joga no controle regra
                Call Posiciona_Texto_Tela(Regra, sValorCampo)

            Case iGrid_PercComissao_Col
                'se for percomissao, joga no controle perccomissao
                Call Posiciona_Texto_Tela(PercComissao, sValorCampo)

            Case iGrid_ValorBase_Col
                'se for valorbase, joga no controle valorbase
                Call Posiciona_Texto_Tela(ValorBase, sValorCampo)

        End Select
    
    End If
    
    'exibe a tela... (para ficar na frente do browser)...
    Me.Show
    
    Exit Sub

Erro_objEventoBrowser_evSelecao:

    Select Case gErr
        
        Case 111785

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154452)

    End Select

    Exit Sub

End Sub

Private Sub Funcoes_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Funcoes_Click()

Dim lErro As Long
Dim objFormulaFuncao As New ClassFormulaFuncao

On Error GoTo Erro_Funcoes_Click

    'se a combo nao estiver preenchida, sai
    If Len(Trim(Funcoes.Text)) = 0 Then Exit Sub
    
    'guarda a informacao da combo no obj
    objFormulaFuncao.sFuncaoCombo = Funcoes.Text
    
    'le a formula visando obter a descricao
    lErro = CF("FormulaFuncao_Le", objFormulaFuncao)
    If lErro <> SUCESSO And lErro <> 36088 Then gError 101625
    
    'se nao achou => erro (possivel exclusao da formula funcao durante a utilizacao da tela)
    If lErro <> SUCESSO Then gError 101626
    
    'coloca a descricao no campo adequado
    Descricao.Text = objFormulaFuncao.sFuncaoDesc
    
    'copia o conteudo da combo para o grid se for o caso
    Call Posiciona_Combo
    
    Exit Sub

Erro_Funcoes_Click:

    Select Case gErr
    
        Case 101625
    
        Case 101626
            Call Rotina_Erro(vbOKOnly, "ERRO_FUNCAO_NAO_CADASTRADA", gErr, objFormulaFuncao.sFuncaoCombo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154453)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Mnemonicos_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Mnemonicos_Click()

Dim lErro As Long
Dim objMnemonicoComissoes As ClassMnemonicoComissoes

On Error GoTo Erro_Mnemonicos_Click

    'se a combo nao estiver preenchida, sai
    If Len(Trim(Mnemonicos.Text)) = 0 Then Exit Sub
    
    'faz com q o obj aponte para o item da colecao referenciado pela combo..
    Set objMnemonicoComissoes = colMnemonicos(Mnemonicos.ListIndex + 1)
    
    'coloca a descricao no devido lugar (campo descricao da tela..)
    Descricao.Text = objMnemonicoComissoes.sDescricao
    
    'copia o conteudo da combo para o grid se for o caso
    Call Posiciona_Combo
    
    Exit Sub

Erro_Mnemonicos_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154454)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Operadores_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Operadores_Click()

Dim lErro As Long
Dim objFormulaOperador As New ClassFormulaOperador

On Error GoTo Erro_Operadores_Click

    'se a combo nao estiver preenchida, sai
    If Len(Trim(Operadores.Text)) = 0 Then Exit Sub
    
    'guarda a informacao da combo no obj
    objFormulaOperador.sOperadorCombo = Operadores.Text
    
    'le o operador visando obter a descricao
    lErro = CF("FormulaOperador_Le", objFormulaOperador)
    If lErro <> SUCESSO And lErro <> 36098 Then gError 101627
    
    'se nao achou => erro (possivel exclusao da formula operador durante a utilizacao da tela)
    If lErro <> SUCESSO Then gError 101628
    
    'coloca a descricao no campo adequado
    Descricao.Text = objFormulaOperador.sOperadorDesc
    
    'copia o conteudo da combo para o grid se for o caso
    Call Posiciona_Combo
    
    Exit Sub

Erro_Operadores_Click:

    Select Case gErr
    
        Case 101627
    
        Case 101628
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_CADASTRADO", gErr, objFormulaOperador.sOperadorCombo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154455)
            
    End Select
    
    Exit Sub

End Sub

Function Posiciona_Combo() As Long
'Coloca o texto no grid, caso alguma coluna e linha esteja selecionada
'a coluna deve ser uma coluna passivel de receber um operador/funcao/mnemonico

    'se existe linha e coluna selecionada
    If GridComissoesRegras.Row > 0 And GridComissoesRegras.Row <= objGridRegras.iLinhasExistentes + 1 And GridComissoesRegras.Col > 0 Then

        'seleciona a coluna
        Select Case GridComissoesRegras.Col

            'se for a coluna de percentual comissao
            Case iGrid_PercComissao_Col
                Call Posiciona_Texto_Tela(PercComissao, Me.ActiveControl.Text)

            'se for a coluna de regra
            Case iGrid_Regra_Col
                Call Posiciona_Texto_Tela(Regra, Me.ActiveControl.Text)

            'se for a coluna de valor base
            Case iGrid_ValorBase_Col
                Call Posiciona_Texto_Tela(ValorBase, Me.ActiveControl.Text)

        End Select

    End If

End Function

Private Sub Posiciona_Texto_Tela(objControl As Control, sTexto As String)
'posiciona o texto sTexto no controle objControl da tela
'funcao adaptada da contabilidade

'Função alterada por Luiz em 07/02/03

Dim iPos As Integer
Dim iTamanho As Integer
Dim sTextoEsq As String
Dim sTextoDir As String

On Error GoTo Erro_Posiciona_Texto_Tela

    'Guarda a posição onde deve ser inserido o conteúdo retornado pelo browser, ou seja, no primeiro espaço vazio à direita do texto onde se encontra o cursor
    iPos = InStr(IIf(objControl.SelStart > 0, objControl.SelStart, 1), objControl.Text, " ")
    
    'Se não encontrou espaço vazio à direita do texto onde se encontra o cursor => coloca o texto retornado pelo browser imediatamenta à direita do texto onde se encontra o cursor
    If iPos = 0 Then iPos = Len(Trim(objControl.Text))
    
    'Guarda o texto posicionado à esquerda do texto a ser inserido
    sTextoEsq = Trim(Mid(objControl.Text, 1, iPos))
    
    'Guarda o texto posicionado à direitsa do texto a ser inserido
    sTextoDir = Trim(Mid(objControl.Text, iPos + 1, Len(objControl.Text)))
    
    'Se os dois últimos caracteres não forem sinais de comparação(>=, <=, <>, =, >, <) => insere um sinal de igualdade
    If Right(Trim(sTextoEsq), 2) <> OPERADOR_MAIORIGUAL And Right(Trim(sTextoEsq), 2) <> OPERADOR_MENORIGUAL And Right(Trim(sTextoEsq), 2) <> OPERADOR_DIFERENTE And Right(Trim(sTextoEsq), 1) <> OPERADOR_IGUAL And Right(Trim(sTextoEsq), 1) <> OPERADOR_MAIOR And Right(Trim(sTextoEsq), 1) <> OPERADOR_MENOR Then sTextoEsq = sTextoEsq & " " & OPERADOR_IGUAL & " "
    
    'Insere no controle o texto passado como parâmetro na posição correta
    objControl.Text = sTextoEsq & " " & sTexto & " " & sTextoDir
    
    'Atualiza a posição do cursor, posicionando-o logo após ao texto que foi inserido no controle
    objControl.SelStart = Len(sTextoEsq) + Len(sTexto) + 1
    
    'Se o controle em questão não é controle ativo => descobre a posição onde o texto será inserido no campo do grid
    If Not (Me.ActiveControl Is objControl) Then
        
        'Se a posição de inserção do texto é maior do que o texto no campo que será atualizado
        If iPos >= Len(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col)) Then
            
            'Indica que não há expressão a ser exibida após o texto que será inserido no grid
            iTamanho = 0
        
        'Senão
        Else
            
            'Guarda o tamanho da expressão que virá após o texto a ser inserido no grid
            iTamanho = Len(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col)) - iPos
        
        End If
        
        'Insere no grid o texto passado como parâmetro na posição correta
        GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col) = Mid(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col), 1, iPos) & sTexto & Mid(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col), iPos + 1, iTamanho)
        
        iAlterado = REGISTRO_ALTERADO
        
    End If
    
    Exit Sub
    
Erro_Posiciona_Texto_Tela:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154456)
    
    End Select
    
End Sub

Private Sub BotaoSubirRegra_Click()

    'se está na primeira linha do grid-> sai
    If GridComissoesRegras.Row <= GridComissoesRegras.FixedRows Then Exit Sub
    
    'se a linha que se quer mover para cima está dentro dos limites das existentes
    If GridComissoesRegras.Row <= objGridRegras.iLinhasExistentes Then
    
        'Inverte a posição da linha atual com a linha de cima
        Call Troca_Linha(GridComissoesRegras.Row, GridComissoesRegras.Row - 1)
    
    End If


End Sub

Private Sub BotaoDescerRegra_Click()

    'se está na última linha do grid-> sai
    If GridComissoesRegras.Row >= objGridRegras.iLinhasExistentes Then Exit Sub
    
    'se a linha que se quer mover para baixo está dentro dos limites das existentes
    If GridComissoesRegras.Row >= GridComissoesRegras.FixedRows Then
    
        'Inverte a posição da linha atual com a linha de cima
        Call Troca_Linha(GridComissoesRegras.Row, GridComissoesRegras.Row + 1)
    
    End If

End Sub

Private Sub Troca_Linha(iLinha1 As Integer, iLinha2 As Integer)
'Troca o conteudo de iLinha1 com o conteudo de iLinha2
'os 2 parametros sao de INPUT

Dim sColVendedor As String
Dim sColIndireta As String
Dim sColRegra As String
Dim sColValorBase As String
Dim sColPercComissao As String
Dim sColPercComissaoEmissao As String

    'Copia o conteudo da linha1 para a memoria
    sColVendedor = GridComissoesRegras.TextMatrix(iLinha1, iGrid_Vendedor_Col)
    sColIndireta = GridComissoesRegras.TextMatrix(iLinha1, iGrid_Indireta_Col)
    sColRegra = GridComissoesRegras.TextMatrix(iLinha1, iGrid_Regra_Col)
    sColValorBase = GridComissoesRegras.TextMatrix(iLinha1, iGrid_ValorBase_Col)
    sColPercComissao = GridComissoesRegras.TextMatrix(iLinha1, iGrid_PercComissao_Col)
    sColPercComissaoEmissao = GridComissoesRegras.TextMatrix(iLinha1, iGrid_PercComissaoEmiss_Col)
    
    'Copia o conteudo da linha2 para a linha1
    GridComissoesRegras.TextMatrix(iLinha1, iGrid_Vendedor_Col) = GridComissoesRegras.TextMatrix(iLinha2, iGrid_Vendedor_Col)
    GridComissoesRegras.TextMatrix(iLinha1, iGrid_Indireta_Col) = GridComissoesRegras.TextMatrix(iLinha2, iGrid_Indireta_Col)
    GridComissoesRegras.TextMatrix(iLinha1, iGrid_Regra_Col) = GridComissoesRegras.TextMatrix(iLinha2, iGrid_Regra_Col)
    GridComissoesRegras.TextMatrix(iLinha1, iGrid_ValorBase_Col) = GridComissoesRegras.TextMatrix(iLinha2, iGrid_ValorBase_Col)
    GridComissoesRegras.TextMatrix(iLinha1, iGrid_PercComissao_Col) = GridComissoesRegras.TextMatrix(iLinha2, iGrid_PercComissao_Col)
    GridComissoesRegras.TextMatrix(iLinha1, iGrid_PercComissaoEmiss_Col) = GridComissoesRegras.TextMatrix(iLinha2, iGrid_PercComissaoEmiss_Col)
    
    'copia para os controles tbm (pois a linha corrente fica sendo a linha1)
    'o campo indireta vai ser atualizado depois.. nao precisa copiar...
    Vendedor.Text = GridComissoesRegras.TextMatrix(iLinha2, iGrid_Vendedor_Col)
    Regra.Text = GridComissoesRegras.TextMatrix(iLinha2, iGrid_Regra_Col)
    ValorBase.Text = GridComissoesRegras.TextMatrix(iLinha2, iGrid_ValorBase_Col)
    PercComissao.Text = GridComissoesRegras.TextMatrix(iLinha2, iGrid_PercComissao_Col)
    PercComissaoEmiss.Text = GridComissoesRegras.TextMatrix(iLinha2, iGrid_PercComissaoEmiss_Col)
    
    'Copia o conteudo da memoria para linha2
    GridComissoesRegras.TextMatrix(iLinha2, iGrid_Vendedor_Col) = sColVendedor
    GridComissoesRegras.TextMatrix(iLinha2, iGrid_Indireta_Col) = sColIndireta
    GridComissoesRegras.TextMatrix(iLinha2, iGrid_Regra_Col) = sColRegra
    GridComissoesRegras.TextMatrix(iLinha2, iGrid_ValorBase_Col) = sColValorBase
    GridComissoesRegras.TextMatrix(iLinha2, iGrid_PercComissao_Col) = sColPercComissao
    GridComissoesRegras.TextMatrix(iLinha2, iGrid_PercComissaoEmiss_Col) = sColPercComissaoEmissao

    'preenche as checks
    Call Grid_Refresh_Checkbox(objGridRegras)
    
    'coloca a linha2 como corrente para que de a impressao de q esta carregando a linha
    GridComissoesRegras.Row = iLinha2
    
End Sub

Private Sub BotaoInserirLinhas_Click()
'insere uma linha no grid

    'coloca o foco no grid
    GridComissoesRegras.SetFocus
    
    'emula a tecla esc (o foco tende a ir para o controle)
    Call SendKeys("{ESC}", True)
    
    'aciona a tecla insert que eh a responsavel por inserir a linha no meio do grid
    Call SendKeys("{INSERT}")
    
    Call SendKeys("{ENTER}", True)
    
End Sub

Private Sub BotaoConsultaCampo_Click()

On Error GoTo Erro_BotaoConsultaCampo_Click
 
    'se nao existem linhas selecionadas => erro
    If GridComissoesRegras.Row = 0 Then gError 101633
    
    'faz uma selecao pela coluna do grid
    Select Case GridComissoesRegras.Col
    
        'se for a coluna de vendedor
        Case iGrid_Vendedor_Col
            
            'se a check indireta estiver marcada
            If GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, iGrid_Indireta_Col) = CStr(vbChecked) Then
                
                'chama o browser de cliente
                Call Chama_Browser_Vendedor
                
            End If
                
        Case iGrid_Regra_Col
            'Analisa o campo para chamar o browser adequado se for o caso
            Call Formula_Analisa(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col), iRegraSelStart)
        
        Case iGrid_ValorBase_Col
            
            'Analisa o campo para chamar o browser adequado se for o caso
            Call Formula_Analisa(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col), iValorBaseSelStart)
        
        Case iGrid_PercComissao_Col

            'Analisa o campo para chamar o browser adequado se for o caso
            Call Formula_Analisa(GridComissoesRegras.TextMatrix(GridComissoesRegras.Row, GridComissoesRegras.Col), iPercComissaoSelStart)
            
    End Select
    
    Exit Sub
    
Erro_BotaoConsultaCampo_Click:

    Select Case gErr
    
        Case 101633
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154457)
            
    End Select
    
    Exit Sub

End Sub

Private Sub VerificaSintaxe_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Formula_Analisa(sFormula As String, iPos As Integer)
'verifica qual browser deve chamar de acordo com o cursor
'sFormula eh parametro de INPUT e guarda o texto que estava no controle
'iPos eh parametro de INPUT e diz onde esta o cursor no momento da chamada do browser
'Recursiva

Dim sMnemonico As String
Dim sBrowser As String
Dim sClasse As String
Dim colSelecao As New Collection
Dim objGenerico As Object
Dim objMnemonicoComissoes As New ClassMnemonicoComissoes
Dim iPosEspacoEsq As Integer
Dim iPosEspacoDir As Integer
Static iAchouComparacao As Integer

On Error GoTo Erro_Formula_Analisa

    'Se não existe fórmula para ser analisada => sai da função
    If Len(Trim(sFormula)) = 0 Then Exit Sub
    
    'Se a posição do cursor é igual a 0 => sai da função
    If iPos = 0 Then Exit Sub
    
    'Descobre o primeiro espaço vazio à direita do texto onde se encontra o cursor
    iPosEspacoDir = IIf(InStr(iPos, sFormula, " ") > 0, InStr(iPos, sFormula, " "), Len(Trim(sFormula)) + 1)
    
    'Descobre o primeiro espaço vazio à esquerda do texto onde se encontra o cursor
    iPosEspacoEsq = IIf(InStrRev(sFormula, " ", iPos) > 0, InStrRev(sFormula, " ", iPos), 1)
    
    'Pega o conteúdo do mnemônico, ou seja, o texto entre o espaço à esquerda e o espaço à direita
    sMnemonico = Mid(sFormula, iPosEspacoEsq, iPosEspacoDir - iPosEspacoEsq)
    
    'Retira espaços vazios e guarda no obj o mnemônico que será pesquisado
    objMnemonicoComissoes.sMnemonico = Trim(sMnemonico)

    'Se os dois últimos caracteres forem sinais de comparação maior igual, menor igual ou diferente => 'retira os sinais para que o mnemônico possa ser pesquisado
    If Right(objMnemonicoComissoes.sMnemonico, 2) = OPERADOR_MAIORIGUAL Or Right(objMnemonicoComissoes.sMnemonico, 2) = OPERADOR_MENORIGUAL Or Right(objMnemonicoComissoes.sMnemonico, 2) = OPERADOR_DIFERENTE Then objMnemonicoComissoes.sMnemonico = Mid(objMnemonicoComissoes.sMnemonico, 1, Len(objMnemonicoComissoes.sMnemonico) - 2)
    
    'Se o último caracter for um sinal de igual, maior ou menor =>         'retira o sinal para que o mnemônico possa ser pesquisado
    If Right(objMnemonicoComissoes.sMnemonico, 1) = OPERADOR_IGUAL Or Right(objMnemonicoComissoes.sMnemonico, 1) = ">" Or Right(objMnemonicoComissoes.sMnemonico, 1) = "<" Then objMnemonicoComissoes.sMnemonico = Mid(objMnemonicoComissoes.sMnemonico, 1, Len(objMnemonicoComissoes.sMnemonico) - 1)
    
    'Se não sobrou mnemônico após retirar o sinal de comparação
    If Len(Trim(objMnemonicoComissoes.sMnemonico)) = 0 Then
    
        'Posiciona o cursor à esquerda do sinal de comparação
        iPos = iPosEspacoEsq - 1
        
        'Se é para continuar procurando um mnemônico válido => verifica se o texto à esquerda do sinal de comparação é um mnemônico com browser
        Call Formula_Analisa(sFormula, iPos)
        
        'Sai da função
        Exit Sub
    
    End If
    
    'verifica se o mnemônico é válido
    If VerificaMnemonico(objMnemonicoComissoes) = True Then

        'instanciar o obj declarado acima como a classe do browser (retornar a mesma..)
        Set objGenerico = CreateObject(objMnemonicoComissoes.sProjetoBrowser & "." & objMnemonicoComissoes.sClasseBrowser)
        
        'chama o browser
        Call Chama_Tela(objMnemonicoComissoes.sNomeBrowser, colSelecao, objGenerico, objEventoBrowser)
        Exit Sub

    End If
    
    Exit Sub

Erro_Formula_Analisa:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154458)
    
    End Select
    
End Sub

Private Function VerificaMnemonico(objMnemonicoComissoes As ClassMnemonicoComissoes) As Boolean
'Verifica se sAspiranteMnemonico (PARAMETRO de INPUT) eh um mnemonico
'se for, retorna verdadeiro, caso o contrario o resultade eh falso

Dim objMnemonicoCorrente As ClassMnemonicoComissoes
Dim sPropriedade As String

    'se o aspirante a mnemonico for uma string_vazia, sai
    If objMnemonicoComissoes.sMnemonico = STRING_VAZIO Then Exit Function

    'para cada mnemonico na colecao de mnemonicos
    For Each objMnemonicoCorrente In colMnemonicos
            
        'verifica se o mnemonico corrente eh igual ao aspirante
        If objMnemonicoCorrente.sMnemonico = objMnemonicoComissoes.sMnemonico Then

            'Faz com que a funcao retorne true e de quebra retorna o browser e da classe...
            VerificaMnemonico = True
            Set objMnemonicoComissoes = objMnemonicoCorrente
'
'            'coloca na variavel global da tela (sProperty) a string que representa a propriedade
'            'do controle que tera q ir pra tela na hora do evento selecao
'            'CUIDADO AO ALTERAR, EFEITO COLATERAL!!!!
'            sProperty = objMnemonicoCorrente.sPropertyBrowser

            Exit Function

        End If

    Next

    VerificaMnemonico = False

End Function
