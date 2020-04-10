VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl RegrasMsgOcx 
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9615
   Begin VB.ComboBox Loc 
      Height          =   315
      ItemData        =   "RegrasMsgOcx.ctx":0000
      Left            =   5310
      List            =   "RegrasMsgOcx.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   180
      Width           =   1800
   End
   Begin VB.ComboBox Doc 
      Height          =   315
      ItemData        =   "RegrasMsgOcx.ctx":0004
      Left            =   1155
      List            =   "RegrasMsgOcx.ctx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   165
      Width           =   2625
   End
   Begin VB.Frame FrameRegras 
      Caption         =   "Regras para cálculo das mensagens"
      Enabled         =   0   'False
      Height          =   4305
      Left            =   120
      TabIndex        =   13
      Top             =   615
      Width           =   9375
      Begin VB.TextBox Detalhe 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   540
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   3135
         Width           =   8730
      End
      Begin VB.TextBox Mensagem 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3495
         MaxLength       =   250
         TabIndex        =   24
         Top             =   2730
         Width           =   4545
      End
      Begin VB.TextBox Regra5 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3585
         MaxLength       =   250
         TabIndex        =   23
         Top             =   2340
         Width           =   2775
      End
      Begin VB.TextBox Regra4 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3690
         MaxLength       =   250
         TabIndex        =   22
         Top             =   1965
         Width           =   2775
      End
      Begin VB.TextBox Regra3 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3675
         MaxLength       =   250
         TabIndex        =   21
         Top             =   1650
         Width           =   2775
      End
      Begin VB.TextBox Regra2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3720
         MaxLength       =   250
         TabIndex        =   20
         Top             =   1290
         Width           =   2775
      End
      Begin VB.CommandButton BotaoLimparGrid 
         Caption         =   "Limpar Grid"
         Height          =   540
         Left            =   4680
         Picture         =   "RegrasMsgOcx.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   3705
         Width           =   1275
      End
      Begin VB.CommandButton BotaoInserirLinhas 
         Height          =   540
         Left            =   6120
         Picture         =   "RegrasMsgOcx.ctx":053A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Insere uma linha no grid, acima da linha atual (Insert)."
         Top             =   3705
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
         TabIndex        =   2
         Top             =   3870
         Value           =   1  'Checked
         Width           =   3600
      End
      Begin VB.TextBox Regra1 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3705
         MaxLength       =   250
         TabIndex        =   0
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton BotaoSubirRegra 
         Height          =   375
         Left            =   8900
         Picture         =   "RegrasMsgOcx.ctx":267C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton BotaoDescerRegra 
         Height          =   375
         Left            =   8900
         Picture         =   "RegrasMsgOcx.ctx":283E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid GridRegras 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4895
         _Version        =   393216
      End
   End
   Begin VB.Frame FrameFormulas 
      Caption         =   "Monte suas fórmulas"
      Enabled         =   0   'False
      Height          =   1365
      Left            =   120
      TabIndex        =   8
      Top             =   4980
      Width           =   9375
      Begin VB.ComboBox Mnemonicos 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   435
         Width           =   3675
      End
      Begin VB.ComboBox Funcoes 
         Height          =   315
         ItemData        =   "RegrasMsgOcx.ctx":2A00
         Left            =   4080
         List            =   "RegrasMsgOcx.ctx":2A02
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   435
         Width           =   3795
      End
      Begin VB.ComboBox Operadores 
         Height          =   315
         Left            =   8045
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   435
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
         TabIndex        =   9
         Top             =   765
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
         Left            =   8055
         TabIndex        =   12
         Top             =   180
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
         TabIndex        =   11
         Top             =   180
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
         TabIndex        =   10
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      DrawStyle       =   1  'Dash
      Height          =   555
      Left            =   7770
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   45
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "RegrasMsgOcx.ctx":2A04
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "RegrasMsgOcx.ctx":2B82
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "RegrasMsgOcx.ctx":30B4
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Localização:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   4140
      TabIndex        =   19
      Top             =   210
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Origem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   17
      Left            =   450
      TabIndex        =   17
      Top             =   195
      Width           =   660
   End
End
Attribute VB_Name = "RegrasMsgOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const NUM_MAX_REGRAS_COMISSOES = 1000

Const KEYCODE_VERIFICASINTAXE = 116

Const GRID_SUBIR_LINHA = "U"
Const GRID_DESCER_LINHA = "D"

'Propriedade iAlterado da tela
Dim iAlterado As Integer

'obj Gerenciador do Grid da tela
Public objGridRegras As AdmGrid

'Colecao de mnemonicos
Dim colMnemonicos As New Collection

Dim iDocAnt As Integer
Dim iLocAnt As Integer

'Variaveis de controle que representam as colunas do grid
Dim iGrid_Regra1_Col As Integer
Dim iGrid_Regra2_Col As Integer
Dim iGrid_Regra3_Col As Integer
Dim iGrid_Regra4_Col As Integer
Dim iGrid_Regra5_Col As Integer
Dim iGrid_Mensagem_Col As Integer

Public Function Form_Load_Ocx() As Object
    Set Form_Load_Ocx = Me
    Caption = "Regras para cálculo das mensagens"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "Regras"
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

    iDocAnt = -1
    iLocAnt = -1

    'Instancia o objeto que gerencia o Grid
    Set objGridRegras = New AdmGrid
    
    'Instancia a colecao global da tela (mnemonicos)
    Set colMnemonicos = New Collection
            
    'Executa inicializacao do Grid
    lErro = Inicializa_GridRegras(objGridRegras)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Carrega as combos da tela
    lErro = Carrega_Combos_Tela()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209029)

    End Select

    Exit Sub

End Sub

Private Function Traz_Regras_Tela() As Long

Dim lErro As Long
Dim objRegras As ClassRegrasMsg
Dim iLinha As Integer
Dim colRegras As New Collection
Dim iDoc As Integer, iLoc As Integer

On Error GoTo Erro_Traz_Regras_Tela

    Call Grid_Limpa(objGridRegras)

    iDoc = Doc.ItemData(Doc.ListIndex)
    iLoc = Loc.ItemData(Loc.ListIndex)
    
    lErro = CF("RegrasMsg_Le", colRegras, iDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inicializa o numero de linhas
    iLinha = 0
    
    'Para cada regra na colecao
     For Each objRegras In colRegras
     
        If objRegras.iTipoMsg = iLoc Then
        
            iLinha = iLinha + 1
            
            GridRegras.TextMatrix(iLinha, iGrid_Regra1_Col) = objRegras.sRegra1
            GridRegras.TextMatrix(iLinha, iGrid_Regra2_Col) = objRegras.sRegra2
            GridRegras.TextMatrix(iLinha, iGrid_Regra3_Col) = objRegras.sRegra3
            GridRegras.TextMatrix(iLinha, iGrid_Regra4_Col) = objRegras.sRegra4
            GridRegras.TextMatrix(iLinha, iGrid_Regra5_Col) = objRegras.sRegra5
            GridRegras.TextMatrix(iLinha, iGrid_Mensagem_Col) = objRegras.sMensagem
            
        End If
    
    Next
    
    'atualiza o numero de linhas existentes
    objGridRegras.iLinhasExistentes = iLinha
    
    Traz_Regras_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Regras_Tela:

    Traz_Regras_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209030)
        
    End Select
        
    Exit Function

End Function

Private Function Carrega_Combos_Tela() As Long
'Responsavel pela carga das comboboxes existentes na tela..

Dim lErro As Long

On Error GoTo Erro_Carrega_Combos_Tela
   
    'Carrega a combo Funcoes
    lErro = Carrega_Funcoes()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Carrega a combo Operadores
    lErro = Carrega_Operacoes()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Doc.AddItem REGRAMSG_TIPODOC_NF_TEXTO
    Doc.ItemData(Doc.NewIndex) = REGRAMSG_TIPODOC_NF

    Doc.AddItem REGRAMSG_TIPODOC_ITEMNF_TEXTO
    Doc.ItemData(Doc.NewIndex) = REGRAMSG_TIPODOC_ITEMNF
        
    Carrega_Combos_Tela = SUCESSO

    Exit Function

Erro_Carrega_Combos_Tela:
    
    Carrega_Combos_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209031)
            
    End Select
    
    Exit Function
    
End Function

Private Function Inicializa_GridRegras(ByVal objGridInt As AdmGrid) As Long
'Inicializa o grid da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridRegras

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Mensagem")
    objGridInt.colColuna.Add ("Condição 1")
    objGridInt.colColuna.Add ("Condição 2")
    objGridInt.colColuna.Add ("Condição 3")
    objGridInt.colColuna.Add ("Condição 4")
    objGridInt.colColuna.Add ("Condição 5")
    
    'campos de edição do grid
    objGridInt.colCampo.Add (Mensagem.Name)
    objGridInt.colCampo.Add (Regra1.Name)
    objGridInt.colCampo.Add (Regra2.Name)
    objGridInt.colCampo.Add (Regra3.Name)
    objGridInt.colCampo.Add (Regra4.Name)
    objGridInt.colCampo.Add (Regra5.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_Mensagem_Col = 1
    iGrid_Regra1_Col = 2
    iGrid_Regra2_Col = 3
    iGrid_Regra3_Col = 4
    iGrid_Regra4_Col = 5
    iGrid_Regra5_Col = 6
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRegras

    'Numero Maximo de Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REGRAS_COMISSOES + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridRegras.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'permite insercao de linhas no meio do grid
    objGridInt.iProibidoIncluirNoMeioGrid = GRID_PERMITIDO_INCLUIR_NO_MEIO
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridRegras = SUCESSO

    Exit Function

Erro_Inicializa_GridRegras:

    Inicializa_GridRegras = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209032)
            
    End Select

    Exit Function

End Function

Private Function Carrega_Mnemonicos(ByVal iDoc As Integer) As Long
'Preenche a combo de mnemonicos com o conteudo do bd

Dim lErro As Long
Dim objMnemonico As ClassMnemonicoRegrasMsg

On Error GoTo Erro_Carrega_Mnemonicos

    Set colMnemonicos = New Collection

    'leitura dos mnemonicos no BD
    lErro = CF("MnemonicoRegrasMsg_Le", colMnemonicos, iDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'limpa a combo
    Mnemonicos.Clear
    Detalhe.Text = ""
    
    'para cada mnemonico na colecao carregada anteriormente
    For Each objMnemonico In colMnemonicos
    
        'adiciona o dito cujo na combo
        Mnemonicos.AddItem objMnemonico.sMnemonico
        
    Next
        
    Carrega_Mnemonicos = SUCESSO

    Exit Function

Erro_Carrega_Mnemonicos:

    Carrega_Mnemonicos = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209033)
            
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
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
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209034)
            
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
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
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209035)
            
    End Select
    
    Exit Function

End Function

Private Function Regras_Critica() As Long
'Funcao que faz a critica da tela Regras

Dim iLinha As Integer
Dim lErro As Long
Dim iIndice As Integer
Dim sRegra As String
Dim iInicio As Integer
Dim iTamanho As Integer
Dim iTipo As Integer

On Error GoTo Erro_Regras_Critica

    If iDocAnt = -1 Then gError 209036
    If iLocAnt = -1 Then gError 209037

    'Para cada linha do grid
    For iLinha = 1 To objGridRegras.iLinhasExistentes
    
        iTipo = TIPO_BOOLEANO
    
        For iIndice = 1 To 6
        
            Select Case iIndice
                Case 1
                    sRegra = GridRegras.TextMatrix(iLinha, iGrid_Regra1_Col)
                Case 2
                    sRegra = GridRegras.TextMatrix(iLinha, iGrid_Regra2_Col)
                Case 3
                    sRegra = GridRegras.TextMatrix(iLinha, iGrid_Regra3_Col)
                Case 4
                    sRegra = GridRegras.TextMatrix(iLinha, iGrid_Regra4_Col)
                Case 5
                    sRegra = GridRegras.TextMatrix(iLinha, iGrid_Regra5_Col)
                Case 6
                    sRegra = GridRegras.TextMatrix(iLinha, iGrid_Mensagem_Col)
                    iTipo = TIPO_TEXTO
            End Select
    
            'Se o campo regra estiver preenchido
            If Len(Trim(sRegra)) > 0 Then
            
                'Validar as regras
                lErro = CF("Valida_Formula_WFW", sRegra, iTipo, iInicio, iTamanho, colMnemonicos)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End If
        
        Next
        
    Next
    
    Regras_Critica = SUCESSO
    
    Exit Function
    
Erro_Regras_Critica:

    Regras_Critica = gErr

    Select Case gErr
    
        Case 209036
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", gErr)
        
        Case 209037
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCALIZACAO_NAO_PREENCHIDA", gErr)
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209038)
            
    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a funcao que ira efetuar a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0

    'fecha a tela apos a gravacao (padrao nas telas de configuracao)
    Call Limpa_Regras

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209039)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colRegras As New Collection
Dim iDoc As Integer, iLoc As Integer

On Error GoTo Erro_Gravar_Registro

     'Exibe uma ampulheta como ponteiro do mouse
     'para que o usuario tenha o feedback da gravacao
     GL_objMDIForm.MousePointer = vbHourglass
     
     'Critica os Dados que serao gravados
     lErro = Regras_Critica()
     If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
          
     'Guarda os dados presentes na tela na colecao de objetos de ClassRegras..
     lErro = Move_Tela_Memoria(colRegras)
     If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

     'grava as regras de comissoes no BD
     lErro = CF("RegrasMsg_Grava", iDocAnt, iLocAnt, colRegras)
     If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209040)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colRegras As Collection) As Long

Dim lErro As Long
Dim objRegras As ClassRegrasMsg
Dim iLinha As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'para cada linha do grid
    For iLinha = 1 To objGridRegras.iLinhasExistentes
    
        'Instancia uma nova area de memoria a ser apontada pelo obj
        Set objRegras = New ClassRegrasMsg
        
        objRegras.iSeq = iLinha
        objRegras.iTipoDoc = iDocAnt
        objRegras.iTipoMsg = iLocAnt
    
        objRegras.sRegra1 = GridRegras.TextMatrix(iLinha, iGrid_Regra1_Col)
        objRegras.sRegra2 = GridRegras.TextMatrix(iLinha, iGrid_Regra2_Col)
        objRegras.sRegra3 = GridRegras.TextMatrix(iLinha, iGrid_Regra3_Col)
        objRegras.sRegra4 = GridRegras.TextMatrix(iLinha, iGrid_Regra4_Col)
        objRegras.sRegra5 = GridRegras.TextMatrix(iLinha, iGrid_Regra5_Col)
        objRegras.sMensagem = GridRegras.TextMatrix(iLinha, iGrid_Mensagem_Col)
       
        'adiciona o obj na colecao
        colRegras.Add objRegras
    
    Next
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209041)
            
    End Select
    
    Exit Function

End Function

Private Sub Limpa_Regras()
'Limpa o grid e marca a check verifica sintaxe...

    Doc.ListIndex = -1
    Loc.ListIndex = -1
    Funcoes.ListIndex = -1
    Operadores.ListIndex = -1
    
    'chama a grid limpa (limpa o grid de regras)
    Call Grid_Limpa(objGridRegras)
    
    'Marca check de VerificaSintaxe
    VerificaSintaxe.Value = vbChecked
    
    Call Limpa_Tela(Me)
    
    iDocAnt = -1
    iLocAnt = -1
    
    iAlterado = 0

End Sub

Private Sub BotaoLimparGrid_Click()
    Call Grid_Limpa(objGridRegras)
    Detalhe.Text = ""
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'chama a teste_salva
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a tela
    Call Limpa_Regras

    'Fecha Comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209042)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub GridRegras_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRegras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegras, iAlterado)
    End If

End Sub

Public Sub GridRegras_GotFocus()
    Call Grid_Recebe_Foco(objGridRegras)
End Sub

Public Sub GridRegras_EnterCell()
    Call Grid_Entrada_Celula(objGridRegras, iAlterado)
End Sub

Public Sub GridRegras_LeaveCell()
    Call Saida_Celula(objGridRegras)
End Sub

Public Sub GridRegras_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridRegras)
End Sub

Public Sub GridRegras_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRegras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegras, iAlterado)
    End If

End Sub

Public Sub GridRegras_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridRegras)

End Sub

Public Sub GridRegras_RowColChange()
    
    Call Grid_RowColChange(objGridRegras)
    If GridRegras.Row > 0 Then
        Detalhe.Text = GridRegras.TextMatrix(GridRegras.Row, GridRegras.Col)
    Else
        Detalhe.Text = ""
    End If
    
End Sub

Public Sub GridRegras_Scroll()
    
    Call Grid_Scroll(objGridRegras)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'inicializa a saida
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
                
        lErro = Saida_Celula_Regra(objGridInt, Me.ActiveControl)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'finaliza a saida
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 209043
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209043
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209044)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Regra(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica do campo ValorBase que está deixando de ser o campo corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim iTipo As Integer

On Error GoTo Erro_Saida_Celula_Regra

    'instancia objcontrole como o controle de regra
    Set objGridInt.objControle = objControle

    'Se o campo está preenchido
    If Len(Trim(objControle.Text)) > 0 And VerificaSintaxe.Value = vbChecked Then
    
        If objControle.Name = Mensagem.Name Then
            iTipo = TIPO_TEXTO
        Else
            iTipo = TIPO_BOOLEANO
        End If
    
         lErro = CF("Valida_Formula_WFW", objControle.Text, iTipo, iInicio, iTamanho, colMnemonicos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    End If
        
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If Len(Trim(objControle.Text)) > 0 And objControle.Name = Mensagem.Name Then
        If GridRegras.Row - GridRegras.FixedRows = objGridRegras.iLinhasExistentes Then
            objGridRegras.iLinhasExistentes = objGridRegras.iLinhasExistentes + 1
        End If
    End If
    
    Saida_Celula_Regra = SUCESSO

    Exit Function

Erro_Saida_Celula_Regra:

    Saida_Celula_Regra = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209045)

    End Select

    Exit Function

End Function

Private Sub Regra1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Regra1_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegras)
End Sub

Private Sub Regra1_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)
End Sub

Private Sub Regra1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Regra1
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Regra2_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Regra2_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegras)
End Sub

Private Sub Regra2_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)
End Sub

Private Sub Regra2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Regra2
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Regra3_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Regra3_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegras)
End Sub

Private Sub Regra3_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)
End Sub

Private Sub Regra3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Regra3
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Regra4_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Regra4_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegras)
End Sub

Private Sub Regra4_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)
End Sub

Private Sub Regra4_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Regra4
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Regra5_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Regra5_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegras)
End Sub

Private Sub Regra5_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)
End Sub

Private Sub Regra5_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Regra5
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Mensagem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRegras)
End Sub

Private Sub Mensagem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)
End Sub

Private Sub Mensagem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Mensagem
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'
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
    If lErro <> SUCESSO And lErro <> 36088 Then gError ERRO_SEM_MENSAGEM
    
    'se nao achou => erro (possivel exclusao da formula funcao durante a utilizacao da tela)
    If lErro <> SUCESSO Then gError 209046
    
    'coloca a descricao no campo adequado
    Descricao.Text = objFormulaFuncao.sFuncaoDesc
    
    'copia o conteudo da combo para o grid se for o caso
    Call Posiciona_Combo
    
    Exit Sub

Erro_Funcoes_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 209046
            Call Rotina_Erro(vbOKOnly, "ERRO_FUNCAO_NAO_CADASTRADA", gErr, objFormulaFuncao.sFuncaoCombo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209047)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Mnemonicos_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Mnemonicos_Click()

Dim lErro As Long
Dim objMnemonico As ClassMnemonicoRegrasMsg

On Error GoTo Erro_Mnemonicos_Click

    'se a combo nao estiver preenchida, sai
    If Len(Trim(Mnemonicos.Text)) = 0 Then Exit Sub
    
    'faz com q o obj aponte para o item da colecao referenciado pela combo..
    For Each objMnemonico In colMnemonicos
        If objMnemonico.sMnemonico = Mnemonicos.Text Then
            Exit For
        End If
    Next
    
    'coloca a descricao no devido lugar (campo descricao da tela..)
    Descricao.Text = objMnemonico.sMnemonicoDesc
    
    'copia o conteudo da combo para o grid se for o caso
    Call Posiciona_Combo
    
    Exit Sub

Erro_Mnemonicos_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209048)
            
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
    If lErro <> SUCESSO And lErro <> 36098 Then gError ERRO_SEM_MENSAGEM
    
    'se nao achou => erro (possivel exclusao da formula operador durante a utilizacao da tela)
    If lErro <> SUCESSO Then gError 209049
    
    'coloca a descricao no campo adequado
    Descricao.Text = objFormulaOperador.sOperadorDesc
    
    'copia o conteudo da combo para o grid se for o caso
    Call Posiciona_Combo
    
    Exit Sub

Erro_Operadores_Click:

    Select Case gErr
       
        Case 209049
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_CADASTRADO", gErr, objFormulaOperador.sOperadorCombo)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209050)
            
    End Select
    
    Exit Sub

End Sub

Function Posiciona_Combo() As Long
'Coloca o texto no grid, caso alguma coluna e linha esteja selecionada
'a coluna deve ser uma coluna passivel de receber um operador/funcao/mnemonico

    'se existe linha e coluna selecionada
    If GridRegras.Row > 0 And GridRegras.Row <= objGridRegras.iLinhasExistentes + 1 And GridRegras.Col > 0 Then

        'seleciona a coluna
        Select Case GridRegras.Col

            Case iGrid_Regra1_Col
                Call Posiciona_Texto_Tela(Regra1, Me.ActiveControl.Text)

            Case iGrid_Regra2_Col
                Call Posiciona_Texto_Tela(Regra2, Me.ActiveControl.Text)

            Case iGrid_Regra3_Col
                Call Posiciona_Texto_Tela(Regra3, Me.ActiveControl.Text)

            Case iGrid_Regra4_Col
                Call Posiciona_Texto_Tela(Regra4, Me.ActiveControl.Text)

            Case iGrid_Regra5_Col
                Call Posiciona_Texto_Tela(Regra5, Me.ActiveControl.Text)

            'se for a coluna de valor base
            Case iGrid_Mensagem_Col
                Call Posiciona_Texto_Tela(Mensagem, Me.ActiveControl.Text)

        End Select

    End If

End Function

Private Sub Posiciona_Texto_Tela(objControl As Control, sTexto As String)

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
    If right(Trim(sTextoEsq), 2) <> OPERADOR_MAIORIGUAL And right(Trim(sTextoEsq), 2) <> OPERADOR_MENORIGUAL And right(Trim(sTextoEsq), 2) <> OPERADOR_DIFERENTE And right(Trim(sTextoEsq), 1) <> OPERADOR_IGUAL And right(Trim(sTextoEsq), 1) <> OPERADOR_MAIOR And right(Trim(sTextoEsq), 1) <> OPERADOR_MENOR Then sTextoEsq = sTextoEsq & " " & OPERADOR_IGUAL & " "
    
    'Insere no controle o texto passado como parâmetro na posição correta
    objControl.Text = sTextoEsq & " " & sTexto & " " & sTextoDir
    
    'Atualiza a posição do cursor, posicionando-o logo após ao texto que foi inserido no controle
    objControl.SelStart = Len(sTextoEsq) + Len(sTexto) + 1
    
    'Se o controle em questão não é controle ativo => descobre a posição onde o texto será inserido no campo do grid
    If Not (Me.ActiveControl Is objControl) Then
        
        'Se a posição de inserção do texto é maior do que o texto no campo que será atualizado
        If iPos >= Len(GridRegras.TextMatrix(GridRegras.Row, GridRegras.Col)) Then
            
            'Indica que não há expressão a ser exibida após o texto que será inserido no grid
            iTamanho = 0
        
        'Senão
        Else
            
            'Guarda o tamanho da expressão que virá após o texto a ser inserido no grid
            iTamanho = Len(GridRegras.TextMatrix(GridRegras.Row, GridRegras.Col)) - iPos
        
        End If
        
        'Insere no grid o texto passado como parâmetro na posição correta
        GridRegras.TextMatrix(GridRegras.Row, GridRegras.Col) = Mid(GridRegras.TextMatrix(GridRegras.Row, GridRegras.Col), 1, iPos) & sTexto & Mid(GridRegras.TextMatrix(GridRegras.Row, GridRegras.Col), iPos + 1, iTamanho)
        
        iAlterado = REGISTRO_ALTERADO
        
    End If
    
    Exit Sub
    
Erro_Posiciona_Texto_Tela:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209051)
    
    End Select
    
End Sub

Private Sub BotaoSubirRegra_Click()

    'se está na primeira linha do grid-> sai
    If GridRegras.Row <= GridRegras.FixedRows Then Exit Sub
    
    'se a linha que se quer mover para cima está dentro dos limites das existentes
    If GridRegras.Row <= objGridRegras.iLinhasExistentes Then
    
        'Inverte a posição da linha atual com a linha de cima
        Call Troca_Linha(GridRegras.Row, GridRegras.Row - 1)
    
    End If


End Sub

Private Sub BotaoDescerRegra_Click()

    'se está na última linha do grid-> sai
    If GridRegras.Row >= objGridRegras.iLinhasExistentes Then Exit Sub
    
    'se a linha que se quer mover para baixo está dentro dos limites das existentes
    If GridRegras.Row >= GridRegras.FixedRows Then
    
        'Inverte a posição da linha atual com a linha de cima
        Call Troca_Linha(GridRegras.Row, GridRegras.Row + 1)
    
    End If

End Sub

Private Sub Troca_Linha(iLinha1 As Integer, iLinha2 As Integer)
'Troca o conteudo de iLinha1 com o conteudo de iLinha2
'os 2 parametros sao de INPUT

Dim iIndice As Integer
Dim asValor(1 To 6) As String

    'Copia o conteudo da linha1 para a memoria
    For iIndice = 1 To GridRegras.Cols - 1
        asValor(iIndice) = GridRegras.TextMatrix(iLinha1, iIndice)
        GridRegras.TextMatrix(iLinha1, iIndice) = GridRegras.TextMatrix(iLinha2, iIndice)
        GridRegras.TextMatrix(iLinha2, iIndice) = asValor(iIndice)
    Next
    
    'coloca a linha2 como corrente para que de a impressao de q esta carregando a linha
    GridRegras.Row = iLinha2
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub BotaoInserirLinhas_Click()
'insere uma linha no grid

    iAlterado = REGISTRO_ALTERADO

    'coloca o foco no grid
    GridRegras.SetFocus
    
    'emula a tecla esc (o foco tende a ir para o controle)
    Call SendKeys("{ESC}", True)
    
    'aciona a tecla insert que eh a responsavel por inserir a linha no meio do grid
    Call SendKeys("{INSERT}")
    
    Call SendKeys("{ENTER}", True)
    
End Sub

Private Sub VerificaSintaxe_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Trata_Doc_Loc() As Long

Dim lErro As Long
Dim iDoc As Integer
Dim iLoc As Integer

On Error GoTo Erro_Trata_Doc_Loc
    
    If Doc.ListIndex <> -1 Then
        iDoc = Doc.ItemData(Doc.ListIndex)
        
        If iDocAnt <> iDoc Then
        
            'Se trocou o a origem testa para ver se não quer salvar antes de trocar o iDocAnt
            lErro = Teste_Salva(Me, iAlterado)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            iAlterado = 0
        
            lErro = Carrega_Mnemonicos(iDoc)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Loc.Clear
            Select Case iDoc
            
                Case REGRAMSG_TIPODOC_NF
                
                    Loc.AddItem "Corpo da Nota"
                    Loc.ItemData(Loc.NewIndex) = REGRAMSG_TIPOMSG_CORPO
                                
                    Loc.AddItem "Dados Adicionais"
                    Loc.ItemData(Loc.NewIndex) = REGRAMSG_TIPOMSG_NORMAL
                
                Case REGRAMSG_TIPODOC_ITEMNF
                
                    Loc.AddItem "Itens"
                    Loc.ItemData(Loc.NewIndex) = REGRAMSG_TIPOMSG_NORMAL
                
            End Select
            
            iDocAnt = iDoc
            
        End If
    Else
    
        'Se trocou o a origem testa para ver se não quer salvar antes de trocar o iDocAnt
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        iDoc = -1
        Call Grid_Limpa(objGridRegras)
        Detalhe.Text = ""
        Mnemonicos.Clear
        Set colMnemonicos = New Collection
        iAlterado = 0
    End If
    
    If Loc.ListIndex <> -1 Then
        
        FrameRegras.Enabled = True
        FrameFormulas.Enabled = True
        
        iLoc = Loc.ItemData(Loc.ListIndex)
        
        If iLocAnt <> iLoc Then
        
            'Se trocou o a localização testa para ver se não quer salvar antes de trocar o iLocAnt
            lErro = Teste_Salva(Me, iAlterado)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            lErro = Traz_Regras_Tela()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            iLocAnt = iLoc
            iAlterado = 0
        
        End If
    
    Else
    
        'Se trocou o a localização testa para ver se não quer salvar antes de trocar o iLocAnt
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        FrameRegras.Enabled = False
        FrameFormulas.Enabled = False
        
        iLoc = -1
        Call Grid_Limpa(objGridRegras)
        Detalhe.Text = ""
        iAlterado = 0
       
    End If
    
    Trata_Doc_Loc = SUCESSO

    Exit Function

Erro_Trata_Doc_Loc:

    Trata_Doc_Loc = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209052)

    End Select

    Exit Function

End Function

Private Sub Doc_Change()
    Call Trata_Doc_Loc
End Sub

Private Sub loc_Change()
    Call Trata_Doc_Loc
End Sub

Private Sub Doc_Click()
    Call Trata_Doc_Loc
End Sub

Private Sub loc_Click()
    Call Trata_Doc_Loc
End Sub
