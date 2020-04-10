VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRankCli 
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   ScaleHeight     =   4380
   ScaleWidth      =   7950
   Begin VB.CheckBox EmpresaToda 
      Caption         =   "Consolidar Empresa Toda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2520
      TabIndex        =   2
      Top             =   730
      Width           =   2595
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Cliente"
      Height          =   1035
      Left            =   360
      TabIndex        =   23
      Top             =   1920
      Width           =   4755
      Begin VB.ComboBox ComboTipo 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2745
      End
      Begin VB.OptionButton OptionUmTipo 
         Caption         =   "Apenas do Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   5
         Top             =   630
         Width           =   1755
      End
      Begin VB.OptionButton OptionTodosTipos 
         Caption         =   "Todos"
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
         Left            =   150
         TabIndex        =   4
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRankCliOcx.ctx":0000
      Left            =   1050
      List            =   "RelOpRankCliOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRankCliOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRankCliOcx.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRankCliOcx.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRankCliOcx.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exibir"
      Height          =   1125
      Left            =   360
      TabIndex        =   18
      Top             =   3120
      Width           =   4770
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   690
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.OptionButton Apenas 
         Caption         =   "Apenas os"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Todos 
         Caption         =   "Todos os clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   7
         Top             =   285
         Value           =   -1  'True
         Width           =   2145
      End
      Begin MSMask.MaskEdBox NumMaxClientes 
         Height          =   300
         Left            =   1410
         TabIndex        =   9
         Top             =   690
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "maiores"
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
         Left            =   2385
         TabIndex        =   20
         Top             =   750
         Width           =   660
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5880
      Picture         =   "RelOpRankCliOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   2800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1380
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataRef 
      Height          =   330
      Left            =   1830
      TabIndex        =   3
      Top             =   1380
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1020
      TabIndex        =   1
      Top             =   780
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   300
      Width           =   615
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   360
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   840
      Width           =   660
   End
   Begin VB.Label LabelDRef 
      Caption         =   "Data Referência:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "RelOpRankCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private WithEvents objEventoPrevVenda As AdmEvento
Attribute objEventoPrevVenda.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoPrevVenda = New AdmEvento

    'Preenche com os Tipos de Clientes
    lErro = PreencheComboTipo()
    If lErro <> SUCESSO Then gError 90374
    
    'Defina todos os tipos
    Call OptionTodosTipos_Click
    
    Todos.Value = True
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 90374

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Unload Me

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoPrevVenda = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 90213
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    Caption = gobjRelatorio.sCodRel
    
    'Preenche combo com as opções de relatório
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 90214

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 90213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 90214
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'Preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCheckTipo As String
Dim sClienteTipo As String
Dim sCheckMaxCli As String
Dim sMaxCli As String
Dim sCheckEmpToda As String

Dim sClientes As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PreencherRelOp
            
    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 90215
    
    If gobjRelatorio.sCodRel = "Maiores Clientes - KGs" Then
        
        'Verifica se houve escolha por consolidar Empresa_Toda
        If EmpresaToda.Value = 0 Then
            iFilialEmpresa = giFilialEmpresa
            sCheckEmpToda = "0"
            gobjRelatorio.sNomeTsk = "MCli_KGS"
    
        ElseIf EmpresaToda.Value = 1 Then
            iFilialEmpresa = EMPRESA_TODA
            sCheckEmpToda = "1"
            gobjRelatorio.sNomeTsk = "MCliKGET"
        End If
    
    End If
    
    If gobjRelatorio.sCodRel = "Maiores Clientes - R$" Then
        
        'Verifica se houve escolha por consolidar Empresa_Toda
        If EmpresaToda.Value = 0 Then
            iFilialEmpresa = giFilialEmpresa
            sCheckEmpToda = "0"
            gobjRelatorio.sNomeTsk = "MCliReal"
    
        ElseIf EmpresaToda.Value = 1 Then
            iFilialEmpresa = EMPRESA_TODA
            sCheckEmpToda = "1"
            gobjRelatorio.sNomeTsk = "MCliReET"
        End If
    
    End If
    
    'Pode Verifica se existe uma Previsão Mensal de Vendas cadastrada com o código passado
    lErro = PrevVendaMensal_Le_Codigo(Codigo.Text, iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 90203 Then gError 90410
    
    'Se não encontro PrevVenda, erro
    If lErro = 90203 Then gError 90396
    
    'Verifica se a Data foi preenchida
    If Len(DataRef.ClipText) = 0 Then gError 90216
    
    'Se a opção para todos os Tipos de Clientes estiver selecionada
    If OptionTodosTipos.Value = True Then
        sCheckTipo = "Todos"
        sClienteTipo = ""
    Else
        If ComboTipo.Text = "" Then gError 90375
        sCheckTipo = "Um"
        sClienteTipo = ComboTipo.Text
    End If
      
    'Se todos os Clientes estiverem selecionados
    If Todos.Value Then
        sCheckMaxCli = "Todos"
        sMaxCli = ""
    Else
       'verificar se o número de Clientes é válido
        If Len(Trim(NumMaxClientes.Text)) = 0 Then gError 90218
        If CInt(NumMaxClientes.Text) = 0 Then gError 90219
        sCheckMaxCli = "Um"
        sMaxCli = NumMaxClientes.Text
    End If
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 90217
    
    lErro = objRelOpcoes.IncluirParametro("NFILIALEMPRESA", CStr(iFilialEmpresa))
    If lErro <> AD_BOOL_TRUE Then gError 90221
            
    lErro = objRelOpcoes.IncluirParametro("TCODIGO", Codigo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90221
    
    lErro = objRelOpcoes.IncluirParametro("TEMPRESATODA", sCheckEmpToda)
    If lErro <> AD_BOOL_TRUE Then gError 90403
    
    If Trim(DataRef.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DREF", DataRef.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DREF", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 90222
        
    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 90376
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", sClienteTipo)
    If lErro <> AD_BOOL_TRUE Then gError 90377
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOCLIENTE", Codigo_Extrai(sClienteTipo))
    If lErro <> AD_BOOL_TRUE Then gError 90377
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOMAXCLI", sCheckMaxCli)
    If lErro <> AD_BOOL_TRUE Then gError 90378
    
    lErro = objRelOpcoes.IncluirParametro("NMAXCLI", sMaxCli)
    If lErro <> AD_BOOL_TRUE Then gError 90220
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr
        
        Case 90220 To 90222
        
        Case 90376 To 90378
        
        Case 90217, 90403, 90410
               
        Case 90375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_PREENCHIDO", gErr)
            ComboTipo.SetFocus
        
        Case 90216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            DataRef.SetFocus

         Case 90218, 90219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INVALIDO2", gErr, Error$)
            NumMaxClientes.SetFocus

        Case 90215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 90223
    
    'pega Código e exibe
    lErro = objRelOpcoes.ObterParametro("TCODIGO", sParam)
    If lErro <> SUCESSO Then gError 90224

    Codigo.Text = sParam
    
    'Pega  Empresa Toda
    lErro = objRelOpcoes.ObterParametro("TEMPRESATODA", sParam)
    If lErro <> SUCESSO Then gError 90404
    
    If sParam = "1" Then
        EmpresaToda.Value = 1
    Else
        If sParam = "0" Then EmpresaToda.Value = 0
    End If
    
    'pega Data Referencia e exibe
    lErro = objRelOpcoes.ObterParametro("DREF", sParam)
    If lErro <> SUCESSO Then gError 90225

    Call DateParaMasked(DataRef, CDate(sParam))

    'Pega  Tipo cliente e Exibe
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then gError 90379
                   
    If sParam = "Todos" Then
        Call OptionTodosTipos_Click
    Else
        'se é "um tipo só" então exibe o tipo
        lErro = objRelOpcoes.ObterParametro("TTIPOCLIENTE", sParam)
        If lErro <> SUCESSO Then gError 90380
                            
        OptionUmTipo.Value = True
        ComboTipo.Enabled = True
        
        If sParam = "" Then
            ComboTipo.ListIndex = -1
        Else
            ComboTipo.Text = sParam
        End If
    End If
    
    'Pega número máximo de Cliente, se existir e exibe
    lErro = objRelOpcoes.ObterParametro("TTIPOMAXCLI", sParam)
    If lErro <> SUCESSO Then gError 90381
    
    If sParam = "Todos" Then
        Todos.Value = True
    Else
        Apenas.Value = True
        lErro = objRelOpcoes.ObterParametro("NMAXCLI", sParam)
        If lErro <> SUCESSO Then gError 90226
    
        NumMaxClientes.Text = sParam
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 90223 To 90226
        
        Case 90379 To 90381

        Case 90404
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function PreencheComboTipo() As Long

Dim lErro As Long
Dim colCodigoDescricaoCliente As New AdmColCodigoNome
Dim objCodigoDescricaoCliente As New AdmCodigoNome

On Error GoTo Erro_PreencheComboTipo
    
    'Preenche a Colecao com os Tipos de clientes
    lErro = CF("Cod_Nomes_Le", "TiposdeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricaoCliente)
    If lErro <> SUCESSO Then gError 90382
    
   'preenche a ListBox ComboTipo com os objetos da colecao
    For Each objCodigoDescricaoCliente In colCodigoDescricaoCliente
        ComboTipo.AddItem objCodigoDescricaoCliente.iCodigo & SEPARADOR & objCodigoDescricaoCliente.sNome
        ComboTipo.ItemData(ComboTipo.NewIndex) = objCodigoDescricaoCliente.iCodigo
    Next
        
    PreencheComboTipo = SUCESSO

    Exit Function
    
Erro_PreencheComboTipo:

    PreencheComboTipo = gErr

    Select Case gErr

    Case 90382
    
    Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de opção"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 90227

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 90228

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 90229

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 90230

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 90227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 90228 To 90230

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 90231

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRANKCLI")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 90232

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 90231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 90232

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    Call OptionTodosTipos_Click
    Todos.Value = True
    ComboOpcoes.SetFocus
    
End Sub

Private Sub BotaoLimpar_Click()
    
    Limpar_Tela
    ComboOpcoes.Text = ""
    EmpresaToda.Value = 0
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 90233

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 90233

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub LabelCodigo_Click()

Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) > 0 Then
        
        'Preenche com o cliente da tela
        objPrevVendaMensal.sCodigo = Codigo.Text
    End If
    
    'Chama Tela ClienteLista
    Call Chama_Tela("PrevVMensalCodLista", colSelecao, objPrevVendaMensal, objEventoPrevVenda)

End Sub

Private Sub objEventoPrevVenda_evSelecao(obj1 As Object)

Dim objPrevVendaMensal As ClassPrevVendaMensal

    Set objPrevVendaMensal = obj1
    
    Codigo.Text = objPrevVendaMensal.sCodigo
    
    Me.Show

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Codigo_Validate

    'Se o código foi preenchido
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'Verifica se houve escolha por consolidar Empresa_Toda
        If EmpresaToda.Value = 0 Then
            iFilialEmpresa = giFilialEmpresa
        ElseIf EmpresaToda.Value = 1 Then
            iFilialEmpresa = EMPRESA_TODA
        End If
    
        'Verifica se existe uma Previsão Mensal de Vendas cadastrada com o código passado
        lErro = PrevVendaMensal_Le_Codigo(Codigo.Text, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 90203 Then gError 90234
        
        'Se não encontro PrevVenda, erro
        If lErro = 90203 Then gError 90235
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 90234
        
        Case 90235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
    

End Sub

Private Sub DataRef_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataRef)

End Sub

Private Sub DataRef_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataRef_Validate

    If Len(DataRef.ClipText) > 0 Then

        lErro = Data_Critica(DataRef.Text)
        If lErro <> SUCESSO Then gError 90236

    End If

    Exit Sub

Erro_DataRef_Validate:

    Cancel = True

    Select Case gErr

        Case 90236

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataRef, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90237

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 90237
            DataRef.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataRef, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90238

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 90238
            DataRef.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub OptionTodosTipos_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_OptionTodosTipos_Click
    
    'Limpa e desabilita a ComboTipo
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = False
    OptionTodosTipos.Value = True
    
    Exit Sub

Erro_OptionTodosTipos_Click:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub OptionUmTipo_Click()

Dim lErro As Long

On Error GoTo Erro_OptionUmTipo_Click
    
    'Limpa Combo Tipo e Abilita
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = True
    ComboTipo.SetFocus
    
    Exit Sub

Erro_OptionUmTipo_Click:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub Todos_Click()
    
    NumMaxClientes.PromptInclude = False
    NumMaxClientes.Text = ""
    NumMaxClientes.PromptInclude = True
    NumMaxClientes.Enabled = False

End Sub

Private Sub Apenas_Click()
    
    NumMaxClientes.Enabled = True
    NumMaxClientes.SetFocus
        
End Sub

Private Sub NumMaxClientes_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumMaxClientes)

End Sub

Private Sub UpDown2_DownClick()

Dim iIndice As Integer

    If Len(Trim(NumMaxClientes.Text)) = 0 Then Exit Sub
    
    iIndice = CInt(NumMaxClientes.Text)
    
    If iIndice = 0 Then Exit Sub
    
    NumMaxClientes.PromptInclude = False
    NumMaxClientes.Text = CStr(iIndice - 1)
    NumMaxClientes.PromptInclude = True
    NumMaxClientes.SetFocus

End Sub

Private Sub UpDown2_UpClick()

Dim iIndice As Integer

    If Len(Trim(NumMaxClientes.Text)) = 0 Then Exit Sub
    
    iIndice = CInt(NumMaxClientes.Text)
    
    NumMaxClientes.PromptInclude = False
    NumMaxClientes.Text = CStr(iIndice + 1)
    NumMaxClientes.PromptInclude = True
    NumMaxClientes.SetFocus


End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Relação por RankCliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRankCli"
    
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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
    End If
        
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelDRef_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDRef, Source, X, Y)
End Sub

Private Sub LabelDRef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDRef, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub


'Subir para RotinasFATUsu
'***Esta função já está também na RelOpDataRefOcx, RelOpPeriodoOcx, RelOpPrevVendaOcx, RelOpRealPrevOcx
Function PrevVendaMensal_Le_Codigo(sCodigo As String, iFilialEmpresa As Integer) As Long
'Verifica se a previsão de Vendas Mensal de códio e FilialEmpresa passados existem

Dim lErro As Long
Dim iFilial As Integer
Dim lComando As Long

On Error GoTo Erro_PrevVendaMensal_Le_Codigo

    'Abertura de comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 90200
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para a Empresa toda
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? ", iFilial, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    Else
        'Pesquisa no BD se existe a Previsão de Vendas Mensais com o código passado, para uma FilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM PrevVendaMensal WHERE Codigo = ? AND FilialEmpresa = ?", iFilial, sCodigo, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 90201
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90202
    
    'PrevVendas não encontradas
    If lErro = AD_SQL_SEM_DADOS Then gError 90203
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    PrevVendaMensal_Le_Codigo = SUCESSO
    
    Exit Function
    
Erro_PrevVendaMensal_Le_Codigo:
    
    PrevVendaMensal_Le_Codigo = gErr
    
    Select Case gErr
        
        Case 90200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90201, 90202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, sCodigo)
        
        Case 90203 'PrevVendas não cadastrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'Fechamento de comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

