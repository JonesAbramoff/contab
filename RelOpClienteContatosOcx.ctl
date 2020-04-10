VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpClienteContatosOcx 
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   ScaleHeight     =   3810
   ScaleWidth      =   6540
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   1335
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   3525
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   600
         TabIndex        =   2
         Top             =   840
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   405
         Width           =   315
      End
      Begin VB.Label LabelClienteAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   900
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpClienteContatosOcx.ctx":0000
      Left            =   1050
      List            =   "RelOpClienteContatosOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   2730
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
      Left            =   4335
      Picture         =   "RelOpClienteContatosOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   870
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4170
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpClienteContatosOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpClienteContatosOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpClienteContatosOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpClienteContatosOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameTipoCliente 
      Caption         =   "Tipo de Cliente"
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   3525
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   825
         Width           =   1950
      End
      Begin VB.OptionButton TipoApenas 
         Caption         =   "Apenas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   855
         Width           =   1050
      End
      Begin VB.OptionButton TipoTodos 
         Caption         =   "Todos os tipos"
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
         Left            =   195
         TabIndex        =   3
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpClienteContatosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos de browser
Private WithEvents objEventoClienteInicial As AdmEvento
Attribute objEventoClienteInicial.VB_VarHelpID = -1
Private WithEvents objEventoClienteFinal As AdmEvento
Attribute objEventoClienteFinal.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa eventos de browser
    Set objEventoClienteInicial = New AdmEvento
    Set objEventoClienteFinal = New AdmEvento
        
    TipoTodos.Value = True
    Tipo.Enabled = False
    Tipo.ListIndex = -1
    
    'Carrega a combo Tipo
    Call Carrega_ComboTipoCliente(Tipo)
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 131427
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167572)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 131428

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 131429

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 131428
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 131429

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167573)

    End Select

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 131430

    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 131430

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167574)

    End Select

End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 131431

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 131431

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167575)

    End Select

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteDe_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteInicial.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objCliente.sNomeReduzido = ClienteInicial.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInicial, "", sOrdenacao)

    Exit Sub
    
Erro_LabelClienteDe_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167576)
    
    End Select
    
End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelClienteAte_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteFinal.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objCliente.sNomeReduzido = ClienteFinal.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFinal, "", sOrdenacao)

    Exit Sub
    
Erro_LabelClienteAte_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167577)
    
    End Select
    
End Sub

Private Sub TipoTodos_Click()

Dim lErro As Long

On Error GoTo Erro_TipoTodos_Click

    'Desabilita o combotipo
    Tipo.ListIndex = -1
    Tipo.Enabled = False
    
    Exit Sub
    
Erro_TipoTodos_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167578)

    End Select

    Exit Sub

End Sub

Private Sub TipoApenas_Click()

Dim lErro As Long

On Error GoTo Erro_TipoApenas_Click

    'Habilita a ComboTipo
    Tipo.Enabled = True
    
    Exit Sub
    
Erro_TipoApenas_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167579)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131432
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 131432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167580)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
     'Limpa a tela
    Call LimpaRelatorioClientes
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr

        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167581)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 131433

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CLIENTESCONTATOS")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 131434

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioClientes()
        If lErro <> SUCESSO Then gError 131435
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 131433
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 131434, 131435

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167582)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 131436

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131437

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 131438
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 131439
    
    'Limpa a tela
    lErro = LimpaRelatorioClientes()
    If lErro <> SUCESSO Then gError 131440
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 131436
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 131437 To 131440
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167583)

    End Select

    Exit Sub

End Sub

Private Function LimpaRelatorioClientes()
'Limpa a tela RelOpRelacClientes

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioClientes

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 131441
    
    ComboOpcoes.Text = ""
    
    TipoTodos.Value = True
    Tipo.Enabled = False
    Tipo.ListIndex = -1
        
    LimpaRelatorioClientes = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioClientes:

    LimpaRelatorioClientes = gErr
    
    Select Case gErr
    
        Case 131441
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167584)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_De As String
Dim sCliente_Ate As String
Dim sTipo As String

On Error GoTo Erro_PreencherRelOp
   
    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sCliente_De, sCliente_Ate, sTipo)
    If lErro <> SUCESSO Then gError 131442
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 131443
        
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_De)
    If lErro <> AD_BOOL_TRUE Then gError 131444
    
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131650
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 131445
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131651
    
    lErro = objRelOpcoes.IncluirParametro("NTIPO", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError 131446

    lErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", Tipo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131652

    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_De, sCliente_Ate, sTipo)
    If lErro <> SUCESSO Then gError 131447
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 131442 To 131447, 131650 To 131652
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167585)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_De As String, sCliente_Ate As String, sTipo As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
     
    'Verifica se o Cliente inicial foi preenchido
    If ClienteInicial.Text <> "" Then
        sCliente_De = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_De = ""
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If ClienteFinal.Text <> "" Then
        sCliente_Ate = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_Ate = ""
    End If
            
    'Verifica se o Cliente Inicial é menor que o final, se não for --> ERRO
    If sCliente_De <> "" And sCliente_Ate <> "" Then
        
        If CLng(sCliente_De) > CLng(sCliente_Ate) Then gError 131448
    End If
    
    'Se a opção para todos os tipos estiver selecionada
    If TipoTodos.Value = True Then
        sTipo = ""
    
    'Se a opção para apenas um tipo estiver selecionada
    Else
        If Tipo.Text = "" Then gError 131449
        sTipo = CStr(Codigo_Extrai(Tipo.Text))
    
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 131448
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
        
        Case 131449
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            Tipo.SetFocus
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167586)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_De As String, sCliente_Ate As String, sTipo As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
      
    'Verifica se o Cliente Inicial foi preenchido
    If sCliente_De <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(CLng(sCliente_De))
        
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If sCliente_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_Ate))
        
    End If
    
    'Se a opção para apenas um tipo estiver selecionada
    If sTipo <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoCliente = " & Forprint_ConvInt(CInt(sTipo))

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167587)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim iTipo As Integer
Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 131450
    
    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 131451
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 131452
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    lErro = objRelOpcoes.ObterParametro("NTIPO", sParam)
    If lErro <> SUCESSO Then gError 131453
                       
    'Preenche o tipo
    If sParam = "" Then
    
        Tipo.ListIndex = -1
        Tipo.Enabled = False
        TipoTodos.Value = True
    
    Else
    
        TipoApenas.Value = True
        Tipo.Enabled = True
        Call Combo_Seleciona_ItemData(Tipo, StrParaInt(sParam))
        
    End If
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 131450 To 131454
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167588)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoClienteInicial = Nothing
    Set objEventoClienteFinal = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***
Private Sub objEventoClienteInicial_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteInicial_evSelecao

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteInicial.Text = objCliente.sNomeReduzido
    
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoClienteInicial_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167589)
    
    End Select

End Sub

Private Sub objEventoClienteFinal_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoClienteFinal_evSelecao

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteFinal.Text = objCliente.sNomeReduzido
    
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoClienteFinal_evSelecao:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167590)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Tipo_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Tipo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 131510

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoCliente.iCodigo = iCodigo

        'Tenta ler TipoCliente com esse código no BD
        lErro = CF("TipoDeCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 28943 Then gError 131511

        'Não encontrou Tipo Cliente no BD
        If lErro = 28943 Then gError 131512

        'Exibe dados de TipoCliente na tela
        Tipo.Text = CStr(iCodigo) & SEPARADOR & objTipoCliente.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 131513

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case Err

        Case 131510, 131511  'Já tratado na rotina chamada

        Case 131512 'Não encontrou Tipo Cliente no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOCLIENTE")

            If vbMsgRes = vbYes Then

                'Chama a tela de TiposDeClientes
                Call Chama_Tela("TipoCliente", objTipoCliente)

            End If

        Case 131513
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167591)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboTipoCliente(ByVal objComboBox As ComboBox)

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_ComboTipoCliente

    'Lê cada código e descrição da tabela TiposDeCliente
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 131458
    
    'Preenche a ComboBox Tipo com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        objComboBox.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Exit Sub

Erro_Carrega_ComboTipoCliente:

    Select Case gErr
    
        Case 131458
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167592)
    
    End Select
    
    Exit Sub
    
End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Clientes x Contatos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpClienteContatos"
    
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



