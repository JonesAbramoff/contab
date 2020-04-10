VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ClienteContatosOcx 
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   KeyPreview      =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   9300
   Begin VB.Frame FrameCliente 
      Caption         =   "Cliente"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   6825
      Begin VB.CommandButton BotaoTrazer 
         Height          =   375
         Left            =   -5520
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Cliente 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         ToolTipText     =   "Digite código, nome reduzido, cgc do cliente ou pressione F3 para consulta."
         Top             =   195
         Width           =   2730
      End
      Begin VB.ComboBox FilialCliente 
         Height          =   315
         Left            =   4485
         TabIndex        =   1
         ToolTipText     =   "Digite o nome ou o código da filial do cliente com quem foi feito o relacionamento."
         Top             =   195
         Width           =   2280
      End
      Begin VB.Label LabelFilialCliente 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Index           =   2
         Left            =   3885
         TabIndex        =   18
         Top             =   255
         Width           =   465
      End
      Begin VB.Label LabelCliente 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   255
         Width           =   660
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Contatos"
      Height          =   4980
      Left            =   120
      TabIndex        =   15
      Top             =   645
      Width           =   9090
      Begin VB.TextBox Codigo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   645
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1935
         Width           =   675
      End
      Begin VB.TextBox OutrosMeioComunic 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2025
         Width           =   2385
      End
      Begin MSMask.MaskEdBox DataNasc 
         Height          =   255
         Left            =   2370
         TabIndex        =   20
         Top             =   1920
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CheckBox Padrao 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Cargo 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   1125
      End
      Begin MSMask.MaskEdBox Telefone 
         Height          =   240
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fax 
         Height          =   240
         Left            =   1800
         TabIndex        =   8
         Top             =   1440
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin VB.TextBox Setor 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   5
         Top             =   960
         Width           =   1125
      End
      Begin VB.TextBox Contato 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   960
         MaxLength       =   50
         TabIndex        =   4
         Top             =   960
         Width           =   2025
      End
      Begin VB.TextBox Email 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1440
         Width           =   2250
      End
      Begin MSFlexGridLib.MSFlexGrid GridContatos 
         Height          =   4560
         Left            =   60
         TabIndex        =   2
         Top             =   225
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   8043
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7065
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ClienteContatos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ClienteContatos.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ClienteContatos.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "ClienteContatos.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "ClienteContatosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Definicoes do Grid de Contatos
Dim objGridContatos As New AdmGrid

Dim iGrid_Codigo_Col As Integer
Dim iGrid_Padrao_Col As Integer
Dim iGrid_Contato_Col As Integer
Dim iGrid_Setor_Col As Integer
Dim iGrid_Telefone_Col As Integer
Dim iGrid_Fax_Col As Integer
Dim iGrid_Email_Col As Integer
Dim iGrid_Cargo_Col As Integer
Dim iGrid_DataNasc_Col As Integer
Dim iGrid_OutMeioComuni_Col As Integer

Public iAlterado As Integer
Public iClienteAlterado As Integer
Dim gbAtualizarGrid As Boolean

Dim giProxCod As Integer

'Evento de browser
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim iFilialClienteAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Function Trata_Parametros(Optional ByVal objClienteContatos As ClassClienteContatos) As Long
'Verifica validade dos parametros passados pela tela chamadora

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se recebeu um objeto com contatos de um cliente
    If Not (objClienteContatos Is Nothing) Then
    
        'Se o código do cliente está preenchido
        If objClienteContatos.lCliente > 0 Then
        
            'Exibe o cliente na tela e faz a validação do mesmo
            Cliente.Text = objClienteContatos.lCliente
            lErro = Valida_Cliente()
            If lErro <> SUCESSO Then gError 102612
            
            'Exibe o cliente na tela e faz a validação do mesmo
            FilialCliente.Text = objClienteContatos.iFilialCliente
            lErro = Valida_FilialCliente()
            If lErro <> SUCESSO Then gError 102613
            
            'Lê e traz os dados do relacionamento para a tela
            lErro = Traz_ClienteContatos_Tela(objClienteContatos)
            If lErro <> SUCESSO Then gError 102569
        
            If Len(Trim(objClienteContatos.sContato)) > 0 Then
               objGridContatos.iLinhasExistentes = objGridContatos.iLinhasExistentes + 1
               GridContatos.TextMatrix(objGridContatos.iLinhasExistentes, iGrid_Contato_Col) = objClienteContatos.sContato
            End If
        
        End If
    
    End If
    
    'Inidica que o click no campo filialcliente deve atualizar o grid
    gbAtualizarGrid = True
    
    iAlterado = 0
    iClienteAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 102569, 102612, 102613
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154201)

    End Select

    iAlterado = 0
    iClienteAlterado = 0
    iFilialClienteAlterado = 0
    
    Exit Function

End Function

Sub Limpa_ClienteContatos()

Dim lErro As Long

    'Limpa Tela
    Call Limpa_Tela(Me)

    'Limpa o grid
    Call Grid_Limpa(objGridContatos)
    
    giProxCod = 1
    
    'Limpa a combo filial
    FilialCliente.Text = ""

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_Load()
'Inicializa a tela

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa o evento de browser
    Set objEventoCliente = New AdmEvento
    
    'Executa inicializacao do GridContatos
    lErro = Inicializa_Grid_Contatos(objGridContatos)
    If lErro <> SUCESSO Then gError 121197

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 121197
            'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154202)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objGridContatos = Nothing
    Set objEventoCliente = Nothing
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Function Traz_ClienteContatos_Tela(objClienteContatos As ClassClienteContatos) As Long
'Carrega o Grid de contatos com os dados trazidos da tabela ClienteContatos

Dim lErro As Long
Dim iLinha As Integer
Dim colClienteContatos As New Collection
Dim objClienteContatos1 As ClassClienteContatos
Dim iCod As Integer

On Error GoTo Erro_Traz_ClienteContatos_Tela

    'Limpa o Grid de Contato
    Call Grid_Limpa(objGridContatos)

    iLinha = 0

    'Le os dados da tabela ContatoGeral e Preenche a colContatoGeral
    lErro = CF("ClienteContatos_Le_Todos", colClienteContatos, objClienteContatos)
    If lErro <> SUCESSO And lErro <> 102574 Then gError 121178

    'Preenche o grid com os objetos da coleção de contato
    'Para cada Contato encontrado
    For Each objClienteContatos1 In colClienteContatos

       iLinha = iLinha + 1

        GridContatos.TextMatrix(iLinha, iGrid_Contato_Col) = objClienteContatos1.sContato
        GridContatos.TextMatrix(iLinha, iGrid_Setor_Col) = objClienteContatos1.sSetor
        GridContatos.TextMatrix(iLinha, iGrid_Email_Col) = objClienteContatos1.sEmail
        GridContatos.TextMatrix(iLinha, iGrid_Telefone_Col) = objClienteContatos1.sTelefone
        GridContatos.TextMatrix(iLinha, iGrid_Fax_Col) = objClienteContatos1.sFax
        GridContatos.TextMatrix(iLinha, iGrid_Cargo_Col) = objClienteContatos1.sCargo
        GridContatos.TextMatrix(iLinha, iGrid_Padrao_Col) = objClienteContatos1.iPadrao
        If objClienteContatos1.dtDataNasc <> DATA_NULA Then GridContatos.TextMatrix(iLinha, iGrid_DataNasc_Col) = Format(objClienteContatos1.dtDataNasc, "dd/mm/yyyy")
        GridContatos.TextMatrix(iLinha, iGrid_Codigo_Col) = CStr(objClienteContatos1.iCodigo)
        GridContatos.TextMatrix(iLinha, iGrid_OutMeioComuni_Col) = objClienteContatos1.sOutrosMeioComunic

        iCod = objClienteContatos1.iCodigo

    Next

    'Guarda o número de linhas existentes
    objGridContatos.iLinhasExistentes = iLinha
   
    giProxCod = iCod + 1
    GridContatos.TextMatrix(iLinha + 1, iGrid_Codigo_Col) = CStr(giProxCod)
    giProxCod = giProxCod + 1
    
    'Atualiza a checkbox do grid
    Call Grid_Refresh_Checkbox(objGridContatos)

    Traz_ClienteContatos_Tela = SUCESSO

    Exit Function

Erro_Traz_ClienteContatos_Tela:

    Traz_ClienteContatos_Tela = gErr

    Select Case gErr

        Case 121178

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154203)

    End Select

    Exit Function

End Function

Private Function Valida_GridContatos() As Long
'Verifica validade do GridContatos

Dim iIndice As Integer

On Error GoTo Erro_Valida_GridContatos

    'Se o cliente não foi preenchido erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 102605
    
    'Se a filial do cliente não foi preenchida => erro
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 102606

    'Se nenhuma linha do grid foi preenchida => erro
    If objGridContatos.iLinhasExistentes <= 0 Then gError 121200

    'Para cada linha existente no grid
    For iIndice = 1 To objGridContatos.iLinhasExistentes

       'Se o contato não foi preenchido => erro
       If Len(Trim(GridContatos.TextMatrix(iIndice, iGrid_Contato_Col))) = 0 Then gError 121201

    Next

    Valida_GridContatos = SUCESSO

    Exit Function

Erro_Valida_GridContatos:

    Valida_GridContatos = gErr

    Select Case gErr

        Case 102605
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102606
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 121200
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)

        Case 121201
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_GRID_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154204)

    End Select

     Exit Function

End Function

Private Function Inicializa_Grid_Contatos(objGridInt As AdmGrid) As Long
'Inicializa o grid de Contatos

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Contatos

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Padrão")
    objGridInt.colColuna.Add ("Contato")
    objGridInt.colColuna.Add ("Setor")
    objGridInt.colColuna.Add ("Cargo")
    objGridInt.colColuna.Add ("Telefone")
    objGridInt.colColuna.Add ("Fax")
    objGridInt.colColuna.Add ("E-Mail")
    objGridInt.colColuna.Add ("Nascimento")
    objGridInt.colColuna.Add ("Outros Meios de Comunicação")

    'campos de edição do grid
    objGridInt.colCampo.Add (Codigo.Name)
    objGridInt.colCampo.Add (Padrao.Name)
    objGridInt.colCampo.Add (Contato.Name)
    objGridInt.colCampo.Add (Setor.Name)
    objGridInt.colCampo.Add (Cargo.Name)
    objGridInt.colCampo.Add (Telefone.Name)
    objGridInt.colCampo.Add (Fax.Name)
    objGridInt.colCampo.Add (Email.Name)
    objGridInt.colCampo.Add (DataNasc.Name)
    objGridInt.colCampo.Add (OutrosMeioComunic.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Codigo_Col = 1
    iGrid_Padrao_Col = 2
    iGrid_Contato_Col = 3
    iGrid_Setor_Col = 4
    iGrid_Cargo_Col = 5
    iGrid_Telefone_Col = 6
    iGrid_Fax_Col = 7
    iGrid_Email_Col = 8
    iGrid_DataNasc_Col = 9
    iGrid_OutMeioComuni_Col = 10
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridContatos

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_CONTATOS

    'Largura da primeira coluna
    GridContatos.ColWidth(0) = 600

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'habilita a rotina grid enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 15

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contatos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contatos:

    Inicializa_Grid_Contatos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154205)

    End Select

    Exit Function

End Function

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_BotaoTrazer_Click

    'Obtém código do cliente e da filial que estão na tela
    lErro = Obtem_Cod_Cliente_Filial(objClienteContatos)
    If lErro <> SUCESSO Then gError 102610
    
    'Traz para a tela os contatos do cliente em questão
    lErro = Traz_ClienteContatos_Tela(objClienteContatos)
    If lErro <> SUCESSO Then gError 102601

    Exit Sub
    
Erro_BotaoTrazer_Click:

    Select Case gErr

        Case 102610
        
        Case 102601
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154206)

    End Select

End Sub

Private Sub Cliente_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = REGISTRO_ALTERADO Then

        'Executa a validação do cliente
        lErro = Valida_Cliente
        If lErro <> SUCESSO Then gError 102604
        
        'Traz os contatos para a tela
        Call BotaoTrazer_Click
            
        iClienteAlterado = 0
        
    End If
    
    Exit Sub
    
Erro_Cliente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102604
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154207)

    End Select

End Sub

Private Function Valida_Cliente()

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Valida_Cliente

    If iClienteAlterado = 0 Then Exit Function

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 102599

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 102600

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliente, iCodFilial)
        
    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialCliente.Clear

    End If
    
    giProxCod = 1
    
    iClienteAlterado = 0

    Valida_Cliente = SUCESSO
    
    Exit Function

Erro_Valida_Cliente:

    Valida_Cliente = gErr
    
    Select Case gErr

        Case 102599, 102600
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154208)

    End Select

    Exit Function

End Function

Private Sub FilialCliente_Change()
    iFilialClienteAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialCliente_Click()

On Error GoTo Erro_FilialCliente_Click

    'Se é para atualizar o grid
    'Essa flag é usada para evitar que no sistema de setas,
    'enquanto está preenchendo o campo filial, o sistema dispare o click
    'diversas vezes e com isso faça a carga do grid repetidamente sem necessidade
    If gbAtualizarGrid Then
    
        'Traz os contatos para a tela
        Call BotaoTrazer_Click
    
    End If
    
    Exit Sub
    
Erro_FilialCliente_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154209)

    End Select

End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Validate

    If iFilialClienteAlterado = REGISTRO_ALTERADO Then
    
        'Executa a validação do cliente
        lErro = Valida_FilialCliente
        If lErro <> SUCESSO Then gError 102611
        
        'Traz os contatos para a tela
        Call BotaoTrazer_Click
    
        iFilialClienteAlterado = 0
        
    End If
    
    Exit Sub
    
Erro_FilialCliente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102611
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154210)

    End Select

End Sub

Private Function Valida_FilialCliente()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sNomeRed As String
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Valida_FilialCliente

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialCliente.Text)) = 0 Then Exit Function

    'Verifica se está preenchida com o item selecionado na ComboBox Filial
    If FilialCliente.Text = FilialCliente.List(FilialCliente.ListIndex) Then Exit Function

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(FilialCliente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102515

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Cliente
        If Len(Trim(Cliente.Text)) = 0 Then gError 102516

        'Lê o Cliente que está na tela
        sNomeRed = Trim(Cliente.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Cliente e Código da Filial
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 102517

        'Se não existe a Filial
        If lErro = 17660 Then gError 102518

        'Encontrou Filial no BD, coloca no Text da Combo
        FilialCliente.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 102519
    
    giProxCod = 1

    Valida_FilialCliente = SUCESSO
    
    Exit Function

Erro_Valida_FilialCliente:

    Valida_FilialCliente = gErr

    Select Case gErr

        Case 102515, 102517

        Case 102516
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 102518
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE1", FilialCliente.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
                'Segura o foco
            End If

        Case 102519
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154211)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Aciona rotina de gravação de registros

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Controla toda a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 121203

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 121203

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154212)

     End Select

     Exit Sub

End Sub

Private Function Move_Tela_Memoria(colClienteContatos As Collection) As Long
'Carrega na memória os dados que estão na tela

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos
Dim iIndice As Integer
Dim lCliente As Long
Dim iFilialCliente As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Obtém código do cliente e da filial que estão na tela
    lErro = Obtem_Cod_Cliente_Filial(objClienteContatos)
    If lErro <> SUCESSO Then gError 102603
    
    'Guarda código do cliente e da filial para serem passados para todos os registros
    lCliente = objClienteContatos.lCliente
    iFilialCliente = objClienteContatos.iFilialCliente
    
    'Para cada linha do Grid
    For iIndice = 1 To objGridContatos.iLinhasExistentes

        'Instancia um novo objClienteContatos
        Set objClienteContatos = New ClassClienteContatos

        'Guarda cliente e filial cliente no obj
        objClienteContatos.lCliente = lCliente
        objClienteContatos.iFilialCliente = iFilialCliente
        
        'Move os dados do grid para o obj
        objClienteContatos.iCodigo = StrParaInt(GridContatos.TextMatrix(iIndice, iGrid_Codigo_Col))
        objClienteContatos.sContato = Trim(GridContatos.TextMatrix(iIndice, iGrid_Contato_Col))
        objClienteContatos.sSetor = Trim(GridContatos.TextMatrix(iIndice, iGrid_Setor_Col))
        objClienteContatos.sCargo = Trim(GridContatos.TextMatrix(iIndice, iGrid_Cargo_Col))
        objClienteContatos.sTelefone = Trim(GridContatos.TextMatrix(iIndice, iGrid_Telefone_Col))
        objClienteContatos.sFax = Trim(GridContatos.TextMatrix(iIndice, iGrid_Fax_Col))
        objClienteContatos.sEmail = Trim(GridContatos.TextMatrix(iIndice, iGrid_Email_Col))
        objClienteContatos.iPadrao = StrParaInt(GridContatos.TextMatrix(iIndice, iGrid_Padrao_Col))
        objClienteContatos.dtDataNasc = StrParaDate(GridContatos.TextMatrix(iIndice, iGrid_DataNasc_Col))
        objClienteContatos.sOutrosMeioComunic = Trim(GridContatos.TextMatrix(iIndice, iGrid_OutMeioComuni_Col))
        
        'Guarda o obj na coleção
        colClienteContatos.Add objClienteContatos

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 102603
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154213)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Controla toda a rotina de gravação

Dim lErro As Long
Dim colClienteContatos As New Collection

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'verifica validade do GridContatos
    lErro = Valida_GridContatos()
    If lErro <> SUCESSO Then gError 121204

    'Move os valores da tela para a colContatoGeral
    lErro = Move_Tela_Memoria(colClienteContatos)
    If lErro <> SUCESSO Then gError 121205

    'Aciona rotinas de gravação no BD
    lErro = CF("ClienteContatos_Grava", colClienteContatos)
    If lErro <> SUCESSO Then gError 121206

    Call Limpa_ClienteContatos
    
    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 121204 To 121206

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154214)

     End Select

     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se existe algo para ser salvo antes de limpar a tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 121207

    'Limpa a Tela
    Call Limpa_ClienteContatos

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 121207

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154215)

     End Select

     Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 121208

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 121208

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154216)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Exclui os Contatos para esse ID e tipo de contato

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_BotaoExcluir_Click

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o cliente não foi preenchido erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 102576
    
    'Se a filial do cliente não foi preenchida => erro
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 102577
    
    'Pede confirmação para exclusão ao usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CLIENTECONTATOS", Cliente.Text, FilialCliente.Text)

    'Se confirma
    If vbMsgRes = vbYes Then

        'Obtém código do cliente e da filial que estão na tela
        lErro = Obtem_Cod_Cliente_Filial(objClienteContatos)
        If lErro <> SUCESSO Then gError 102609
        
        'exclui os contatos
        lErro = CF("ClienteContatos_Exclui", objClienteContatos)
        If lErro <> SUCESSO Then gError 121209

        'Limpa o GridContatos
        Call Grid_Limpa(objGridContatos)

        iAlterado = 0

    End If

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 121209, 102609

        Case 102576
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102577
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154217)

    End Select

    Exit Sub

End Sub


Public Sub GridContatos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If

End Sub

Public Sub GridContatos_GotFocus()
    Call Grid_Recebe_Foco(objGridContatos)
End Sub

Public Sub GridContatos_EnterCell()
    Call Grid_Entrada_Celula(objGridContatos, iAlterado)
End Sub

Public Sub GridContatos_LeaveCell()
    Call Saida_Celula(objGridContatos)
End Sub

Public Sub GridContatos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
    iLinhasExistentesAnterior = objGridContatos.iLinhasExistentes
    iItemAtual = GridContatos.Row
    Call Grid_Trata_Tecla1(KeyCode, objGridContatos)
    If objGridContatos.iLinhasExistentes < iLinhasExistentesAnterior Then
        GridContatos.TextMatrix(objGridContatos.iLinhasExistentes + 1, iGrid_Codigo_Col) = GridContatos.TextMatrix(iLinhasExistentesAnterior + 1, iGrid_Codigo_Col)
        GridContatos.TextMatrix(iLinhasExistentesAnterior + 1, iGrid_Codigo_Col) = ""
    End If
End Sub

Public Sub GridContatos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContatos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContatos, iAlterado)
    End If

End Sub

Public Sub GridContatos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContatos)
End Sub

Public Sub GridContatos_RowColChange()
    Call Grid_RowColChange(objGridContatos)
End Sub

Public Sub GridContatos_Scroll()
    Call Grid_Scroll(objGridContatos)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Contato
        Case iGrid_Contato_Col
            lErro = Saida_Celula_Contato(objGridInt)
            If lErro <> SUCESSO Then gError 121216

        'Setor
        Case iGrid_Setor_Col
            lErro = Saida_Celula_Setor(objGridInt)
            If lErro <> SUCESSO Then gError 121217

        'Telefone
        Case iGrid_Telefone_Col
            lErro = Saida_Celula_Telefone(objGridInt)
            If lErro <> SUCESSO Then gError 121218

        'Fax
        Case iGrid_Fax_Col
            lErro = Saida_Celula_Fax(objGridInt)
            If lErro <> SUCESSO Then gError 121219

        'Email
        Case iGrid_Email_Col
            lErro = Saida_Celula_Email(objGridInt)
            If lErro <> SUCESSO Then gError 121220

        'Cargo
        Case iGrid_Cargo_Col
            lErro = Saida_Celula_Cargo(objGridInt)
            If lErro <> SUCESSO Then gError 121221

        'DataNasc
        Case iGrid_DataNasc_Col
            lErro = Saida_Celula_Data(objGridInt)
            If lErro <> SUCESSO Then gError 121221

        'Outros Meios
        Case iGrid_OutMeioComuni_Col
            lErro = Saida_Celula_Padrao(objGridInt, OutrosMeioComunic)
            If lErro <> SUCESSO Then gError 121221

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121228

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 121216 To 121228

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154218)

    End Select

    Exit Function

End Function

Private Sub Padrao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub Padrao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub Padrao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = Padrao
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Cargo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cargo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub Cargo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub Cargo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = Cargo
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Contato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Contato_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub Contato_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub Contato_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = Contato
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Telefone_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Telefone_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub Telefone_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub Telefone_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = Telefone
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Fax_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Fax_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub Fax_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub Fax_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = Fax
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Email_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Email_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub Email_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub Email_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = Email
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Setor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Setor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub Setor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub Setor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = Setor
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Contato(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Contato que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Contato

    Set objGridInt.objControle = Contato

    If GridContatos.Row - GridContatos.FixedRows = objGridInt.iLinhasExistentes And Len(Trim(Contato.Text)) > 0 Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        GridContatos.TextMatrix(GridContatos.Row + 1, iGrid_Codigo_Col) = CStr(giProxCod)
        giProxCod = giProxCod + 1
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121235

    Saida_Celula_Contato = SUCESSO

    Exit Function

Erro_Saida_Celula_Contato:

    Saida_Celula_Contato = gErr

    Select Case gErr

        Case 121235
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154219)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Setor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Setor que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Setor

    Set objGridInt.objControle = Setor

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121236

    Saida_Celula_Setor = SUCESSO

    Exit Function

Erro_Saida_Celula_Setor:

    Saida_Celula_Setor = gErr

    Select Case gErr

        Case 121236
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154220)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Telefone(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Telefone que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Telefone

    Set objGridInt.objControle = Telefone

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121237

    Saida_Celula_Telefone = SUCESSO

    Exit Function

Erro_Saida_Celula_Telefone:

    Saida_Celula_Telefone = gErr

    Select Case gErr

        Case 121237
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154221)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Fax(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Fax que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Fax

    Set objGridInt.objControle = Fax

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121238

    Saida_Celula_Fax = SUCESSO

    Exit Function

Erro_Saida_Celula_Fax:

    Saida_Celula_Fax = gErr

    Select Case gErr

        Case 121238
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154222)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Email(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Email que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Email

    Set objGridInt.objControle = Email

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121239

    Saida_Celula_Email = SUCESSO

    Exit Function

Erro_Saida_Celula_Email:

    Saida_Celula_Email = gErr

    Select Case gErr

        Case 121239
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154223)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cargo(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Cargo que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Cargo

    Set objGridInt.objControle = Cargo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 121241

    Saida_Celula_Cargo = SUCESSO

    Exit Function

Erro_Saida_Celula_Cargo:

    Saida_Celula_Cargo = gErr

    Select Case gErr

        Case 121241
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154224)

    End Select

    Exit Function

End Function

Private Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelCliente_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_LabelCliente_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154225)
    
    End Select
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
    
    End If

End Sub

'**** TRATAMENTO DO SISTEMA DE SETAS ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objClienteContatos As New ClassClienteContatos
Dim objCampoValor As AdmCampoValor
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ClienteContatos_ClientesDistintos"

    'Obtém código do cliente e da filial que estão na tela
    lErro = Obtem_Cod_Cliente_Filial(objClienteContatos)
    If lErro <> SUCESSO Then gError 102617

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Cliente", objClienteContatos.lCliente, 0, "Cliente"
    colCampoValor.Add "NomeCliente", Trim(Cliente.Text), STRING_CLIENTE_NOME_REDUZIDO, "NomeCliente"
    colCampoValor.Add "FilialCliente", objClienteContatos.iFilialCliente, 0, "FilialCliente"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case 102617
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154226)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD
'Esse Tela_Preenche foge do padrão, pois selecionar o campo
'na combo Campo é suficiente para que o restante da tela seja preenchido

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos
Dim iIndice As Integer

On Error GoTo Erro_Tela_Preenche

    'Evita que o grid seja atualizado enquanto está preenchendo
    'e validando a filial do cliente
    gbAtualizarGrid = False
    
    'Guarda o código do campo em questão no obj
    objClienteContatos.lCliente = colCampoValor.Item("Cliente").vValor
    objClienteContatos.iFilialCliente = colCampoValor.Item("FilialCliente").vValor

    'Exibe o cliente na tela e faz a validação do mesmo
    Cliente.Text = objClienteContatos.lCliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102615
    
    'Exibe o cliente na tela e faz a validação do mesmo
    FilialCliente.Text = objClienteContatos.iFilialCliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 102616

    'Preenche a tela com os valores para o campo em questão
    lErro = Traz_ClienteContatos_Tela(objClienteContatos)
    If lErro <> SUCESSO Then gError 102608
    
    gbAtualizarGrid = True
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 102607, 102608, 102615, 102616
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154227)

    End Select

    Exit Sub

End Sub
'**** FIM DO TRATAMENTO DO SISTEMA DE SETAS ****

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    
    Caption = "Clientes x Contatos"
    
    Call Form_Load

End Function

Public Function Name() As String
    Name = "ClienteContatos"
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

Private Sub Unload(objme As Object)

 RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Function Obtem_Cod_Cliente_Filial(objClienteContatos As ClassClienteContatos) As Long
'Obtém o código do cliente e da filial que estão na tela e guarda-os no objClienteContatos

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_Obtem_Cod_Cliente_Filial

    'Se o cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        '*** Leitura do cliente a partir do nome reduzido para obter o seu código ***
        
        'Guarda o nome reduzido do cliente
        objcliente.sNomeReduzido = Trim(Cliente.Text)
        
        'Faz a leitura do cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 102578
        
        'Se não encontrou o cliente => erro
        If lErro = 12348 Then gError 102579
        
        'Guarda no obj o código do cliente
        objClienteContatos.lCliente = objcliente.lCodigo
        
        '*** Fim da leitura de cliente ***
        
        'Guarda no obj o código da filial do cliente
        objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    
    End If

    Obtem_Cod_Cliente_Filial = SUCESSO

    Exit Function

Erro_Obtem_Cod_Cliente_Filial:

    Obtem_Cod_Cliente_Filial = gErr

    Select Case gErr

        Case 102578

        Case 102579
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154228)

    End Select

End Function

'*** ROTINA_GRID_ENABLE ************
Public Sub Rotina_Grid_Enable(ByVal iLinha As Integer, ByVal objControl As Object, ByVal iLocalChamada As Integer)

Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Seleciona o controle atual
    Select Case objControl.Name

        'se for o campo código
        Case Contato.Name
            
            'Se o campo contato não estiver preenchido e os campos cliente e filial estiverem preenchidos
            If Len(Trim(GridContatos.TextMatrix(GridContatos.Row, iGrid_Contato_Col))) = 0 And Len(Trim(Cliente.Text)) > 0 And Len(Trim(FilialCliente.Text)) > 0 Then
            
                'habilita o controle
                objControl.Enabled = True
            
            'Senão
            Else
                
                'Desabilita o controle
                objControl.Enabled = False
            
            End If
        
        'Se for qualquer campo do grid diferente do campo conato
        Case Setor.Name, Cargo.Name, Telefone.Name, Fax.Name, Email.Name
        
            'Se o campo contato estiver preenchido
            If Len(Trim(GridContatos.TextMatrix(GridContatos.Row, iGrid_Contato_Col))) > 0 Then
            
                'habilita o controle
                objControl.Enabled = True
            
            'Senão
            Else
                
                'Desabilita o controle
                objControl.Enabled = False
            
            End If
            
        
    End Select
    
    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154229)
            
    End Select
    
    Exit Sub

End Sub
'*****************

Private Sub Padrao_Click()
    Call Trata_Item_Padrao
End Sub

Private Function Trata_Item_Padrao() As Long
'Impede que 2 ou mais itens sejam configurados como item padrão

Dim iLinha As Integer

On Error GoTo Erro_Trata_Item_Padrao

    'Para cada item do Grid
    For iLinha = 1 To objGridContatos.iLinhasExistentes
    
        'Se o item estiver configurado como padrão e não for o item da linha atual
        If StrParaInt(GridContatos.TextMatrix(iLinha, iGrid_Padrao_Col)) = MARCADO And iLinha <> GridContatos.Row Then
        
            'desmarca a opção padrão para esse item, pois apenas um item pode ser considerado padrão
            GridContatos.TextMatrix(iLinha, iGrid_Padrao_Col) = DESMARCADO
        
        End If
    
    Next
    
    'Faz um refresh no grid para atualizar as figuras de marcado / desmarcado
    Call Grid_Refresh_Checkbox(objGridContatos)

    Trata_Item_Padrao = SUCESSO
    
    Exit Function

Erro_Trata_Item_Padrao:

    Trata_Item_Padrao = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154230)

    End Select

End Function

Private Sub Cliente_Preenche()

Dim sNomeReduzidoParte As String
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134018

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134018

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154231)

    End Select
    
    Exit Sub

End Sub

Private Sub DataNasc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataNasc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub DataNasc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub DataNasc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = DataNasc
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub OutrosMeioComunic_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OutrosMeioComunic_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContatos)
End Sub

Private Sub OutrosMeioComunic_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContatos)
End Sub

Private Sub OutrosMeioComunic_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridContatos.objControle = OutrosMeioComunic
    lErro = Grid_Campo_Libera_Foco(objGridContatos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Data(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = DataNasc

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataNasc.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataNasc.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188450)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function
