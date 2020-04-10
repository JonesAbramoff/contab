VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpComisPart 
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   ScaleHeight     =   3540
   ScaleWidth      =   8250
   Begin VB.Frame Frame1 
      Caption         =   "Vendedores"
      Height          =   1005
      Left            =   375
      TabIndex        =   20
      Top             =   1515
      Width           =   5355
      Begin VB.OptionButton OptionApenasUm 
         Caption         =   "Apenas um"
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
         Left            =   165
         TabIndex        =   23
         Top             =   645
         Width           =   1545
      End
      Begin VB.OptionButton OptionTodos 
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
         Height          =   240
         Left            =   165
         TabIndex        =   21
         Top             =   255
         Value           =   -1  'True
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   3060
         TabIndex        =   22
         Top             =   600
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Enabled         =   0   'False
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
         Left            =   2100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   660
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   825
      Left            =   375
      TabIndex        =   15
      Top             =   2535
      Width           =   5355
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   16
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3255
         TabIndex        =   17
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   360
         Width           =   360
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Baixa"
      Height          =   735
      Left            =   375
      TabIndex        =   8
      Top             =   735
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownBaixaDe 
         Height          =   315
         Left            =   2235
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox BaixaDe 
         Height          =   285
         Left            =   1050
         TabIndex        =   10
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownBaixaAte 
         Height          =   315
         Left            =   4500
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox BaixaAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   12
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2940
         TabIndex        =   14
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   330
         Width           =   285
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpComisPartOcx.ctx":0000
      Left            =   2340
      List            =   "RelOpComisPartOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   240
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
      Left            =   6150
      Picture         =   "RelOpComisPartOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   825
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5970
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpComisPartOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpComisPartOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpComisPartOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpComisPartOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1605
      TabIndex        =   7
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpComisPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 123159

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 123160

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 123159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 123160

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 123161

    ComboOpcoes.Text = ""

    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 123162

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 123161, 123162

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError 123163

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case 123163

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

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
        If lErro <> SUCESSO Then gError 123164

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 123164

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 123165

    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 123165

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 123166

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 123166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 123167

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 123168

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 123169

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 123170

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 123167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 123168, 123169, 123170

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
    If ComboOpcoes.ListIndex = -1 Then gError 123171

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 123172

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 123171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 123172

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 123173

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 123173

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'Preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sCheckVend As String
Dim sVendedor As String

On Error GoTo Erro_PreencherRelOp

    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sVendedor, sCheckVend)
    If lErro <> SUCESSO Then gError 123174

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 123175

    'Preenche o Cliente Inicial
    lErro = objRelOpcoes.IncluirParametro("NCLINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 123176

    lErro = objRelOpcoes.IncluirParametro("TCLIINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 123177
    
    'Preenche o Cliente Final
    lErro = objRelOpcoes.IncluirParametro("NCLIFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 123178

    lErro = objRelOpcoes.IncluirParametro("TCLIFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 123179
    
    'Preenche Baixa Inicial
    If BaixaDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", BaixaDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 123180

    'Preenche Baixa Final
    If BaixaAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", BaixaAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 123181
    
    'Preenche o Vendedor
    If sCheckVend = "UM" Then
    
        lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", sVendedor)
        If lErro <> AD_BOOL_TRUE Then Error 123182
    
        lErro = objRelOpcoes.IncluirParametro("TOPVEND", sCheckVend)
        If lErro <> AD_BOOL_TRUE Then Error 123183
    
    End If
        
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_I, sCliente_F, sVendedor, sCheckVend)
    If lErro <> SUCESSO Then gError 123184

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 123174 To 123184

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sVendedor As String, sCheckVend As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Vendedor

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica Cliente Inicial e Final
    If ClienteInicial.Text <> "" Then
        sCliente_I = CStr(LCodigo_Extrai(ClienteInicial.Text))
    Else
        sCliente_I = ""
    End If

    If ClienteFinal.Text <> "" Then
        sCliente_F = CStr(LCodigo_Extrai(ClienteFinal.Text))
    Else
        sCliente_F = ""
    End If

    If sCliente_I <> "" And sCliente_F <> "" Then

        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 123185

    End If

    'data inicial não pode ser maior que a data final
    If Trim(BaixaDe.ClipText) <> "" And Trim(BaixaAte.ClipText) <> "" Then

         If CDate(BaixaDe.Text) > CDate(BaixaAte.Text) Then gError 123186

    End If

    'Se a opção para todos os vendedores estiver selecionada
    If OptionTodos.Value = True Then
        sCheckVend = "TODOS"
    Else
        'Se a opção para apenas um vendedor estiver selecionada
        If Len(Trim(Vendedor.Text)) = 0 Then gError 123187
        sVendedor = Vendedor.Text
        sCheckVend = "UM"
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 123185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus

        Case 123186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_INICIAL_MAIOR", gErr)
            BaixaDe.SetFocus

        Case 123187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_PREENCHIDO", gErr)
            Vendedor.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_I As String, sCliente_F As String, sVendedor As String, sCheckVend As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'Coloca na Expressão o Cliente Inicial e Final
    If sCliente_I <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(CLng(sCliente_I))

    If sCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_F))

    End If

    'Se a opção para apenas um vendedor estiver selecionada
    If sCheckVend = "UM" Then

        If Len(Trim(sVendedor)) > 0 Then sExpressao = "Vendedor = " & Forprint_ConvInt(Codigo_Extrai(sVendedor))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 123188

    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLINIC", sParam)
    If lErro <> SUCESSO Then gError 123189

    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)

    'pega Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIFIM", sParam)
    If lErro <> SUCESSO Then gError 123190

    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)

    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 123191

    Call DateParaMasked(BaixaDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 123192

    Call DateParaMasked(BaixaAte, CDate(sParam))
    
    'pega vendedor e exibe
    lErro = objRelOpcoes.ObterParametro("TOPVEND", sParam)
    If lErro <> SUCESSO Then Error 123193
    
    'Se a Opção todos Vendedores estiver preenchida
    If sParam = "UM" Then

        lErro = objRelOpcoes.ObterParametro("TVENDEDOR", sParam)
        If lErro <> SUCESSO Then Error 123194

        OptionApenasUm.Value = True
        Vendedor.Text = sParam
        Call Vendedor_Validate(bSGECancelDummy)

    ElseIf sParam = "Todos" Then

        OptionTodos.Value = True
        Vendedor.Enabled = False
        LabelVendedor.Enabled = False

    End If

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 123188 To 123194

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    BaixaDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    BaixaAte.Text = Format(gdtDataAtual, "dd/mm/yy")

    'defina todos os tipos
    OptionTodos.Value = True

    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub BaixaAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(BaixaAte)

End Sub

Private Sub BaixaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaixaAte_Validate

    If Len(BaixaAte.ClipText) > 0 Then

        lErro = Data_Critica(BaixaAte.Text)
        If lErro <> SUCESSO Then gError 123195

    End If

    Exit Sub

Erro_BaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 123195

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BaixaDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(BaixaDe)

End Sub

Private Sub BaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaixaDe_Validate

    If Len(BaixaDe.ClipText) > 0 Then

        lErro = Data_Critica(BaixaDe.Text)
        If lErro <> SUCESSO Then gError 123196

    End If

    Exit Sub

Erro_BaixaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 123196

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set objEventoVendedor = Nothing

End Sub

Private Sub OptionApenasUm_Click()

    Vendedor.Enabled = True
    LabelVendedor.Enabled = True

End Sub

Private Sub OptionTodos_Click()

    Vendedor.Text = ""
    Vendedor.Enabled = False
    LabelVendedor.Enabled = False

End Sub

Private Sub UpDownBaixaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_DownClick

    lErro = Data_Up_Down_Click(BaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 123197

    Exit Sub

Erro_UpDownBaixaDe_DownClick:

    Select Case Err

        Case 123197
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_UpClick

    lErro = Data_Up_Down_Click(BaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 123198

    Exit Sub

Erro_UpDownBaixaDe_UpClick:

    Select Case gErr

        Case 123198
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaAte_DownClick

    lErro = Data_Up_Down_Click(BaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 123199

    Exit Sub

Erro_UpDownBaixaAte_DownClick:

    Select Case gErr

        Case 123199
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaAte_UpClick

    lErro = Data_Up_Down_Click(BaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 123200

    Exit Sub

Erro_UpDownBaixaAte_UpClick:

    Select Case Err

        Case 123201
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Comissões por Participante"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpComisPart"

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

Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click
    
    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche o Cliente Final com o Codigo selecionado
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    'Preenche o Cliente Final com Codigo - Descricao
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche o Cliente Inical com o codigo
    ClienteInicial.Text = CStr(objCliente.lCodigo)
    
    'Preenche o Cliente Inicial com codigo - Descricao
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub LabelVendedor_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_LabelVendedor_Click
    
    If Len(Trim(Vendedor.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

   Exit Sub

Erro_LabelVendedor_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub


