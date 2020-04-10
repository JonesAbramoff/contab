VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpPosPADataOcx 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   ScaleHeight     =   3285
   ScaleWidth      =   6495
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4260
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPosPADataOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPosPADataOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPosPADataOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPosPADataOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox CheckPulaPag 
      Caption         =   "Pula página a cada novo fornecedor"
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
      Left            =   180
      TabIndex        =   7
      Top             =   2880
      Width           =   3660
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
      Left            =   4380
      Picture         =   "RelOpPosPADataOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPosPADataOcx.ctx":0A96
      Left            =   900
      List            =   "RelOpPosPADataOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   285
      Width           =   3060
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fornecedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   180
      TabIndex        =   0
      Top             =   1425
      Width           =   3825
      Begin MSMask.MaskEdBox FornecedorInicial 
         Height          =   300
         Left            =   750
         TabIndex        =   1
         Top             =   315
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox FornecedorFinal 
         Height          =   300
         Left            =   765
         TabIndex        =   2
         Top             =   780
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label LabelFornFim 
         Caption         =   "Final:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LabelFornInic 
         Caption         =   "Inicial:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   1995
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   855
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   855
      TabIndex        =   14
      Top             =   855
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data:"
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
      Left            =   180
      TabIndex        =   16
      Top             =   870
      Width           =   525
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
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   330
      Width           =   690
   End
End
Attribute VB_Name = "RelOpPosPADataOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giFocoInicial As Boolean
Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private WithEvents objEventoFornecedorInic As AdmEvento
Attribute objEventoFornecedorInic.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorFim As AdmEvento
Attribute objEventoFornecedorFim.VB_VarHelpID = -1


Private Sub FornecedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorFinal_Validate

    If Len(Trim(FornecedorFinal.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(FornecedorFinal, objFornecedor, 0)
        If lErro <> SUCESSO Then Error 48776

    End If

    Exit Sub

Erro_FornecedorFinal_Validate:

    Cancel = True


    Select Case Err

        Case 48776

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171244)

    End Select

End Sub

Private Sub FornecedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorInicial_Validate

    If Len(Trim(FornecedorInicial.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(FornecedorInicial, objFornecedor, 0)
        If lErro <> SUCESSO Then Error 48777

    End If

    Exit Sub

Erro_FornecedorInicial_Validate:

    Cancel = True


    Select Case Err

        Case 48777

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171245)

    End Select

End Sub

Private Sub LabelFornInic_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornInic_Click

    If Len(Trim(FornecedorInicial.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorInicial.Text)
    End If

    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorInic)

   Exit Sub

Erro_LabelFornInic_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171246)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornFim_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornFim_Click

    If Len(Trim(FornecedorFinal.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorFinal.Text)
    End If

    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorFim)

   Exit Sub

Erro_LabelFornFim_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171247)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedorFim_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    FornecedorFinal.Text = CStr(objFornecedor.lCodigo)
    Call FornecedorFinal_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornecedorInic_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    FornecedorInicial.Text = CStr(objFornecedor.lCodigo)
    Call FornecedorInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Function PreencheComboOpcoes(sCodRel As String) As Long
'preenche o Combo de Opções com as opções referentes a sCodRel

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_PreencheComboOpcoes

    'le os nomes das opcoes do relatório existentes no BD
    lErro = CF("RelOpcoes_Le_Todos", sCodRel, colRelParametros)
    If lErro <> SUCESSO Then Error 23021

    'preenche o ComboBox com os nomes das opções do relatório
    For Each objRelOpcoes In colRelParametros
        ComboOpcoes.AddItem objRelOpcoes.sNome
    Next

    PreencheComboOpcoes = SUCESSO

    Exit Function

Erro_PreencheComboOpcoes:

    PreencheComboOpcoes = Err

    Select Case Err

        Case 23021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171248)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpRazao() As Long
'faz a crítica da data inicial e da data final

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_RelOpRazao

    'data não pode ser vazia
    If Len(Data.ClipText) = 0 Then Error 23025
    
    Critica_Datas_RelOpRazao = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpRazao:

    Critica_Datas_RelOpRazao = Err

    Select Case Err
        
        Case 23025
            Data.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171249)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 23027

    'pega Fornecedor Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TFORNINIC", sParam)
    If lErro <> SUCESSO Then Error 23040

    FornecedorInicial.Text = CStr(sParam)

    'pega Fornecedor Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TFORNFIM", sParam)
    If lErro <> SUCESSO Then Error 23041

    FornecedorFinal.Text = CStr(sParam)

    'pega 'Pula página a cada novo conta' e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPULAPAGQBR1", sParam)
    If lErro <> SUCESSO Then Error 23042

    If sParam = "S" Then CheckPulaPag.Value = 1

    'pega data inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 23043
    
    'pega data e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro <> SUCESSO Then Error 23044

    Data.PromptInclude = False
    Data.Text = sParam
    Data.PromptInclude = True

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23027, 23040, 23041, 23042, 23043, 23044

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171250)

    End Select

    Exit Function

End Function


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iPer As Integer
Dim iExercicio As Integer
Dim sCheck As String
Dim sDt As String
Dim sFornecedor_I As String, sFornecedor_F As String

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpRazao
    If lErro <> SUCESSO Then Error 23045

'    lErro = Obtem_Periodo_Exercicio(iPer, iExercicio, sDt)
'    If lErro <> SUCESSO Then Error 23046

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23047

    'Pegar parametros da tela
    sFornecedor_I = FornecedorInicial.Text
    lErro = objRelOpcoes.IncluirParametro("TFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then Error 23048

    sFornecedor_F = FornecedorFinal.Text
    lErro = objRelOpcoes.IncluirParametro("TFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then Error 23049

    sFornecedor_I = LCodigo_Extrai(FornecedorInicial.Text)
    lErro = objRelOpcoes.IncluirParametro("NFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then Error 23048

    sFornecedor_F = LCodigo_Extrai(FornecedorFinal.Text)
    lErro = objRelOpcoes.IncluirParametro("NFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then Error 23049

    'Pula Página a Cada Novo Fornecedor
    If CheckPulaPag.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If

    lErro = objRelOpcoes.IncluirParametro("TPULAPAGQBR1", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 23050

'    lErro = objRelOpcoes.IncluirParametro("NPERFIM", CStr(iPer))
'    If lErro <> AD_BOOL_TRUE Then Error 23052

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(iExercicio))
    If lErro <> AD_BOOL_TRUE Then Error 23053
      
    lErro = objRelOpcoes.IncluirParametro("DDATA", Data.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23056

    'Se fornecedor final preenchido
    If Len(Trim(FornecedorFinal.Text)) <> 0 Then

        'Verificar se Fornecedor Final é maior que Fornecedor Inicial
        If LCodigo_Extrai(FornecedorFinal.Text) < LCodigo_Extrai(FornecedorInicial.Text) Then Error 23061

    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sDt)
    If lErro <> SUCESSO Then Error 23058

    '???Call Acha_Nome_TSK(sDtIni_I)

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 23045, 23046, 23047, 23048, 23049, 23050

        Case 23052, 23053, 23056, 23057, 23058

        Case 23061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_FINAL_MENOR", Err, Error$)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171251)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sDt As String) As Long
'monta a expressão de seleção que será incluida dinamicamente para a execucao do relatorio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If FornecedorInicial.Text <> "" Then sExpressao = "Fornecedor >= " & Forprint_ConvLong(LCodigo_Extrai(FornecedorInicial.Text))

    If FornecedorFinal.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Fornecedor <= " & Forprint_ConvLong(LCodigo_Extrai(FornecedorFinal.Text))
    End If
    
'    'se a data não coincide com o período
'    If Data.Text <> sDt Then
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "LancData <= " & Forprint_ConvData(CDate(Data.Text))
'    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171252)

    End Select

    Exit Function

End Function


Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 24976

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 48773

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 48773

        Case 24976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171253)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 23038

    vbMsgRes = Rotina_Aviso(vbYesNo, "EXCLUSAO_RELOPRAZAOCP")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 23039

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 23038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 23039

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171254)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23037

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23037

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171255)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de opção"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 23034

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23035

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 23036

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 59495

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 23035, 23036, 59495

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171256)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    CheckPulaPag.Value = 0

    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

Dim lErro As Long

On Error GoTo Erro_ComboOpcoes_Click

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Le", gobjRelOpcoes)
    If (lErro <> SUCESSO) Then Error 23032

    lErro = PreencherParametrosNaTela(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23033

    Exit Sub

Erro_ComboOpcoes_Click:

    Select Case Err

        Case 23032, 23033

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171257)

    End Select

    Exit Sub

End Sub


Private Sub Data_Validate(Cancel As Boolean)

Dim sData As String
Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text
        lErro = Data_Critica(sData)
        If lErro <> SUCESSO Then Error 23031

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 23031

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171258)

    End Select

    Exit Sub

End Sub


Private Sub Form_Load()

Dim colCodigoDescricao As New AdmCollCodigoNome
Dim lErro As Long, iIndice As Integer
Dim objCodigoDescricao As AdmlCodigoNome

On Error GoTo Erro_OpcoesRel_Form_Load

    giFocoInicial = 1

    Set objEventoFornecedorInic = New AdmEvento
    Set objEventoFornecedorFim = New AdmEvento

'    'Preenche a listbox Fornecedores
'    'Le cada codigo e Nome Reduzido da tabela Fornecedores
'    lErro = CF("LCod_Nomes_Le","Fornecedores", "Codigo", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 23023
'
'    'preenche a listbox Fornecedores com os objetos da colecao colCodigoDescricao
'    For Each objCodigoDescricao In colCodigoDescricao
'
'        FornecedoresList.AddItem objCodigoDescricao.sNome
'        FornecedoresList.ItemData(FornecedoresList.NewIndex) = objCodigoDescricao.lCodigo
'
'    Next
'
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23022, 23023

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171259)

    End Select

    Unload Me

    Exit Sub

End Sub
'
'Private Sub FornecedorFinal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objFornecedor As New ClassFornecedor
'Dim iCodFilial As Integer
'
'On Error GoTo Erro_FornecedorFinal_Validate
'
'    giFocoInicial = 0
'
'    lErro = TP_Fornecedor_Le(FornecedorFinal, objFornecedor, iCodFilial)
'    If lErro Then Error 23059
'
'    Exit Sub
'
'Erro_FornecedorFinal_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 23059
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171260)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub FornecedorInicial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objFornecedor As New ClassFornecedor
'Dim iCodFilial As Integer
'
'On Error GoTo Erro_FornecedorInicial_Validate
'
'    giFocoInicial = 1
'
'    lErro = TP_Fornecedor_Le(FornecedorInicial, objFornecedor, iCodFilial)
'    If lErro <> SUCESSO Then Error 23060
'
'    Exit Sub
'Erro_FornecedorInicial_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 23060
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171261)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub FornecedoresList_DblClick()
'
'Dim sListBoxItem As String
'Dim lErro As Long
'
'On Error GoTo Erro_FornecedoresList_DblClick
'
'    'Se não há Fornecedor selecionado sai da rotina
'    If FornecedoresList.ListIndex = -1 Then Exit Sub
'
'    'Pega o nome reduzido do Fornecedor e joga no Fornecedor que teve o último foco
'    sListBoxItem = Trim(FornecedoresList.List(FornecedoresList.ListIndex))
'
'    'Verifica se o nome reduzido do fornecedor está vazio
'    If Len(sListBoxItem) = 0 Then Error 23067
'
'    If giFocoInicial = 0 Then
'
'        FornecedorFinal.Text = sListBoxItem
'        Exit Sub
'
'    End If
'
'    FornecedorInicial.Text = sListBoxItem
'
'    Exit Sub
'
'Erro_FornecedoresList_DblClick:
'
'    Select Case Err
'
'        Case 23067
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_VAZIO", Err, Error$)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171262)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'


Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23028

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 23028

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171263)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23062

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 23062

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171264)

    End Select

    Exit Sub

End Sub

'Function Obtem_Periodo_Exercicio(iPer As Integer, iExercicio As Integer, sDt As String) As Long
'
'Dim objPer As New ClassPeriodo
'Dim lErro As Long
'
'On Error GoTo Erro_Obtem_Periodo_Exercicio
'
'    'pega o período da Data Final
'    lErro = CF("Periodo_Le",CDate(Data.Text), objPer)
'    If lErro <> SUCESSO Then Error 23065
'
'    iPer = objPer.iPeriodo
'    iExercicio = objPer.iExercicio
'    sDt = objPer.dtData
'
'    Obtem_Periodo_Exercicio = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_Periodo_Exercicio:
'
'    Obtem_Periodo_Exercicio = Err
'
'    Select Case Err
'
'        Case 23065
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171265)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoFornecedorFim = Nothing
    Set objEventoFornecedorInic = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data)

End Sub



Private Sub LabelFornInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornInic, Source, X, Y)
End Sub

Private Sub LabelFornInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornInic, Button, Shift, X, Y)
End Sub

Private Sub LabelFornFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornFim, Source, X, Y)
End Sub

Private Sub LabelFornFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornFim, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID = IDH_RELOP_POSFORN
    Set Form_Load_Ocx = Me
    Caption = "Posição Contábil de Adiantamentos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpPosPAData"

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



