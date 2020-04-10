VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpPainelCaixas 
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   6705
   Begin VB.CheckBox ApenasCaixaCentral 
      Caption         =   "Apenas Caixa Central"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   2205
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   4095
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   285
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   885
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelCaixaDe 
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
         Height          =   195
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelCaixaAte 
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
         Height          =   195
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   945
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPainelCaixas.ctx":0000
      Left            =   1080
      List            =   "RelOpPainelCaixas.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPainelCaixas.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpPainelCaixas.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpPainelCaixas.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpPainelCaixas.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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
      Left            =   4733
      Picture         =   "RelOpPainelCaixas.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   945
      Width           =   1605
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
      TabIndex        =   9
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPainelCaixas"
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

'Obj utilizado para o browser de Caixas
Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1

Dim giCaixaInicial As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCaixa = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170573)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Limpa Objetos da memoria
    Set objEventoCaixa = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116549

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116550

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 116550

        Case 116549
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170574)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Aciona a Rotina de exclusão das opções de relatório

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116551

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116552

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116551
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116552

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170575)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()
'Aciona rotinas que que checam as opções do relatório e ativam impressão do mesmo

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'aciona rotina que checa opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116555
    
    'Se a checkBox Apenas Caixa Central não estiver setada
    If CInt(ApenasCaixaCentral.Value) = vbUnchecked Then
    
        'Guarda o nome do tsk
        gobjRelatorio.sNomeTsk = "painelcx"
        
    'Senão
    Else
    
        'Guarda o nome do tsk
        gobjRelatorio.sNomeTsk = "painelct"

    End If
    
    'Chama rotina que excuta a impressão do relatório
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 116555

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170576)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 116556

    'Chama rotina que checa as opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116557

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Aciona rotina que grava opções do relatório
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116558

    'Testa se nome no combo esta igual ao nome em gobjRelOpçoes.sNome
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116559

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116556
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116557 To 116559

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170577)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'Aciona Rotinas de Limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116560

    'Limpa Combo Opções
    ComboOpcoes.Text = ""
    
    ComboOpcoes.SetFocus
    
    'Desmarca opção Apenas Caixa Central
    ApenasCaixaCentral.Value = Unchecked
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 116560

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170578)

    End Select

    Exit Sub

End Sub

Private Sub ApenasCaixaCentral_Click()

    If ApenasCaixaCentral.Value = Checked Then
        
        'Limpa as TextBox
        CaixaDe.Text = ""
        CaixaAte.Text = ""
        
        'Desabilita a label e as textBox
        CaixaDe.Enabled = False
        LabelCaixaDe.Enabled = False
        CaixaAte.Enabled = False
        LabelCaixaAte.Enabled = False
    
    Else
    
        'Habilita as Labels e as TextBox
        CaixaDe.Enabled = True
        LabelCaixaDe.Enabled = True
        CaixaAte.Enabled = True
        LabelCaixaAte.Enabled = True
        
    End If
    
End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    'inicializa variavel bSGECancelDummy
    bSGECancelDummy = False
    
    'Limpa a Tela
    lErro = Limpa_Tela
    If lErro <> SUCESSO Then gError 116794
    
    'Carrega parametros do relatorio gravado
    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 116573
            
    'pega parâmetro Caixa Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAINIC", sParam)
    If lErro <> SUCESSO Then gError 116574
    
    'Preenche campo CaixaDe
    CaixaDe.Text = sParam
   
    'verifica validade de CaixaDe
    Call CaixaDe_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116683
    
    'pega parâmetro Caixa Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAFIM", sParam)
    If lErro <> SUCESSO Then gError 116575
    
    'Preenche campo CaixaAte
    CaixaAte.Text = sParam
    
    'verifica validade de CaixaAte
    Call CaixaAte_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then gError 116684
    
    'Pega Apenas Caixa Central
    lErro = objRelOpcoes.ObterParametro("NCAIXACENTRAL", sParam)
    If lErro <> SUCESSO Then gError 128145
    
    'verifica se Apenas Caixa Central esta marcado no relatorio carregado
    If sParam = Checked Then
        
        ApenasCaixaCentral.Value = Checked
    
    Else
    
        ApenasCaixaCentral.Value = Unchecked
        
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116573 To 116575
        
        Case 116794, 128145
        
        Case 116683
            CaixaDe.Text = ""
            
        Case 116684
            CaixaAte.Text = ""
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170579)

    End Select

    Exit Function

End Function

Private Sub CaixaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(CaixaAte)
End Sub

Private Sub CaixaDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(CaixaDe)
End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Private Sub CaixaAte_Validate(Cancel As Boolean)
'Verifica validade de CaixaAte

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaAte_Validate

    giCaixaInicial = 0

    If Len(Trim(CaixaAte.Text)) > 0 Then

        'instancia o obj
        Set objCaixa = New ClassCaixa

        'preenche o obj c/ o cod e filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler a Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaAte, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116562

        'código inexistente
        If lErro = 116175 Then gError 116563

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116563

    End If
 
    Exit Sub

Erro_CaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116562

        Case 116563
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, CaixaAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170580)

    End Select

End Sub

Private Sub CaixaDe_Validate(Cancel As Boolean)
'Verifica validade de CaixaDe

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaAte_Validate

    giCaixaInicial = 1

    If Len(Trim(CaixaDe.Text)) > 0 Then
        
        'instancia o obj
        Set objCaixa = New ClassCaixa
        
        'preenche o obj c/ o cod e filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaDe, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116564

        'código inexistente
        If lErro = 116175 Then gError 116565

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116565

    End If
    
    Exit Sub
    
Erro_CaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116564

        Case 116565
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, CaixaDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170581)

    End Select
    
    Exit Sub

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelCaixaAte_Click()

Dim objCaixa As New ClassCaixa
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCaixaAte_Click

    giCaixaInicial = 0
    
    sSelecao = "CaixaCod <> ?"
    colSelecao.Add CODIGO_CAIXA_CENTRAL
    
    If Len(Trim(CaixaAte.Text)) > 0 Then
        'Preenche com a caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
    End If
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)
    
    Else
    
        'Chama Tela Caixa
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)
    
    End If
    
    Exit Sub
    
Erro_LabelCaixaAte_Click:
    
    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170582)

    End Select

    Exit Sub

End Sub

Private Sub LabelCaixaAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelCaixaAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCaixaDe_Click()

Dim objCaixa As New ClassCaixa
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCaixaDe_Click

    giCaixaInicial = 1
    
    sSelecao = "CaixaCod <> ?"
    colSelecao.Add CODIGO_CAIXA_CENTRAL
    
    If Len(Trim(CaixaDe.Text)) > 0 Then
        'Preenche com a caixa  da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
    End If
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)
    
    Else
    
        'Chama Tela de caixa
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa, sSelecao)
    
    End If
    
    Exit Sub
    
Erro_LabelCaixaDe_Click:

   Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170583)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelCaixaDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Controle_MouseDown(LabelCaixaDe, Button, Shift, X, Y)
End Sub

Private Sub objEventoCaixa_evSelecao(obj1 As Object)

Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixa_evSelecao
    
    Set objCaixa = obj1

    'se controle atual é o CaixaDe
    If giCaixaInicial = 1 Then

        'Preenche campo CaixaDe
        CaixaDe.Text = CStr(objCaixa.iCodigo)
        
        Call CaixaDe_Validate(bSGECancelDummy)

    'Se controle atual é o CaixaAte
    Else

       'Preenche campo CaixaAte
       CaixaAte.Text = CStr(objCaixa.iCodigo)
       
       Call CaixaAte_Validate(bSGECancelDummy)

    End If

    Me.Show

    Exit Sub

Erro_objEventoCaixa_evSelecao:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170584)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Verifica se a tecla F3 (Browser) foi acionada
    If KeyCode = KEYCODE_BROWSER Then

        'Verifica se o campo atual é o CaixaDe ou o CaixaAte
        If Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click
        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click
        End If

    End If

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCaixa_I As String
Dim sCaixa_F As String

On Error GoTo Erro_PreencherRelOp

    'Verifica Parametros , e formata os mesmos
    lErro = Formata_E_Critica_Parametros(sCaixa_I, sCaixa_F)
    If lErro <> SUCESSO Then gError 116566

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116567

    lErro = objRelOpcoes.IncluirParametro("DDATAATUAL", gdtDataHoje)
    If lErro <> AD_BOOL_TRUE Then gError 116568
    
    'Inclui parametro de CaixaDe
    lErro = objRelOpcoes.IncluirParametro("NCAIXAINIC", sCaixa_I)
    If lErro <> AD_BOOL_TRUE Then gError 116568

    'Inclui parametro de CaixaDe
    lErro = objRelOpcoes.IncluirParametro("TCAIXAINIC", Trim(CaixaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 102468

    'Inclui parametro de CaixaAte
    lErro = objRelOpcoes.IncluirParametro("NCAIXAFIM", sCaixa_F)
    If lErro <> AD_BOOL_TRUE Then gError 116569

    'Inclui parametro de CaixaAte
    lErro = objRelOpcoes.IncluirParametro("TCAIXAFIM", Trim(CaixaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 102469

    If giFilialEmpresa <> EMPRESA_TODA Then
        
        'Inclui Parametro Filial Empresa
        lErro = objRelOpcoes.IncluirParametro("NFILIALEMPRESA", CStr(giFilialEmpresa))
        If lErro <> AD_BOOL_TRUE Then gError 116570

    End If
    
    'Inclui parametro de Apenas Caixa Central
    lErro = objRelOpcoes.IncluirParametro("NCAIXACENTRAL", CInt(ApenasCaixaCentral.Value))
    If lErro <> AD_BOOL_TRUE Then gError 128146
        
    'Aciona Rotina que monta_expressão que será usada para gerar relatório
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCaixa_I, sCaixa_F)
    If lErro <> SUCESSO Then gError 116571

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 116566 To 116571, 102468, 102469, 128146

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170585)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCaixa_I As String, sCaixa_F As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'Formata CaixaDe
    If CaixaDe.ClipText <> "" Then
        sCaixa_I = CStr(Codigo_Extrai(CaixaDe.Text))
    Else
        sCaixa_I = ""
    End If

    'Formata CaixaAte
    If CaixaAte.ClipText <> "" Then
        sCaixa_F = CStr(Codigo_Extrai(CaixaAte.Text))
    Else
        sCaixa_F = ""
    End If

    'verifica se CaixaDe é maior que o CaixaAte
    If Trim(CaixaDe.ClipText) <> "" And Trim(CaixaAte.ClipText) <> "" Then

         If CInt(sCaixa_I) > CInt(sCaixa_F) Then gError 116572

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 116572
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXAINICIAL_MAIOR_CAIXAFINAL", gErr)
            CaixaDe.SetFocus

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170586)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCaixa_I As String, sCaixa_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'Verifica se campo ClienteDe foi preenchido
    If Trim(CaixaDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressao o Valor de CaixaDe
        sExpressao = sExpressao & "Caixa >= " & sCaixa_I

    End If

    'Verifica se campo ClienteAte foi preenchido
    If Trim(CaixaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de ClienteAte
        sExpressao = sExpressao & "Caixa <= " & sCaixa_F

    End If

    If giFilialEmpresa <> EMPRESA_TODA Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de Filial Empresa
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)

    End If
    
    'Verifica se a expressão foi preenchido com algum filtro
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170587)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela()
'Limpa os campos da tela , quando é chamada uma opção de relatorio para a tela

On Error GoTo Erro_Limpa_Tela

    'Limpa campos de Caixa
    CaixaDe.Text = ""
    CaixaAte.Text = ""

    Exit Function
    
Erro_Limpa_Tela:

    Select Case gErr

        Case Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170588)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Painel de Caixas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPainelCaixas"

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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub LabelCaixaAte_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCaixaAte, Source, X, Y)
End Sub

Private Sub LabelCaixaDe_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCaixaDe, Source, X, Y)
End Sub

'Public Function TP_Caixa_Le(objCaixaMaskEdBox As Object, ByVal objCaixa As ClassCaixa) As Long
''Lê a Caixa com Código ou NomeRed em objCaixaMaskEdBox.Text
''Devolve em objCaixa. Coloca código-NomeReduzido no .Text
'
'Dim sCaixa As String
'Dim iCodigo As Integer
'Dim Caixa As Object
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_TP_Caixa_Le
'
'    Set Caixa = objCaixaMaskEdBox
'    sCaixa = Trim(Caixa.Text)
'
'    'Tenta extrair código de sCaixa
'    iCodigo = Codigo_Extrai(sCaixa)
'
'    'Indica a filial empresa ativa
'    objCaixa.iFilialEmpresa = giFilialEmpresa
'
'    'Se é do tipo código
'    If iCodigo > 0 Then
'
'        objCaixa.iCodigo = iCodigo
'
'        'verifica se o codigo existe
'        lErro = CF("Caixas_Le", objCaixa)
'        If lErro <> SUCESSO And lErro <> 79405 Then gError 116174
'
'        'sem dados
'        If lErro = 79405 Then gError 116175
'
'        Caixa.Text = objCaixa.iCodigo & SEPARADOR & objCaixa.sNomeReduzido
'
'    Else  'Se é do tipo String
'
'         objCaixa.sNomeReduzido = sCaixa
'
'         'verifica se o nome reduzido existe
'         lErro = CF("Caixa_Le_NomeReduzido", objCaixa)
'         If lErro <> SUCESSO And lErro <> 79582 Then gError 116176
'
'         'sem dados
'         If lErro = 79582 Then gError 116177
'
'         'NomeControle.text recebe codigo - nome_reduzido
'         Caixa.Text = objCaixa.iCodigo & SEPARADOR & objCaixa.sNomeReduzido
'
'    End If
'
'    TP_Caixa_Le = SUCESSO
'
'    Exit Function
'
'Erro_TP_Caixa_Le:
'
'    TP_Caixa_Le = gErr
'
'    Select Case gErr
'
'        Case 116176, 116174 'Tratados nas rotinas chamadas
'
'        Case 116175, 116177 'Caixa com Codigo / NomeReduzido não cadastrado
'
'        Case Else
'             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170589)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
