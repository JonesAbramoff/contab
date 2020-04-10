VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpRelCaixas 
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   6330
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   3495
      Begin VB.OptionButton StatusFechado 
         Caption         =   "Fechado"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton StatusTodos 
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
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton StatusAberto 
         Caption         =   "Aberto"
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
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelCaixas.ctx":0000
      Left            =   1080
      List            =   "RelOpRelCaixas.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   2670
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelCaixas.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpRelCaixas.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpRelCaixas.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpRelCaixas.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   11
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
      Left            =   4260
      Picture         =   "RelOpRelCaixas.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   945
      Width           =   1605
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3495
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   285
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   765
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         PromptChar      =   "_"
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
         Left            =   285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   825
         Width           =   360
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
      TabIndex        =   15
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelCaixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'evento do browser
Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1

'variavel de controle do browser
Dim giCaixaInicial As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'instancia o obj
    Set objEventoCaixa = New AdmEvento
              
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172448)

    End Select

    Exit Sub
    
End Sub

Private Function Formata_E_Critica_Parametros(sCaixaI As String, sCaixaF As String) As Long
'Verifica se o parâmetro inicial é maior que o final

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
         
    'critica Caixa Inicial e Final
    If Trim(CaixaDe.Text) <> "" Then
        sCaixaI = Codigo_Extrai(CaixaDe.Text)
    Else
        sCaixaI = ""
    End If
    
    If Trim(CaixaAte.Text) <> "" Then
        sCaixaF = Codigo_Extrai(CaixaAte.Text)
    Else
        sCaixaF = ""
    End If
        
    'Caixa Inicial não pode ser maiso do que a Final
    If sCaixaI <> "" And sCaixaF <> "" Then
    
        If StrParaInt(sCaixaI) > StrParaInt(sCaixaF) Then gError 116196
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 116196
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INICIAL_MAIOR", gErr)
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172449)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCaixaI As String, sCaixaF As String, sStatus As String) As Long
'monta a expressão de seleção de relatório

Dim lErro As Long
Dim sExpressao As String

On Error GoTo Erro_Monta_Expressao_Selecao
    
    'monta expressão da Caixa
    If sCaixaI <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CAIXA >= " & Forprint_ConvInt(StrParaInt(sCaixaI))
        
    End If
    
    If sCaixaF <> "" Then
         
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CAIXA <= " & Forprint_ConvInt(StrParaInt(sCaixaF))
        
    End If
            
    'monta expressão do Status
    If sStatus = CStr(CAIXA_STATUS_ABERTO) Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "STATUS = " & Forprint_ConvInt(CAIXA_STATUS_ABERTO)
    
    ElseIf sStatus = CStr(CAIXA_STATUS_FECHADO) Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "STATUS = " & Forprint_ConvInt(CAIXA_STATUS_FECHADO)
    
    End If
            
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressão o valor de Filial Empresa
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
    
    End If
    
    'passa a expressão completa para o obj
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172450)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCaixaI As String
Dim sCaixaF As String
Dim sStatus As String

On Error GoTo Erro_PreencherRelOp
       
    'verifica se os dados são validos
    lErro = Formata_E_Critica_Parametros(sCaixaI, sCaixaF)
    If lErro <> SUCESSO Then gError 116197

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116198
         
    'inclui parametro da Caixa
    lErro = objRelOpcoes.IncluirParametro("NCAIXAINIC", sCaixaI)
    If lErro <> AD_BOOL_TRUE Then gError 116199

    lErro = objRelOpcoes.IncluirParametro("NCAIXAFIM", sCaixaF)
    If lErro <> AD_BOOL_TRUE Then gError 116200
      
    'inclui parametro da Caixa (controle)
    lErro = objRelOpcoes.IncluirParametro("TCAIXAINIC", CaixaDe)
    If lErro <> AD_BOOL_TRUE Then gError 116226

    lErro = objRelOpcoes.IncluirParametro("TCAIXAFIM", CaixaAte)
    If lErro <> AD_BOOL_TRUE Then gError 116227
      
    'inclui parametro do status
    If StatusAberto.Value = True Then
        sStatus = CAIXA_STATUS_ABERTO
    ElseIf StatusFechado.Value = True Then
        sStatus = CAIXA_STATUS_FECHADO
    ElseIf StatusTodos.Value = True Then
        sStatus = 2
    End If
      
    'se o Status for <> de Todos, inclui o parametro
    If sStatus <> "" Then
        lErro = objRelOpcoes.IncluirParametro("NSTATUS", sStatus)
    End If
    
    'monta a expressão final
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCaixaI, sCaixaF, sStatus)
    If lErro <> SUCESSO Then gError 116201

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 116197 To 116201, 116226, 116227
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172451)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 116202

   'pega parâmetro Caixa Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAINIC", sParam)
    If lErro <> SUCESSO Then gError 116203
    
    CaixaDe.Text = sParam
    Call CaixaDe_Validate(bSGECancelDummy)
    
    'pega parâmetro Caixa Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAFIM", sParam)
    If lErro <> SUCESSO Then gError 116204
    
    CaixaAte.Text = sParam
    Call CaixaAte_Validate(bSGECancelDummy)
             
    'pega o parametro Status e exibe
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then gError 116205
             
    If sParam = CStr(CAIXA_STATUS_ABERTO) Then
        StatusAberto.Value = True
    ElseIf sParam = CStr(CAIXA_STATUS_FECHADO) Then
        StatusFechado.Value = True
    ElseIf sParam = "" Then
        StatusTodos.Value = True
    End If
             
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116202 To 116205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172452)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()
'envia p/ o arquivo a opção desejada

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'preenche as opções de relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116206

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 116206

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172453)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 116207
    
    Set gobjRelOpcoes = objRelOpcoes
    Set gobjRelatorio = objRelatorio
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116208
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116207, 116208
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172454)

    End Select

    Exit Function

End Function

Private Sub CaixaDe_Validate(Cancel As Boolean)
'valida a o cód/ nome do caixa

Dim lErro As Long
Dim objCaixa As ClassCaixa

On Error GoTo Erro_CaixaDe_Validate

    giCaixaInicial = 1

    If Len(Trim(CaixaDe.Text)) > 0 Then
        
        'instancia o obj
        Set objCaixa = New ClassCaixa
        
        'preenche o obj c/ o cod e filial
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler Caixa (Código ou nome)
        lErro = CF("TP_Caixa_Le1", CaixaDe, objCaixa)
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116209

        'código inexistente
        If lErro = 116175 Then gError 116210

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116211

    End If
    
    Exit Sub

Erro_CaixaDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116209

        Case 116210, 116211
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INEXISTENTE", gErr, CaixaDe.Text)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172455)

    End Select

    Exit Sub

End Sub

Private Sub CaixaAte_Validate(Cancel As Boolean)
'valida a o cód/ nome do caixa

Dim lErro As Long
Dim objCaixa As ClassCaixa

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
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116212

        'código inexistente
        If lErro = 116175 Then gError 116213

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116214

    End If
 
    Exit Sub

Erro_CaixaAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116212
            
        Case 116213, 116214
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INEXISTENTE", gErr, CaixaAte.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172456)

    End Select

    Exit Sub

End Sub

Private Sub LabelCaixaDe_Click()
'sub chamadora do browser

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaDe_Click

    giCaixaInicial = 1
    
    If Len(Trim(CaixaDe.Text)) > 0 Then
        'Preenche com a caixa  da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
    End If
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa)
    
    Else
    
        'Chama Tela de caixa
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)
    
    End If
    
    Exit Sub

Erro_LabelCaixaDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172457)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCaixa_evSelecao(obj1 As Object)
'evento de inclusao de item selecionado no browser caixa

Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixa_evSelecao

    Set objCaixa = obj1
    
    'Preenche campo Caixa
    If giCaixaInicial = 1 Then
        CaixaDe.Text = objCaixa.iCodigo
        CaixaDe_Validate (bSGECancelDummy)
    Else
        CaixaAte.Text = objCaixa.iCodigo
        CaixaAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoCaixa_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 172458)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelCaixaAte_Click()
'sub chamadora do browser caixa

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaAte_Click

    giCaixaInicial = 0
    
    If Len(Trim(CaixaAte.Text)) > 0 Then
        'Preenche com a caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
    End If
    
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'Chama Tela CaixaLista
        Call Chama_Tela("CaixaTodosLista", colSelecao, objCaixa, objEventoCaixa)
    
    Else
    
        'Chama Tela Caixa
        Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)
    
    End If
    
    Exit Sub

Erro_LabelCaixaAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172459)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()
'exclui a opção de relatorio selecionada

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116215

    'pergunta se deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO", ComboOpcoes.Text)

    'se sim
    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116216

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa a tela
        Call BotaoLimpar_Click
                
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116215
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116216

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172460)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Trim(ComboOpcoes.Text) = "" Then gError 116217

    'preenche o arquivo C c/ a opção de relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116218

    'carrega o obj com a opção da tela
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'grava a opção
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116219

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116220
    
    'limpa a tela
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116217
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116218, 116219, 116220

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172461)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'limpa a tela
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa o relatorio
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116221
    
    'posiciona o cursor e limpa a combo opcoes
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    StatusTodos.Value = True
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116221
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172462)

    End Select

    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click
        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click
        End If
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_RELOP_PEDIDOS_NAO_ENTREGUES
    Set Form_Load_Ocx = Me
    Caption = "Relação de Caixas"   '???
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelCaixas"
    
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

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelCaixaDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCaixaDe, Source, X, Y)
End Sub

Private Sub LabelCaixaDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCaixaDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCaixaAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCaixaAte, Source, X, Y)
End Sub

Private Sub LabelCaixaAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCaixaAte, Button, Shift, X, Y)
End Sub
