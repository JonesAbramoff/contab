VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpMovCaixa 
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   6690
   Begin VB.Frame FrameOperador 
      Caption         =   "Operador"
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   4215
      Begin MSMask.MaskEdBox OperadorDe 
         Height          =   315
         Left            =   690
         TabIndex        =   16
         Top             =   285
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OperadorAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   18
         Top             =   285
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelOperadorDe 
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelOperadorAte 
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
         Left            =   2160
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   720
         TabIndex        =   9
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   12
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataAte 
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
         Left            =   2265
         TabIndex        =   11
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelDataDe 
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
         Left            =   300
         TabIndex        =   8
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4215
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   690
         TabIndex        =   4
         Top             =   285
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   2685
         TabIndex        =   6
         Top             =   285
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   19
         PromptChar      =   " "
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
         Left            =   2160
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   5
         Top             =   345
         Width           =   360
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   3
         Top             =   360
         Width           =   315
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
      Left            =   4740
      Picture         =   "RelOpMovCaixa.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   945
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpMovCaixa.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpMovCaixa.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpMovCaixa.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpMovCaixa.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMovCaixa.ctx":0A96
      Left            =   1080
      List            =   "RelOpMovCaixa.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   2670
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
      TabIndex        =   0
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMovCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'eventos dos browsers
Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1
Private WithEvents objEventoOperador As AdmEvento
Attribute objEventoOperador.VB_VarHelpID = -1


'variaveis de controle de browser
Dim giOperadorInicial As Integer
Dim giCaixaInicial As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    'instancia o obj
    Set objEventoCaixa = New AdmEvento
    Set objEventoOperador = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169990)

    End Select

    Exit Sub
    
End Sub

Private Function Formata_E_Critica_Parametros(sCaixaI As String, sCaixaF As String, sOperadorI As String, sOperadorF As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
         
    'atribui o codigo a caixa inicia e final
    If CaixaDe.Text <> "" Then
        sCaixaI = Codigo_Extrai(CaixaDe.Text)
    Else
        sCaixaI = ""
    End If
    
    If CaixaAte.Text <> "" Then
        sCaixaF = Codigo_Extrai(CaixaAte.Text)
    Else
        sCaixaF = ""
    End If
    
    'critica Caixa Inicial e Final
    If sCaixaI <> "" And sCaixaF <> "" Then
        If StrParaInt(sCaixaI) > StrParaInt(sCaixaF) Then gError 116108
    End If
     
    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 116109
    End If
    
    'atribui o cod. operador inicial e final
    If OperadorDe.Text <> "" Then
        sOperadorI = Codigo_Extrai(OperadorDe.Text)
    Else
        sOperadorI = ""
    End If
        
    If OperadorAte.Text <> "" Then
        sOperadorF = Codigo_Extrai(OperadorAte.Text)
    Else
        sOperadorF = ""
    End If
              
    'operador inicial não pode ser maior de que o final
    If sOperadorI <> "" And sOperadorF <> "" Then
        If StrParaInt(sOperadorI) > StrParaInt(sOperadorF) Then gError 116110
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 116108
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INICIAL_MAIOR", gErr)
            CaixaDe.SetFocus
        
         Case 116109
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
            
        Case 116110
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_INICIAL_MAIOR", gErr)
            OperadorDe.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169991)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCaixaI As String, sCaixaF As String, sOperadorI As String, sOperadorF As String) As Long
'monta a expressão de seleção de relatório

Dim lErro As Long
Dim sExpressao As String

On Error GoTo Erro_Monta_Expressao_Selecao

    'monta expressão da caixa
    If sCaixaI <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa >= " & Forprint_ConvInt(StrParaInt(sCaixaI))
        
    End If

    If sCaixaF <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa <= " & Forprint_ConvInt(StrParaInt(sCaixaF))
        
    End If
    
    'monta a expressão da data
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(StrParaDate(DataDe.Text))

    End If
    
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(StrParaDate(DataAte.Text))

    End If
        
    'monta a expressão do operador
    If sOperadorI <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Operador >= " & Forprint_ConvInt(StrParaInt(sOperadorI))
        
    End If

    If sOperadorF <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Operador <= " & Forprint_ConvInt(StrParaInt(sOperadorF))

    End If
    
    'faz o filtro da filial_empresa
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169992)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCaixaI As String
Dim sCaixaF As String
Dim sOperadorI As String
Dim sOperadorF As String

On Error GoTo Erro_PreencherRelOp
       
    'formata os campos
    lErro = Formata_E_Critica_Parametros(sCaixaI, sCaixaF, sOperadorI, sOperadorF)
    If lErro <> SUCESSO Then gError 116111

    'limpa o obj
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 116112
         
    'preenche o arquivo com o operador
    lErro = objRelOpcoes.IncluirParametro("NOPERADORI", sOperadorI)
    If lErro <> AD_BOOL_TRUE Then gError 116113

    '???Luiz: NOPERADORF ok ivan
    lErro = objRelOpcoes.IncluirParametro("NOPERADORF", sOperadorF)
    If lErro <> AD_BOOL_TRUE Then gError 116114
            
    'preenche o arquivo c/ o controle Operador
    lErro = objRelOpcoes.IncluirParametro("TOPERADORI", Trim(OperadorDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116224

    lErro = objRelOpcoes.IncluirParametro("TOPERADORF", Trim(OperadorAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116225
    
    'preenche o arquivo com o codigo da caixa
    lErro = objRelOpcoes.IncluirParametro("NCAIXAI", sCaixaI)
    If lErro <> AD_BOOL_TRUE Then gError 116115
    
    lErro = objRelOpcoes.IncluirParametro("NCAIXAF", sCaixaF)
    If lErro <> AD_BOOL_TRUE Then gError 116116
    
    'preenche o arquivo c/ o controle Caixa
    lErro = objRelOpcoes.IncluirParametro("TCAIXAI", Trim(CaixaDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116222

    lErro = objRelOpcoes.IncluirParametro("TCAIXAF", Trim(CaixaAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 116223
    
    'preenche o arquivo com a data
    If DataDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 116117

    If DataAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 116118
    
    'preenche o arquivo c/ a filial (se <> de EMPRESA_TODA)
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        lErro = objRelOpcoes.IncluirParametro("NFILIAL", giFilialEmpresa)
        If lErro <> AD_BOOL_TRUE Then gError 116251
      
    End If
     
    'monta a expressão final
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCaixaI, sCaixaF, sOperadorI, sOperadorF)
    If lErro <> SUCESSO Then gError 116119

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 116111 To 116119, 116222, 116223, 116224, 116225
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169993)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 116138

   'pega parâmetro Caixa Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAI", sParam)
    If lErro <> SUCESSO Then gError 116139
    
    CaixaDe.Text = sParam
    Call CaixaDe_Validate(bSGECancelDummy)
    
    'pega parâmetro Caixa Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCAIXAF", sParam)
    If lErro <> SUCESSO Then gError 116140
    
    CaixaAte.Text = sParam
    Call CaixaAte_Validate(bSGECancelDummy)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 116143

    Call DateParaMasked(DataDe, StrParaDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 116144

    Call DateParaMasked(DataAte, StrParaDate(sParam))
        
    'pega o parametro Operador inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NOPERADORI", sParam)
    If lErro <> SUCESSO Then gError 116141
    
    OperadorDe.Text = sParam
    Call OperadorDe_Validate(bSGECancelDummy)
    
    'pega o parametro Operador final e exibe
    lErro = objRelOpcoes.ObterParametro("NOPERADORF", sParam)
    If lErro <> SUCESSO Then gError 116142
    
    OperadorAte.Text = sParam
    Call OperadorAte_Validate(bSGECancelDummy)
         
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 116138 To 116144

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169994)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Unload(Cancel As Integer)
'libera os obj

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoOperador = Nothing
    Set objEventoCaixa = Nothing
    
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Trim(ComboOpcoes.Text) = "" Then gError 116145

    'preenche o arquivo c/ as opções
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 116146

    'carrega o obj com a opção da tela
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'grava a opção
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 116147

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 116148
    
    'limpa a tela
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116145
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 116146, 116147, 116148

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169995)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'exclui a opção de rel. selecionada

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 116149

    'pergunta se deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO", ComboOpcoes.Text)

    'se a resp. for sim
    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 116150

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa a tela
        Call BotaoLimpar_Click
                
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 116149
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 116150

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169996)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Private Sub objEventoOperador_evSelecao(obj1 As Object)
'evento de inclusão de item selecionado no browser Operador

Dim objOperador As ClassOperador

On Error GoTo Erro_objEventoOperador_evSelecao

    Set objOperador = obj1
    
    'Preenche campo Operador
    If giOperadorInicial = 1 Then
        OperadorDe.Text = objOperador.iCodigo
        OperadorDe_Validate (bSGECancelDummy)
    Else
        OperadorAte.Text = objOperador.iCodigo
        OperadorAte_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoOperador_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 169997)

    End Select
    
    Exit Sub

End Sub

Private Sub OperadorDe_Validate(Cancel As Boolean)
'faz a validação do cód/nome do operador

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_OperadorDe_Validate

    giOperadorInicial = 1

    If Len(Trim(OperadorDe.Text)) > 0 Then
        
        'Tenta ler Operador(Código ou nome)
        lErro = CF("TP_Operador_Le", OperadorDe, objOperador)
        If lErro <> SUCESSO And lErro <> 117117 And lErro <> 117119 Then gError 116151

        'cód. não cadastrado
        If lErro = 117117 Then gError 116167

        'nome não cadastrado
        If lErro = 117119 Then gError 116191

    End If
    
    Exit Sub

Erro_OperadorDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116151

        Case 116167
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_INEXISTENTE", gErr, objOperador.iCodigo)

        Case 116191
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NOMERED_INEXISTENTE", gErr, objOperador.sNome)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169998)

    End Select

    Exit Sub

End Sub

Private Sub OperadorAte_Validate(Cancel As Boolean)
'faz a validação do cód/nome do operador

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_OperadorAte_Validate

    giOperadorInicial = 0
    
    If Len(Trim(OperadorAte.Text)) > 0 Then
    
        'Tenta ler o Operador (Código ou nome)
        lErro = CF("TP_Operador_Le", OperadorAte, objOperador)
        If lErro <> SUCESSO And lErro <> 117117 And lErro <> 117119 Then gError 116152

        'cód inexistente
        If lErro = 117117 Then gError 116166

        'nome inexistente
        If lErro = 117119 Then gError 116190

    End If
    
    Exit Sub

Erro_OperadorAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116152
            
        Case 116166
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_INEXISTENTE", gErr, objOperador.iCodigo)
        
        Case 116190
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NOMERED_INEXISTENTE", gErr, objOperador.sNome)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169999)

    End Select

    Exit Sub

End Sub

Private Sub LabelCaixaDe_Click()
'sub chamadora do browse caixa

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaDe_Click

    giCaixaInicial = 1
    
    If Len(Trim(CaixaDe.Text)) > 0 Then
        'Preenche com a caixa  da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaDe.Text)
    End If
    
    'Chama Tela de caixa
    Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)
    
    Exit Sub
    
Erro_LabelCaixaDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170000)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCaixa_evSelecao(obj1 As Object)
'evento de inclusão de item selecionado no browse caixa

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170001)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelCaixaAte_Click()
'sub chamadora do browse caixa

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection

On Error GoTo Erro_LabelCaixaAte_Click

    giCaixaInicial = 0
    
    If Len(Trim(CaixaAte.Text)) > 0 Then
        'Preenche com a caixa da tela
        objCaixa.iCodigo = Codigo_Extrai(CaixaAte.Text)
    End If
    
    'Chama Tela Caixa
    Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)

    Exit Sub

Erro_LabelCaixaAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170002)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelOperadorDe_Click()
'sub chamadora do browse operador

Dim objOperador As New ClassOperador
Dim colSelecao As Collection

On Error GoTo Erro_LabelOperadorDe_Click

    giOperadorInicial = 1
    
    If Len(Trim(OperadorDe.Text)) > 0 Then
        'Preenche com o Operador da tela
        objOperador.iCodigo = Codigo_Extrai(OperadorDe.Text)
         
    End If
    
    'Chama Tela de Operador
    Call Chama_Tela("OperadorLista", colSelecao, objOperador, objEventoOperador)
    
    Exit Sub

Erro_LabelOperadorDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170003)

    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelOperadorAte_Click()
'sub chamadora do browse operador

Dim objOperador As New ClassOperador
Dim colSelecao As Collection

On Error GoTo Erro_LabelOperadorAte_Click

    giOperadorInicial = 0
    
    If Len(Trim(OperadorAte.Text)) > 0 Then
        'Preenche com o Operador da tela
        objOperador.iCodigo = Codigo_Extrai(OperadorAte.Text)
    
    End If
    
    'Chama Tela Operador
    Call Chama_Tela("OperadorLista", colSelecao, objOperador, objEventoOperador)

    Exit Sub

Erro_LabelOperadorAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170004)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()
'manda p/ o arquivo a opção de relatorio

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'preenche as opções de relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 116153

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 116153

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170005)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'verifica se é possivel diminuir a data
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116154

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 116154
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170006)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'verifica se é possivel aumentar a data
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116155

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 116155
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170007)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'verifica se é possivel diminuir a data
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116156

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 116156
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170008)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()
'aumenta a data
 
Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'verifica se é possivel aumenta a data
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116157

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 116157
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170009)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe)
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'valida a data
 
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    ' se a data estiver preenchida
    If Len(DataDe.ClipText) > 0 Then
    
        'verifica se ela é valida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 116158

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 116158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170010)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte)
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'valida a data

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'se a data estiver preenchida
    If Len(DataAte.ClipText) > 0 Then

        'verifica se ela é valida
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 116159

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 116159

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170011)

    End Select

    Exit Sub

End Sub

Private Sub CaixaDe_Validate(Cancel As Boolean)
'valida o cód ou nome reduzido do caixa

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
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116160

        'código inexistente
        If lErro = 116175 Then gError 116164

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116184

    End If
    
    Exit Sub
    
Erro_CaixaDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116160

        Case 116164
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INEXISTENTE", gErr, objCaixa.iCodigo)
            
        Case 116184
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NOMERED_INEXISTENTE", gErr, objCaixa.sNomeReduzido)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170012)

    End Select

    Exit Sub

End Sub

Private Sub CaixaAte_Validate(Cancel As Boolean)
'valida o cód ou nome reduzido do caixa

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
        If lErro <> SUCESSO And lErro <> 116175 And lErro <> 116177 Then gError 116161

        'código inexistente
        If lErro = 116175 Then gError 116165

        'nome_reduzido inexistente
        If lErro = 116177 Then gError 116185

    End If
 
    Exit Sub

Erro_CaixaAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116161
            
        Case 116165
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_INEXISTENTE", gErr, objCaixa.iCodigo)
            
        Case 116185
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NOMERED_INEXISTENTE", gErr, objCaixa.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170013)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
'instancia os gobjs e carrega a combo opções

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se gobjrelatorio <> nothing
    If Not (gobjRelatorio Is Nothing) Then gError 116192
    
    'instancia os objs globias
    Set gobjRelOpcoes = objRelOpcoes
    Set gobjRelatorio = objRelatorio
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 116193
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 116193, 116192
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170014)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'limpa a tela
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa o relatorio
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 116162
    
    'posiciona o cursor na combo opções
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 116162
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170015)

    End Select

    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click
        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click
        ElseIf Me.ActiveControl Is OperadorDe Then
            Call LabelOperadorDe_Click
        ElseIf Me.ActiveControl Is OperadorAte Then
            Call LabelOperadorAte_Click
        End If
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_RELOP_PEDIDOS_NAO_ENTREGUES
    Set Form_Load_Ocx = Me
    Caption = "Movimentos de Caixa"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpMovCaixa"
    
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

Private Sub LabelOperadorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOperadorAte, Source, X, Y)
End Sub

Private Sub LabelOperadorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOperadorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelOperadorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOperadorDe, Source, X, Y)
End Sub

Private Sub LabelOperadorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOperadorDe, Button, Shift, X, Y)
End Sub

Private Sub labelDataAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataAte, Source, X, Y)
End Sub

Private Sub LabelDataAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataDe, Source, X, Y)
End Sub

Private Sub LabelDataDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataDe, Button, Shift, X, Y)
End Sub

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
