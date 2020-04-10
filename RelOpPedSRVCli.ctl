VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPedSRVCli 
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   5670
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3405
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPedSRVCli.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPedSRVCli.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPedSRVCli.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPedSRVCli.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   1185
      Left            =   210
      TabIndex        =   21
      Top             =   780
      Width           =   3105
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   630
         TabIndex        =   1
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   1605
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   705
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   630
         TabIndex        =   3
         Top             =   720
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   315
         Width           =   345
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   225
         TabIndex        =   22
         Top             =   780
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   660
      Left            =   210
      TabIndex        =   18
      Top             =   2805
      Width           =   5310
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   615
         TabIndex        =   7
         Top             =   240
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   8
         Top             =   240
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   300
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   285
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendedores"
      Height          =   720
      Left            =   210
      TabIndex        =   15
      Top             =   2025
      Width           =   5310
      Begin MSMask.MaskEdBox VendedorInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   270
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   6
         Top             =   285
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorAte 
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
         TabIndex        =   17
         Top             =   330
         Width           =   360
      End
      Begin VB.Label LabelVendedorDe 
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   315
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
      Left            =   3690
      Picture         =   "RelOpPedSRVCli.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   735
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPedSRVCli.ctx":0A96
      Left            =   810
      List            =   "RelOpPedSRVCli.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   330
      Width           =   2475
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
      Index           =   0
      Left            =   105
      TabIndex        =   24
      Top             =   375
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPedSRVCli"
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

Dim giClienteInicial As Integer
Dim giVendedorInicial As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
        
    Call Define_Padrao
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171071)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    giClienteInicial = 1
    giVendedorInicial = 1
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171072)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 37872
    
    'pega vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDINIC", sParam)
    If lErro Then gError 37873
    
    VendedorInicial.Text = sParam
    Call VendedorInicial_Validate(bSGECancelDummy)
    
    'pega  vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDFIM", sParam)
    If lErro Then gError 37874
    
    VendedorFinal.Text = sParam
    Call VendedorFinal_Validate(bSGECancelDummy)
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro Then gError 37875
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro Then gError 37876
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 37877

    Call DateParaMasked(DataInicial, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 37878

    Call DateParaMasked(DataFinal, CDate(sParam))
             
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 37872 To 37878

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171073)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 37868
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 37868
        
        Case 29884
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171074)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoCliente = Nothing
    Set objEventoVendedor = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 43184
    
    ComboOpcoes.Text = ""
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 43229
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 43184, 43229
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171075)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sVend_I As String
Dim sVend_F As String

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sVend_I, sVend_F)
    If lErro <> SUCESSO Then gError 37892
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 37893
    
    lErro = objRelOpcoes.IncluirParametro("NVENDINIC", sVend_I)
    If lErro <> AD_BOOL_TRUE Then gError 37894
    
    lErro = objRelOpcoes.IncluirParametro("TVENDINIC", VendedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54847

    lErro = objRelOpcoes.IncluirParametro("NVENDFIM", sVend_F)
    If lErro <> AD_BOOL_TRUE Then gError 37895
    
    lErro = objRelOpcoes.IncluirParametro("TVENDFIM", VendedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54848
         
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 37896
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54849

    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 37897
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54850
    
    lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(StrParaDate(DataInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 37898

    lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(StrParaDate(DataFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 37899
        
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sVend_I, sVend_F, sCliente_I, sCliente_F)
    If lErro <> SUCESSO Then gError 37900
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 37892 To 37900

        Case 54847, 54848, 54849, 54850
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171076)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 37901

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 37902

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 43185
    
        ComboOpcoes.Text = ""
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then gError 43230
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 37901
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 37902, 43185, 43230

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171077)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 37903

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 37903

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171078)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 37904

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 37905

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 37906

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 43186
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 37904
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 37905, 37906, 43186

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171079)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sVend_I As String, sVend_F As String, sCliente_I As String, sCliente_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

   If sVend_I <> "" Then sExpressao = "Vendedor >= " & Forprint_ConvInt(CInt(sVend_I))

   If sVend_F <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(CInt(sVend_F))
    End If
   
   If sCliente_I <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(CLng(sCliente_I))
   End If

   If sCliente_F <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_F))
    End If
     
    If Trim(DataInicial.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))
    End If
    
    If Trim(DataFinal.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171080)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sVend_I As String, sVend_F As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'critica vendedor Inicial e Final
    If VendedorInicial.Text <> "" Then
        sVend_I = CStr(Codigo_Extrai(VendedorInicial.Text))
    Else
        sVend_I = ""
    End If
    
    If VendedorFinal.Text <> "" Then
        sVend_F = CStr(Codigo_Extrai(VendedorFinal.Text))
    Else
        sVend_F = ""
    End If
    
    If sVend_I <> "" And sVend_F <> "" Then
        If CInt(sVend_I) > CInt(sVend_F) Then gError 37907
    End If
   
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
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 37908
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 37909
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 37907
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
            VendedorInicial.SetFocus
        
        Case 37908
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
        
         Case 37909
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171081)

    End Select

    Exit Function

End Function

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 37911

    End If
    
    giClienteInicial = 1
    
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37911
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171082)

    End Select

End Sub


Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 37912

    End If
    
    giClienteInicial = 0
 
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37912
             Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171083)

    End Select

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 37913

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37913

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171084)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 37914

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37914

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171085)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelVendedorAte_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 0
    
    If Len(Trim(VendedorFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorFinal.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub LabelVendedorDe_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 1
    
    If Len(Trim(VendedorInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorInicial.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    If giClienteInicial = 1 Then
        ClienteInicial.Text = CStr(objCliente.lCodigo)
        Call ClienteInicial_Validate(bSGECancelDummy)
    Else
        ClienteFinal.Text = CStr(objCliente.lCodigo)
        Call ClienteFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorInicial = 1 Then
        VendedorInicial.Text = CStr(objVendedor.iCodigo)
        Call VendedorInicial_Validate(bSGECancelDummy)
    Else
        VendedorFinal.Text = CStr(objVendedor.iCodigo)
        Call VendedorFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37915

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 37915
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171086)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37916

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 37916
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171087)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 37917

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 37917
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171088)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37918

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 37918
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171089)

    End Select

    Exit Sub

End Sub

Private Sub VendedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorInicial_Validate

    If Len(Trim(VendedorInicial.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorInicial, objVendedor, 0)
        If lErro <> SUCESSO Then gError 37924

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 37924
            'Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171090)

    End Select

End Sub

Private Sub VendedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorFinal_Validate

    If Len(Trim(VendedorFinal.Text)) > 0 Then

        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
        If lErro <> SUCESSO Then gError 37925

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 37925
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171091)

    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PEDIDO_VENDEDOR_CLIENTE
    Set Form_Load_Ocx = Me
    Caption = "Pedidos de Serviço por cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPedSRVCli"
    
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
        
        If Me.ActiveControl Is VendedorInicial Then
            Call LabelVendedorDe_Click
        ElseIf Me.ActiveControl Is VendedorFinal Then
            Call LabelVendedorAte_Click
        ElseIf Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
             Call LabelClienteAte_Click
        End If
    
    End If

End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorAte, Source, X, Y)
End Sub

Private Sub LabelVendedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
End Sub

Private Sub LabelVendedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorDe, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

