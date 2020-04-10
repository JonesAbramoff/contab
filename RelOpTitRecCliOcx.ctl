VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpTitRecCli 
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   ScaleHeight     =   3705
   ScaleWidth      =   8205
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   825
      Left            =   465
      TabIndex        =   16
      Top             =   2520
      Width           =   5355
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   600
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   345
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
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Documento"
      Height          =   660
      Left            =   480
      TabIndex        =   15
      Top             =   1725
      Width           =   5355
      Begin VB.OptionButton OptionNF 
         Caption         =   "NF"
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
         Left            =   1320
         TabIndex        =   22
         Top             =   330
         Width           =   960
      End
      Begin VB.OptionButton OptionND 
         Caption         =   "ND"
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
         Left            =   3270
         TabIndex        =   21
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Baixa"
      Height          =   705
      Left            =   465
      TabIndex        =   8
      Top             =   885
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
      Left            =   5985
      Picture         =   "RelOpTitRecCliOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   975
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5820
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTitRecCliOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTitRecCliOcx.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTitRecCliOcx.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTitRecCliOcx.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTitRecCliOcx.ctx":0A96
      Left            =   2025
      List            =   "RelOpTitRecCliOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   330
      Width           =   2730
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
      Left            =   1230
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTitRecCli"
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

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 123090
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 123091
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 123090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 123091
            
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
    If lErro <> SUCESSO Then gError 123092
    
    ComboOpcoes.Text = ""
     
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 123093
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 123092, 123093
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
   
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate
    
    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 123094

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 123094
            
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
        If lErro <> SUCESSO Then gError 123095

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 123095
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
  
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 123096
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 123096
        
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
    If ComboOpcoes.Text = "" Then gError 123097

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 123098

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 123099
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 123100
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 123097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 123098, 123099, 123100
        
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
    If ComboOpcoes.ListIndex = -1 Then gError 123101

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 123102

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 123101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 123102

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 123103
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 123103

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sTipoDoc As String

On Error GoTo Erro_PreencherRelOp
    
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sTipoDoc)
    If lErro <> SUCESSO Then gError 123104

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 123105
         
    'Preenche o Cliente Inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 123106
    
    'Preenche o Cliente Final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 123108
                    
    'Preenche Baixa Inicial
    If BaixaDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", BaixaDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 123110
    
    'Preenche Baixa Final
    If BaixaAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", BaixaAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 123111
    
    'Preenche o tipo de Cliente
    lErro = objRelOpcoes.IncluirParametro("TTIPODOC", sTipoDoc)
    If lErro <> AD_BOOL_TRUE Then gError 123126
    
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_I, sCliente_F, sTipoDoc)
    If lErro <> SUCESSO Then gError 123127

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 123104 To 123111
        
        Case 123126, 123127
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sTipoDoc As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o TipoCliente

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
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 123112
        
    End If
            
    'data inicial não pode ser maior que a data final
    If Trim(BaixaDe.ClipText) <> "" And Trim(BaixaAte.ClipText) <> "" Then
    
         If CDate(BaixaDe.Text) > CDate(BaixaAte.Text) Then gError 123113
    
    End If
    
    'Se a opção NF estiver selecionada
    If OptionNF.Value = True Then
        sTipoDoc = "NF"
    Else
        sTipoDoc = "ND"
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 123112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
                               
        Case 123113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_INICIAL_MAIOR", gErr)
            BaixaDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_I As String, sCliente_F As String, sTipoDoc As String) As Long
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
    'Coloca a data de Baixa Inicial e Final
    If Trim(BaixaDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Baixa >= " & Forprint_ConvData(CDate(BaixaDe.Text))

    End If
    
    If Trim(BaixaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Baixa <= " & Forprint_ConvData(CDate(BaixaAte.Text))

    End If
    
    'Coloca o tipo de documento
    If sExpressao <> "" Then sExpressao = sExpressao & " E "
    
    If sTipoDoc = "ND" Then
        sExpressao = sExpressao & "TDoc = " & Forprint_ConvTexto("ND")
    Else
        sExpressao = sExpressao & "TDoc <> " & Forprint_ConvTexto("ND")
    End If
    
    objRelOpcoes.sSelecao = sExpressao

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
    If lErro <> SUCESSO Then gError 123114
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 123115
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 123116
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXINIC", sParam)
    If lErro <> SUCESSO Then gError 123117

    Call DateParaMasked(BaixaDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DBXFIM", sParam)
    If lErro <> SUCESSO Then gError 123118

    Call DateParaMasked(BaixaAte, CDate(sParam))
    
    'exibe o tipo do documento
    lErro = objRelOpcoes.ObterParametro("TTIPODOC", sParam)
    If lErro <> SUCESSO Then gError 123119
        
    If sParam = "NF" Then
        OptionNF.Value = True
    Else
        OptionND.Value = True
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 123114 To 123119
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    BaixaDe.Text = Format(gdtDataAtual - 1, "dd/mm/yy")
    BaixaAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'defina o tipo
    OptionNF.Value = True
    
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
        If lErro <> SUCESSO Then gError 123120

    End If

    Exit Sub

Erro_BaixaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 123120

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
        If lErro <> SUCESSO Then gError 123121

    End If

    Exit Sub

Erro_BaixaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 123121

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
    
End Sub

Private Sub UpDownBaixaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_DownClick

    lErro = Data_Up_Down_Click(BaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 123122
    
    Exit Sub

Erro_UpDownBaixaDe_DownClick:

    Select Case Err

        Case 123122
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
    If lErro <> SUCESSO Then gError 123123

    Exit Sub

Erro_UpDownBaixaDe_UpClick:

    Select Case gErr

        Case 123123
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
    If lErro <> SUCESSO Then gError 123124

    Exit Sub

Erro_UpDownBaixaAte_DownClick:

    Select Case gErr

        Case 123124
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
    If lErro <> SUCESSO Then gError 123125

    Exit Sub

Erro_UpDownBaixaAte_UpClick:

    Select Case Err

        Case 123125
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
    Caption = "Títulos a Receber Liquidados por Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitRecCli"
    
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

