VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpBaixasCR 
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ScaleHeight     =   6795
   ScaleWidth      =   6510
   Begin VB.Frame Frame5 
      Caption         =   "Região de Venda"
      Height          =   1755
      Left            =   240
      TabIndex        =   42
      Top             =   4650
      Width           =   5955
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   525
         Left            =   4350
         Picture         =   "RelOpBaixasCRPur.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   900
         Width           =   1530
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   525
         Left            =   4350
         Picture         =   "RelOpBaixasCRPur.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   255
         Width           =   1530
      End
      Begin VB.ListBox ListRegioes 
         Height          =   1410
         Left            =   75
         Style           =   1  'Checkbox
         TabIndex        =   30
         Top             =   240
         Width           =   4290
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Vendedor"
      Height          =   555
      Left            =   240
      TabIndex        =   40
      Top             =   4050
      Width           =   5955
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   180
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label VendedorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   41
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de Documento"
      Height          =   765
      Left            =   225
      TabIndex        =   39
      Top             =   3255
      Width           =   5955
      Begin VB.OptionButton TipoDocApenas 
         Caption         =   "Apenas:"
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
         Left            =   330
         TabIndex        =   27
         Top             =   465
         Width           =   1050
      End
      Begin VB.OptionButton TipoDocTodos 
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
         Left            =   330
         TabIndex        =   26
         Top             =   225
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.ComboBox TipoDocSeleciona 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "RelOpBaixasCRPur.ctx":21FC
         Left            =   1380
         List            =   "RelOpBaixasCRPur.ctx":21FE
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   375
         Width           =   4395
      End
   End
   Begin VB.CheckBox CheckAnalitico 
      Caption         =   "Exibe baixas que não geraram movimento de conta corrente"
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
      Left            =   240
      TabIndex        =   33
      Top             =   6480
      Width           =   5580
   End
   Begin VB.Frame Frame3 
      Caption         =   "Conta Corrente"
      Height          =   840
      Left            =   225
      TabIndex        =   22
      Top             =   2400
      Width           =   5955
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   1335
         TabIndex        =   25
         Top             =   450
         Width           =   4395
      End
      Begin VB.OptionButton ApenasCta 
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
         Left            =   285
         TabIndex        =   24
         Top             =   510
         Width           =   1050
      End
      Begin VB.OptionButton TodasCtas 
         Caption         =   "Todas"
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
         Left            =   285
         TabIndex        =   23
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Digitação"
      Height          =   705
      Left            =   210
      TabIndex        =   16
      Top             =   645
      Width           =   3990
      Begin MSComCtl2.UpDown DigitacaoDeUpDown 
         Height          =   315
         Left            =   1695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DigitacaoDe 
         Height          =   315
         Left            =   555
         TabIndex        =   3
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown DigitacaoAteUpDown 
         Height          =   315
         Left            =   3600
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DigitacaoAte 
         Height          =   315
         Left            =   2445
         TabIndex        =   4
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
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
         Left            =   1995
         TabIndex        =   20
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Baixa"
      Height          =   990
      Left            =   210
      TabIndex        =   11
      Top             =   1395
      Width           =   5985
      Begin VB.ComboBox ComboDias 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "RelOpBaixasCRPur.ctx":2200
         Left            =   1305
         List            =   "RelOpBaixasCRPur.ctx":220A
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   570
         Width           =   1365
      End
      Begin VB.OptionButton FaixaDatas 
         Caption         =   "Faixa de Datas"
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
         Left            =   210
         TabIndex        =   35
         Top             =   225
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.OptionButton ApenasDias 
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
         Left            =   225
         TabIndex        =   34
         Top             =   615
         Width           =   1050
      End
      Begin MSComCtl2.UpDown BaixaDeUpDown 
         Height          =   315
         Left            =   3585
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox BaixaDe 
         Height          =   315
         Left            =   2445
         TabIndex        =   1
         Top             =   180
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown BaixaAteUpDown 
         Height          =   315
         Left            =   5505
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   165
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox BaixaAte 
         Height          =   315
         Left            =   4350
         TabIndex        =   2
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumDias 
         Height          =   315
         Left            =   2790
         TabIndex        =   37
         Top             =   570
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "dia(s)"
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
         Left            =   3480
         TabIndex        =   38
         Top             =   615
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Left            =   3900
         TabIndex        =   15
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   225
         Width           =   345
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpBaixasCRPur.ctx":2225
      Left            =   1110
      List            =   "RelOpBaixasCRPur.ctx":2227
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
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
      Left            =   4395
      Picture         =   "RelOpBaixasCRPur.ctx":2229
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4350
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         HelpContextID   =   1000
         Left            =   1605
         Picture         =   "RelOpBaixasCRPur.ctx":232B
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpBaixasCRPur.ctx":24A9
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpBaixasCRPur.ctx":29DB
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpBaixasCRPur.ctx":2B65
         Style           =   1  'Graphical
         TabIndex        =   7
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   420
      TabIndex        =   21
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpBaixasCR"
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

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 59587
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 59588
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 59588
        
        Case 59587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167225)

    End Select

    Exit Function

End Function

Private Sub BaixaAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(BaixaAte)

End Sub

Private Sub BaixaDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(BaixaDe)

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 59589
    
    ComboOpcoes.Text = ""
    
    Call Define_Padrao
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 59589
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167226)

    End Select

    Exit Sub
   
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoVendedor = New AdmEvento
    
    lErro = PreencheComboContas()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Carrega_TipoDocumento(TipoDocSeleciona)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CarregaList_Regioes
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167227)

    End Select

    Exit Sub

End Sub

Function PreencheComboContas() As Long

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As New AdmCodigoNome

On Error GoTo Erro_PreencheComboContas

    'Carrega a Coleção de Contas
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 59821

    'Preenche a ComboBox CodConta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    PreencheComboContas = SUCESSO

    Exit Function
    
Erro_PreencheComboContas:

    PreencheComboContas = Err

    Select Case Err

        Case 59821
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167228)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 59592

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 59593

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 59594
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 59595
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 59592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 59593, 59594, 59595
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167229)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 59596

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 59597

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 59596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 59597

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167230)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOpExec(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 59598

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 59598

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167231)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sCheckTipo As String
Dim sClienteTipo As String
Dim sCheckCobrador As String
Dim sCobrador As String
Dim sCheckContas As String
Dim sConta As String
Dim sTipoDoc As String, iIndice As Integer
Dim iVendedor As Integer, iNRegiao As Integer
Dim sRegiao As String, sListCount As String

On Error GoTo Erro_PreencherRelOp
    
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sTipoDoc, sCheckContas, sConta, iVendedor)
    If lErro <> SUCESSO Then Error 59599

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 59600
             
    lErro = Trata_DatasBaixa(objRelOpcoes)
    If lErro <> SUCESSO Then Error 59601
             
    'Preenche a data da digitacao inicial
    If Trim(DigitacaoDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGINIC", DigitacaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 59603
    
    'Preenche data da digitacao final
    If Trim(DigitacaoAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGFIM", DigitacaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 59604

    lErro = objRelOpcoes.IncluirParametro("TTIPODOC", sTipoDoc)
    If lErro <> AD_BOOL_TRUE Then Error 59831

    'Preenche a Conta Corrente
    lErro = objRelOpcoes.IncluirParametro("TCONTA", sConta)
    If lErro <> AD_BOOL_TRUE Then Error 59831
    
    lErro = objRelOpcoes.IncluirParametro("TCONTACORRENTE", ContaCorrente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59832

    'Preenche com a Opcao Conta Corrente(Todas Contas ou uma Cnta)
    lErro = objRelOpcoes.IncluirParametro("TTODCONTAS", sCheckContas)
    If lErro <> AD_BOOL_TRUE Then Error 59833

    'Preenche com o Exibir Devolução / Crédito
    lErro = objRelOpcoes.IncluirParametro("NEXIBTIT", CStr(CheckAnalitico.Value))
    If lErro <> AD_BOOL_TRUE Then Error 47822

    'Preenche com o Exibir Titulo a Titulo
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(iVendedor))
    If lErro <> AD_BOOL_TRUE Then gError 47822
    
    iNRegiao = 1
    'Percorre toda a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            sRegiao = Codigo_Extrai(ListRegioes.List(iIndice))
            'Inclui todas as Regioes que foram slecionados
            lErro = objRelOpcoes.IncluirParametro("NLIST" & SEPARADOR & iNRegiao, sRegiao)
            If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
            iNRegiao = iNRegiao + 1
        End If
    Next
    sListCount = iNRegiao - 1
    
    'Inclui o numero de Clientes selecionados na Lista
    lErro = objRelOpcoes.IncluirParametro("NLISTCOUNT", sListCount)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCheckContas, sConta, sTipoDoc)
    If lErro <> SUCESSO Then Error 59605

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 59599, 59600, 59601, 59603, 59604, 59605, 59831, 59832, 59833, 47822
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167232)

    End Select

    Exit Function

End Function

Function Formata_E_Critica_Parametros(sTipoDoc As String, sCheckContas As String, sConta As String, iVendedor As Integer) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim iIndice As Integer, iAchou As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros

    If TipoDocApenas.Value = True Then
        sTipoDoc = SCodigo_Extrai(TipoDocSeleciona.Text)
    Else
        sTipoDoc = ""
    End If
    
    If FaixaDatas.Value = True Then
    
        'Pelo menos um par De/Ate tem que estar Preenchido senão -----> Error
        If Trim(BaixaDe.ClipText) = "" Or Trim(BaixaAte.ClipText) = "" Then
            If Trim(DigitacaoDe.ClipText) = "" Or Trim(DigitacaoAte.ClipText) = "" Then gError 59606
        End If
        
        'data da Baixa inicial não pode ser maior que a Baixa final
        If Trim(BaixaDe.ClipText) <> "" And Trim(BaixaAte.ClipText) <> "" Then
        
             If CDate(BaixaDe.Text) > CDate(BaixaAte.Text) Then Error 59607
        
        End If
    
    Else
    
        If Len(NumDias.Text) = 0 Or ComboDias.ListIndex = -1 Then
            If Trim(DigitacaoDe.ClipText) = "" Or Trim(DigitacaoAte.ClipText) = "" Then gError 178896
        End If
    
    End If
    
    
    'data daDigitacao da Baixa inicial não pode ser maior que a data da digitacao da Baixa final
    If Trim(DigitacaoDe.ClipText) <> "" And Trim(DigitacaoAte.ClipText) <> "" Then
    
         If CDate(DigitacaoDe.Text) > CDate(DigitacaoAte.Text) Then gError 59608
    
    End If
    
    'Se a opção para todas as Contas estiver selecionada
    If TodasCtas.Value = True Then
        sCheckContas = "Todas"
        sConta = ""
    
    'Se a opção para apenas uma Conta estiver selecionada
    Else
        'Tem que indicar a Conta
        If ContaCorrente.Text = "" Then gError 59838
        sCheckContas = "Uma"
        sConta = ContaCorrente.Text
    
    End If

    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
    
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError ERRO_SEM_MENSAGEM

        iVendedor = objVendedor.iCodigo

    End If
    
    'Limpa a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            iAchou = 1
            Exit For
        End If
    Next
       
    If iAchou = 0 Then gError 207095
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 59606, 178896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UMA_DATA_NAO_PREENCHIDA", gErr)
            
        Case 59607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_INICIAL_MAIOR", gErr)
        
        Case 59608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DIGITACAO_BAIXA_INICIAL_MAIOR", gErr)
              
        Case 59838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", gErr)
            ContaCorrente.SetFocus
                            
        Case 207095
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_ROTA_SELECIONADA", gErr)
                           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167233)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCheckContas As String, sConta As String, sTipoDoc As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim sSub As String, iIndice As Integer
Dim iCount As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
    
    'Se a opção para apenas uma conta estiver selecionada
    If sCheckContas = "Uma" Then

        If CheckAnalitico.Value = 1 Then
            sExpressao = "(Conta = " & Forprint_ConvInt(Codigo_Extrai(sConta))
            sExpressao = sExpressao & " OU NumMovCta = 0)"
        Else
            sExpressao = "Conta = " & Forprint_ConvInt(Codigo_Extrai(sConta))
        End If
        
    End If
    
    lErro = Exp_Sel_DatasBaixa(sExpressao)
    If lErro <> SUCESSO Then gError 178894
        
    If Trim(DigitacaoDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao >= " & Forprint_ConvData(CDate(DigitacaoDe.Text))

    End If
    
    If Trim(DigitacaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(DigitacaoAte.Text))

    End If
        
    If giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> gobjCR.iFilialCentralizadora Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(CInt(giFilialEmpresa))
    End If
    
    If Trim(sTipoDoc) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoDoc <= " & Forprint_ConvTexto(sTipoDoc)

    End If
    
    sSub = ""
    iCount = 0
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) Then
            iCount = iCount + 1
            If sSub <> "" Then sSub = sSub & " OU "
            sSub = sSub & " Regiao = " & Forprint_ConvInt(ListRegioes.ItemData(iIndice))
        End If
    Next
    
    'Se selecionou só alguns
    If Len(Trim(sSub)) <> 0 And iCount <> ListRegioes.ListCount Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "(" & sSub & ")"

    End If
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case 178894

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167234)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim sCobrador As String
Dim sConta As String
Dim iIndice As Integer
Dim sListCount As String, iIndiceRel As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 59609
   
    lErro = Preenche_DatasBaixa(objRelOpcoes)
    If lErro <> SUCESSO Then Error 59610
   
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDIGINIC", sParam)
    If lErro <> SUCESSO Then Error 59612

    Call DateParaMasked(DigitacaoDe, CDate(sParam))
       
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDIGFIM", sParam)
    If lErro <> SUCESSO Then Error 59613

    Call DateParaMasked(DigitacaoAte, CDate(sParam))
    
    'pega conta e Exibe
    lErro = objRelOpcoes.ObterParametro("TTODCONTAS", sParam)
    If lErro <> SUCESSO Then Error 59841
                   
    If sParam = "Todas" Then
    
        Call TodasCtas_Click
    
    Else
        'se é apenas uma entao exibe esta
        lErro = objRelOpcoes.ObterParametro("TCONTA", sConta)
        If lErro <> SUCESSO Then Error 59842
                            
        ApenasCta.Value = True
        ContaCorrente.Enabled = True
        CheckAnalitico.Enabled = True
        
        If sConta = "" Then
            ContaCorrente.ListIndex = -1
        Else
            ContaCorrente.Text = sConta
        End If
    End If

    lErro = objRelOpcoes.ObterParametro("NEXIBTIT", sParam)
    If lErro <> SUCESSO Then Error 47835
       
    CheckAnalitico.Value = CInt(sParam)
    
    'pega o tipo de documento
    lErro = objRelOpcoes.ObterParametro("TTIPODOC", sParam)
    If lErro <> SUCESSO Then Error 47835

    If Len(Trim(sParam)) > 0 Then
        TipoDocTodos.Value = False
        TipoDocApenas.Value = True
        For iIndice = 0 To TipoDocSeleciona.ListCount - 1
            If SCodigo_Extrai(TipoDocSeleciona.List(iIndice)) = sParam Then
                TipoDocSeleciona.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoDocTodos.Value = True
        TipoDocApenas.Value = False
        TipoDocSeleciona.ListIndex = -1
    End If

    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError 47835
    
    If StrParaInt(sParam) <> 0 Then
        Vendedor.Text = sParam
        Call Vendedor_Validate(bSGECancelDummy)
    End If
    
    'Obtem o numero de Regioes selecionados na Lista
    lErro = objRelOpcoes.ObterParametro("NLISTCOUNT", sListCount)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    'Percorre toda a Lista
    
    For iIndice = 0 To ListRegioes.ListCount - 1
        
        'Percorre todas as Regieos que foram slecionados
        For iIndiceRel = 1 To StrParaInt(sListCount)
            lErro = objRelOpcoes.ObterParametro("NLIST" & SEPARADOR & iIndiceRel, sParam)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            'Se o cliente não foi excluido
            If sParam = Codigo_Extrai(ListRegioes.List(iIndice)) Then
                'Marca as Regioes que foram gravados
                ListRegioes.Selected(iIndice) = True
            End If
        Next
    Next
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 59609, 59610, 59611, 59612, 59613, 47835
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167235)

    End Select

    Exit Function

End Function

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se foi preenchida a ComboBox
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 59846

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro <> SUCESSO Then Error 59847

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 59846 'Tratado na rotina chamada
    
        Case 59847
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, ContaCorrente.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167236)

    End Select

    Exit Sub

End Sub



Private Sub NumDias_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumDias)

End Sub

Private Sub TodasCtas_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TodasCtas_Click
    
    'Limpa e desabilita a ComboTipo
    ContaCorrente.ListIndex = -1
    ContaCorrente.Text = ""
    ContaCorrente.Enabled = False
    TodasCtas.Value = True
    
    CheckAnalitico.Enabled = False
    
    Exit Sub

Erro_TodasCtas_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167237)

    End Select

    Exit Sub
    
End Sub

Function Define_Padrao() As Long
'preenche padroes (valores default) na tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    BaixaDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    BaixaAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    DigitacaoDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DigitacaoAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'defina todas as contas
    Call TodasCtas_Click
    
    TipoDocTodos.Value = True
    TipoDocApenas.Value = False
    TipoDocSeleciona.ListIndex = -1

    'define Exibir Devolução / Crédito como Padrao
    CheckAnalitico.Value = 1

    Call Limpa_ListRegioes

    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167238)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ApenasCta_Click()

Dim lErro As Long

On Error GoTo Erro_OptionUmTipo_Click
    
    'Limpa Combo Tipo e Abilita
    ContaCorrente.ListIndex = -1
    ContaCorrente.Enabled = True
    ContaCorrente.SetFocus
    
    CheckAnalitico.Enabled = True
    
    Exit Sub

Erro_OptionUmTipo_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167239)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub BaixaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaixaAte_Validate

    If Len(BaixaAte.ClipText) > 0 Then
        
        lErro = Data_Critica(BaixaAte.Text)
        If lErro <> SUCESSO Then Error 59614

    End If

    Exit Sub

Erro_BaixaAte_Validate:

    Cancel = True


    Select Case Err

        Case 59614

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167240)

    End Select

    Exit Sub

End Sub

Private Sub BaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaixaDe_Validate

    If Len(BaixaDe.ClipText) > 0 Then

        lErro = Data_Critica(BaixaDe.Text)
        If lErro <> SUCESSO Then Error 59615

    End If

    Exit Sub

Erro_BaixaDe_Validate:

    Cancel = True


    Select Case Err

        Case 59615

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167241)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DigitacaoAte)

End Sub

Private Sub DigitacaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DigitacaoAte_Validate

    If Len(DigitacaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(DigitacaoAte.Text)
        If lErro <> SUCESSO Then Error 59616

    End If

    Exit Sub

Erro_DigitacaoAte_Validate:

    Cancel = True


    Select Case Err

        Case 59616

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167242)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DigitacaoDe)

End Sub

Private Sub DigitacaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DigitacaoDe_Validate

    If Len(DigitacaoDe.ClipText) > 0 Then

        lErro = Data_Critica(DigitacaoDe.Text)
        If lErro <> SUCESSO Then Error 59617

    End If

    Exit Sub

Erro_DigitacaoDe_Validate:

    Cancel = True


    Select Case Err

        Case 59617

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167243)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoVendedor = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Private Sub DigitacaoDeUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoDeUpDown_DownClick

    lErro = Data_Up_Down_Click(DigitacaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 59618

    Exit Sub

Erro_DigitacaoDeUpDown_DownClick:

    Select Case Err

        Case 59618
            DigitacaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167244)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoDeUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoDeUpDown_UpClick

    lErro = Data_Up_Down_Click(DigitacaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 59619

    Exit Sub

Erro_DigitacaoDeUpDown_UpClick:

    Select Case Err

        Case 59619
            DigitacaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167245)

    End Select

    Exit Sub

End Sub
    
Private Sub BaixaDeUpDoWn_DownClick()

Dim lErro As Long

On Error GoTo Erro_BaixaDeUpDoWn_DownClick

    lErro = Data_Up_Down_Click(BaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 59620

    Exit Sub

Erro_BaixaDeUpDoWn_DownClick:

    Select Case Err

        Case 59620
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167246)

    End Select

    Exit Sub

End Sub

Private Sub BaixaDeUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_BaixaDeUpDown_UpClick

    lErro = Data_Up_Down_Click(BaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 59621

    Exit Sub

Erro_BaixaDeUpDown_UpClick:

    Select Case Err

        Case 59621
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167247)

    End Select

    Exit Sub
    
End Sub

Private Sub BaixaAteUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_BaixaAteUpDown_DownClick

    lErro = Data_Up_Down_Click(BaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 59622

    Exit Sub

Erro_BaixaAteUpDown_DownClick:

    Select Case Err

        Case 59622
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167248)

    End Select

    Exit Sub

End Sub

Private Sub BaixaAteUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_BaixaAteUpDown_UpClick

    lErro = Data_Up_Down_Click(BaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 59623
    
    Exit Sub

Erro_BaixaAteUpDown_UpClick:

    Select Case Err

        Case 59623
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167249)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoAteUpDown_DownClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoAteUpDown_DownClick

    lErro = Data_Up_Down_Click(DigitacaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 59624

    Exit Sub

Erro_DigitacaoAteUpDown_DownClick:

    Select Case Err

        Case 59624
            DigitacaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167250)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoAteUpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_DigitacaoAteUpDown_UpClick

    lErro = Data_Up_Down_Click(DigitacaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 59625

    Exit Sub

Erro_DigitacaoAteUpDown_UpClick:

    Select Case Err

        Case 59625
            DigitacaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167251)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BAIXASCR
    Set Form_Load_Ocx = Me
    Caption = "Relação de Baixas no Contas a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBaixasCR"
    
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




Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub ApenasDias_Click()
    BaixaDe.Enabled = False
    BaixaDeUpDown.Enabled = False
    BaixaAte.Enabled = False
    BaixaAteUpDown.Enabled = False
    NumDias.Enabled = True
    ComboDias.Enabled = True
End Sub

Private Sub FaixaDatas_Click()
    BaixaDe.Enabled = True
    BaixaDeUpDown.Enabled = True
    BaixaAte.Enabled = True
    BaixaAteUpDown.Enabled = True
    NumDias.Enabled = False
    ComboDias.Enabled = False
End Sub

Private Function Trata_DatasBaixa(objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim lNumDias As Long

On Error GoTo Erro_Trata_DatasBaixa

    If FaixaDatas.Value = True Then

        'Preenche data baixa inicial
        If Trim(BaixaDe.ClipText) <> "" Then
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", BaixaDe.Text)
        Else
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
        End If
        If lErro <> AD_BOOL_TRUE Then gError 178882
        
        'Preenche data da baixa Final
        If Trim(BaixaAte.ClipText) <> "" Then
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", BaixaAte.Text)
        Else
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
        End If
        If lErro <> AD_BOOL_TRUE Then gError 178883
    
    Else
    
        If ComboDias.ListIndex = -1 Then
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
            
        ElseIf ComboDias.ItemData(ComboDias.ListIndex) = REL_OPCOES_ULTIMOS_DIAS Then
        
            If Len(NumDias.Text) = 0 Then gError 178884
        
            lNumDias = StrParaLong(NumDias.Text) - 1
    
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", "DATA_HOJE()-" & CStr(lNumDias))
            If lErro <> AD_BOOL_TRUE Then gError 178885
    
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", "DATA_HOJE()")
            If lErro <> AD_BOOL_TRUE Then gError 178886
            
        Else
            
            If Len(NumDias.Text) = 0 Then gError 178887
        
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", "DATA_HOJE()+1")
            If lErro <> AD_BOOL_TRUE Then gError 178888
    
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", "DATA_HOJE()+" & NumDias.Text)
            If lErro <> AD_BOOL_TRUE Then gError 178889
    
        End If
    
    
    End If

    Trata_DatasBaixa = SUCESSO
    
    Exit Function

Erro_Trata_DatasBaixa:

    Trata_DatasBaixa = gErr

    Select Case gErr

        Case 178882, 178883, 178885, 178886, 178888, 178889
        
        Case 178884, 178887
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_DIAS_ZERADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178890)

    End Select

    Exit Function

End Function

Private Function Preenche_DatasBaixa(objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim lNumDias As Long
Dim iPos As Integer
Dim sParam As String
Dim sParam1 As String
Dim iIndice As Integer

On Error GoTo Erro_Preenche_DatasBaixa


    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXINIC", sParam)
    If lErro <> SUCESSO Then gError 178891
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXFIM", sParam1)
    If lErro <> SUCESSO Then gError 178892
    
    iPos = InStr(sParam, "DATA_HOJE()")
    
    If iPos = 0 Then
        FaixaDatas.Value = True
        Call DateParaMasked(BaixaDe, CDate(sParam))
        Call DateParaMasked(BaixaAte, CDate(sParam1))
    Else
        ApenasDias.Value = True
        If sParam1 = "DATA_HOJE()" Then
            For iIndice = 0 To ComboDias.ListCount - 1
                If ComboDias.ItemData(iIndice) = REL_OPCOES_ULTIMOS_DIAS Then
                    ComboDias.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            lNumDias = StrParaInt(Mid(sParam, 13)) + 1
            
            NumDias.Text = lNumDias

        Else

            For iIndice = 0 To ComboDias.ListCount - 1
                If ComboDias.ItemData(iIndice) = REL_OPCOES_PROXIMOS_DIAS Then
                    ComboDias.ListIndex = iIndice
                    Exit For
                End If
            Next

            NumDias.Text = Mid(sParam1, 13)

        End If

    End If

    Preenche_DatasBaixa = SUCESSO
    
    Exit Function

Erro_Preenche_DatasBaixa:

    Preenche_DatasBaixa = gErr

    Select Case gErr

        Case 178891, 178892
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178893)

    End Select

    Exit Function

End Function

Function PreencherRelOpExec(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sCheckTipo As String
Dim sClienteTipo As String
Dim sCheckCobrador As String
Dim sCobrador As String
Dim sCheckContas As String
Dim sConta As String
Dim sTipoDoc As String
Dim iVendedor As Integer

On Error GoTo Erro_PreencherRelOpExec
    
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sTipoDoc, sCheckContas, sConta, iVendedor)
    If lErro <> SUCESSO Then Error 59599

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 59600
             
    lErro = Trata_DatasBaixaExec(objRelOpcoes)
    If lErro <> SUCESSO Then Error 59601
             
    'Preenche a data da digitacao inicial
    If Trim(DigitacaoDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGINIC", DigitacaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 59603
    
    'Preenche data da digitacao final
    If Trim(DigitacaoAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGFIM", DigitacaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 59604

    lErro = objRelOpcoes.IncluirParametro("TTIPODOC", sTipoDoc)
    If lErro <> AD_BOOL_TRUE Then Error 59831

    'Preenche a Conta Corrente
    lErro = objRelOpcoes.IncluirParametro("TCONTA", sConta)
    If lErro <> AD_BOOL_TRUE Then Error 59831
    
    lErro = objRelOpcoes.IncluirParametro("TCONTACORRENTE", ContaCorrente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 59832

    'Preenche com a Opcao Conta Corrente(Todas Contas ou uma Cnta)
    lErro = objRelOpcoes.IncluirParametro("TTODCONTAS", sCheckContas)
    If lErro <> AD_BOOL_TRUE Then Error 59833

    'Preenche com o Exibir Devolução / Crédito
    lErro = objRelOpcoes.IncluirParametro("NEXIBTIT", CStr(CheckAnalitico.Value))
    If lErro <> AD_BOOL_TRUE Then Error 47822

    'Preenche com o Exibir Titulo a Titulo
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(iVendedor))
    If lErro <> AD_BOOL_TRUE Then gError 47822

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCheckContas, sConta, sTipoDoc)
    If lErro <> SUCESSO Then Error 59605

    PreencherRelOpExec = SUCESSO

    Exit Function

Erro_PreencherRelOpExec:

    PreencherRelOpExec = Err

    Select Case Err

        Case 59599, 59600, 59601, 59603, 59604, 59605, 59831, 59832, 59833, 47822
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167232)

    End Select

    Exit Function

End Function

Private Function Trata_DatasBaixaExec(objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim lNumDias As Long

On Error GoTo Erro_Trata_DatasBaixaExec

    If FaixaDatas.Value = True Then

        'Preenche data baixa inicial
        If Trim(BaixaDe.ClipText) <> "" Then
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", BaixaDe.Text)
        Else
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
        End If
        If lErro <> AD_BOOL_TRUE Then gError 178882
        
        'Preenche data da baixa Final
        If Trim(BaixaAte.ClipText) <> "" Then
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", BaixaAte.Text)
        Else
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
        End If
        If lErro <> AD_BOOL_TRUE Then gError 178883
    
    Else
    
        If ComboDias.ListIndex = -1 Then
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
            
        ElseIf ComboDias.ItemData(ComboDias.ListIndex) = REL_OPCOES_ULTIMOS_DIAS Then
        
            If Len(NumDias.Text) = 0 Then gError 178884
        
            lNumDias = StrParaLong(NumDias.Text) - 1
    
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", Format(DateAdd("d", -lNumDias, gdtDataHoje), "dd/mm/yy"))
            If lErro <> AD_BOOL_TRUE Then gError 178885
    
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", Format(gdtDataHoje, "dd/mm/yy"))
            If lErro <> AD_BOOL_TRUE Then gError 178886
            
        Else
            
            If Len(NumDias.Text) = 0 Then gError 178887
        
            lErro = objRelOpcoes.IncluirParametro("DBXINIC", Format(DateAdd("d", 1, gdtDataHoje), "dd/mm/yy"))
            If lErro <> AD_BOOL_TRUE Then gError 178888
    
            lErro = objRelOpcoes.IncluirParametro("DBXFIM", Format(DateAdd("d", lNumDias, gdtDataHoje), "dd/mm/yy"))
            If lErro <> AD_BOOL_TRUE Then gError 178889
    
        End If
    
    
    End If

    Trata_DatasBaixaExec = SUCESSO
    
    Exit Function

Erro_Trata_DatasBaixaExec:

    Trata_DatasBaixaExec = gErr

    Select Case gErr

        Case 178882, 178883, 178885, 178886, 178888, 178889
        
        Case 178884, 178887
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_DIAS_ZERADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178890)

    End Select

    Exit Function

End Function

Private Function Exp_Sel_DatasBaixa(sExpressao As String) As Long

Dim lNumDias As Long

On Error GoTo Erro_Exp_Sel_DatasBaixa

    If FaixaDatas.Value = True Then

        If Trim(BaixaDe.ClipText) <> "" Then
        
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Baixa >= " & Forprint_ConvData(CDate(BaixaDe.Text))
        
        End If
        
        If Trim(BaixaAte.ClipText) <> "" Then
    
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "Baixa <= " & Forprint_ConvData(CDate(BaixaAte.Text))
    
        End If

    Else

        If ComboDias.ListIndex > -1 Then


            If ComboDias.ItemData(ComboDias.ListIndex) = REL_OPCOES_ULTIMOS_DIAS Then
            
                lNumDias = StrParaLong(NumDias.Text) - 1
        
                If sExpressao <> "" Then sExpressao = sExpressao & " E "
        
                sExpressao = sExpressao & "Baixa >= " & Forprint_ConvData(DateAdd("d", -lNumDias, gdtDataHoje))
        
                If sExpressao <> "" Then sExpressao = sExpressao & " E "
        
                sExpressao = sExpressao & "Baixa <= " & Forprint_ConvData(gdtDataHoje)
                
            Else
                
            
                If sExpressao <> "" Then sExpressao = sExpressao & " E "
            
                sExpressao = sExpressao & "Baixa >= " & Forprint_ConvData(DateAdd("d", 1, gdtDataHoje))
        
                If sExpressao <> "" Then sExpressao = sExpressao & " E "
                
                sExpressao = sExpressao & "Baixa <= " & Forprint_ConvData(DateAdd("d", lNumDias, gdtDataHoje))
                
            End If

        End If

    End If

    Exp_Sel_DatasBaixa = SUCESSO
    
    Exit Function

Erro_Exp_Sel_DatasBaixa:

    Exp_Sel_DatasBaixa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178895)

    End Select

    Exit Function

End Function

Public Sub TipoDocApenas_Click()

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = True

End Sub

Public Sub TipoDocTodos_Click()

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona.ListIndex = COMBO_INDICE

End Sub

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 190523
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 190523

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190524)

    End Select

    Exit Function

End Function

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim dPercComissao As Double

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If Len(Trim(Vendedor.Text)) <> 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Vendedor, objVendedor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209207)
    
    End Select

End Sub

Public Sub VendedorLabel_Click()

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    Me.Show

    Vendedor.SetFocus 'Inserido por Wagner
    
    Exit Sub

End Sub


Function CarregaList_Regioes() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_CarregaList_Regioes
    
    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 207090

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        ListRegioes.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        ListRegioes.ItemData(ListRegioes.NewIndex) = objCodigoDescricao.iCodigo
    Next

    CarregaList_Regioes = SUCESSO

    Exit Function

Erro_CarregaList_Regioes:

    CarregaList_Regioes = gErr

    Select Case gErr

        Case 207900

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172566)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcar_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoDesmarcar_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Sub Limpa_ListRegioes()

Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Public Function RetiraNomes_Sel(colRegioes As Collection) As Long
'Retira da combo todos os nomes que não estão selecionados

Dim iIndice As Integer
Dim lCodRegiao As Long

    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            lCodRegiao = LCodigo_Extrai(ListRegioes.List(iIndice))
            colRegioes.Add lCodRegiao
        End If
    Next
    
End Function
