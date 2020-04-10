VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpJurosRecOcx 
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   LockControls    =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   7980
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5640
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpJurosRecOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpJurosRecOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpJurosRecOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpJurosRecOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Baixa"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   735
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownBaixaDe 
         Height          =   315
         Left            =   2400
         TabIndex        =   17
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
         Left            =   1245
         TabIndex        =   1
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownBaixaAte 
         Height          =   315
         Left            =   4500
         TabIndex        =   18
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
         TabIndex        =   2
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yy"
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
         Height          =   195
         Left            =   2940
         TabIndex        =   20
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
         Height          =   195
         Left            =   870
         TabIndex        =   19
         Top             =   330
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpJurosRecOcx.ctx":0994
      Left            =   1275
      List            =   "RelOpJurosRecOcx.ctx":0996
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
      Left            =   5820
      Picture         =   "RelOpJurosRecOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   825
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digitação da Baixa"
      Height          =   705
      Left            =   120
      TabIndex        =   11
      Top             =   1575
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownDigitacaoBaixaDe 
         Height          =   315
         Left            =   2400
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DigitacaoBaixaDe 
         Height          =   285
         Left            =   1230
         TabIndex        =   3
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
      Begin MSComCtl2.UpDown UpDownDigitacaoBaixaAte 
         Height          =   315
         Left            =   4500
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DigitacaoBaixaAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   4
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
      Begin VB.Label Label4 
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
         Left            =   870
         TabIndex        =   15
         Top             =   330
         Width           =   315
      End
      Begin VB.Label Label5 
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
         Left            =   2940
         TabIndex        =   14
         Top             =   330
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
      Left            =   570
      TabIndex        =   21
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpJurosRecOcx"
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

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 47732
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 47736
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 47736
        
        Case 47732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169440)

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
    If lErro <> SUCESSO Then Error 47733
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 47785
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47733, 47785
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169441)

    End Select

    Exit Sub
   
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 47739
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 47739
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169442)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 47743

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47744

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47745
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47746
    
    Call BotaoLimpar_Click
           
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47744, 47745, 47746
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169443)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 47748

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 47749

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 47748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 47749

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169444)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47752

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 47752

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169445)

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

On Error GoTo Erro_PreencherRelOp
    
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then Error 47757

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 47758
         
    'Preenche o Cliente Inicial
    If BaixaDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", BaixaDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47759
    
    'Preenche o Cliente Final
    If BaixaAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", BaixaAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DBXFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47760
                    
    'Preenche o tipo do Cliente
    If DigitacaoBaixaDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGBXINIC", DigitacaoBaixaDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGBXINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47761
    
    'Preenche com a Opcao Tipocliente(TodosClientes ou um Cliente)
    If DigitacaoBaixaAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDIGBXFIM", DigitacaoBaixaAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDIGBXFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47762
    
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 47766

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 47757, 47758, 47759, 47760, 47761, 47762, 47763, 47764, 47765, 47766
        
        Case 47783, 47784, 47787
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169446)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente e Cobrador

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'data da Baixa inicial não pode ser maior que a Baixa final
    If Trim(BaixaDe.ClipText) <> "" And Trim(BaixaAte.ClipText) <> "" Then
    
         If CDate(BaixaDe.Text) > CDate(BaixaAte.Text) Then Error 47780
    
    End If
    
    'data da Digitacao da Baixa inicial não pode ser maior que a data da digitacao da Baixa final
    If Trim(DigitacaoBaixaDe.ClipText) <> "" And Trim(DigitacaoBaixaAte.ClipText) <> "" Then
    
         If CDate(DigitacaoBaixaDe.Text) > CDate(DigitacaoBaixaAte.Text) Then Error 48585
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
        Case 47780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_INICIAL_MAIOR", Err)
        
        Case 48585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DIGITACAO_BAIXA_INICIAL_MAIOR", Err)
              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169447)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
    
    If Trim(BaixaDe.ClipText) <> "" Then sExpressao = "Baixa >= " & Forprint_ConvData(CDate(BaixaDe.Text))
    
    If Trim(BaixaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Baixa <= " & Forprint_ConvData(CDate(BaixaAte.Text))

    End If
        
    If Trim(DigitacaoBaixaDe.ClipText) <> "" Then

        sExpressao = "Emissao >= " & Forprint_ConvData(CDate(DigitacaoBaixaDe.Text))

    End If
    
    If Trim(DigitacaoBaixaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(DigitacaoBaixaAte.Text))

    End If
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169448)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim sCobrador As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 47770
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXINIC", sParam)
    If lErro <> SUCESSO Then Error 47771
    
    Call DateParaMasked(BaixaDe, CDate(sParam))
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DBXFIM", sParam)
    If lErro <> SUCESSO Then Error 47781

    Call DateParaMasked(BaixaAte, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDIGBXINIC", sParam)
    If lErro <> SUCESSO Then Error 47782

    Call DateParaMasked(DigitacaoBaixaDe, CDate(sParam))
       
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDIGBXFIM", sParam)
    If lErro <> SUCESSO Then Error 47788

    Call DateParaMasked(DigitacaoBaixaAte, CDate(sParam))
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 47770, 47771, 47772, 47773, 47774, 47775, 47776
        
        Case 47777, 47781, 47782, 47788
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169449)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long
Dim DataAux As Date

On Error GoTo Erro_Define_Padrao

    'Seta como padrao o dia anterior
    DataAux = gdtDataAtual - 1
    
'    'Define Data de Referencia como data atual
'    BaixaDe.Text = Format(DataAux, "dd/mm/yy")
'
'    'Define Data de Referencia como data atual
'    BaixaAte.Text = Format(DataAux, "dd/mm/yy")
'
'    DigitacaoBaixaDe.Text = Format(DataAux, "dd/mm/yy")
'
'    DigitacaoBaixaAte.Text = Format(DataAux, "dd/mm/yy")
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169450)
    
    End Select
    
    Exit Function
    
End Function

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
        If lErro <> SUCESSO Then Error 47789

    End If

    Exit Sub

Erro_BaixaAte_Validate:

    Cancel = True


    Select Case Err

        Case 47789

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169451)

    End Select

    Exit Sub

End Sub

Private Sub BaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaixaDe_Validate

    If Len(BaixaDe.ClipText) > 0 Then

        lErro = Data_Critica(BaixaDe.Text)
        If lErro <> SUCESSO Then Error 47790

    End If

    Exit Sub

Erro_BaixaDe_Validate:

    Cancel = True


    Select Case Err

        Case 47790

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169452)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoBaixaAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DigitacaoBaixaAte)

End Sub

Private Sub DigitacaoBaixaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DigitacaoBaixaAte_Validate

    If Len(DigitacaoBaixaAte.ClipText) > 0 Then
        
        lErro = Data_Critica(DigitacaoBaixaAte.Text)
        If lErro <> SUCESSO Then Error 47789

    End If

    Exit Sub

Erro_DigitacaoBaixaAte_Validate:

    Cancel = True


    Select Case Err

        Case 47789

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169453)

    End Select

    Exit Sub

End Sub

Private Sub DigitacaoBaixaDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DigitacaoBaixaDe)

End Sub

Private Sub DigitacaoBaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DigitacaoBaixaDe_Validate

    If Len(DigitacaoBaixaDe.ClipText) > 0 Then

        lErro = Data_Critica(DigitacaoBaixaDe.Text)
        If lErro <> SUCESSO Then Error 47790

    End If

    Exit Sub

Erro_DigitacaoBaixaDe_Validate:

    Cancel = True


    Select Case Err

        Case 47790

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169454)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
 
 End Sub

Private Sub UpDownDigitacaoBaixaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDigitacaoBaixaDe_DownClick

    lErro = Data_Up_Down_Click(DigitacaoBaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47850

    Exit Sub

Erro_UpDownDigitacaoBaixaDe_DownClick:

    Select Case Err

        Case 47850
            DigitacaoBaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169455)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDigitacaoBaixaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDigitacaoBaixaDe_UpClick

    lErro = Data_Up_Down_Click(DigitacaoBaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47851

    Exit Sub

Erro_UpDownDigitacaoBaixaDe_UpClick:

    Select Case Err

        Case 47851
            DigitacaoBaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169456)

    End Select

    Exit Sub

End Sub
    
Private Sub UpDownBaixaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_DownClick

    lErro = Data_Up_Down_Click(BaixaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47852

    Exit Sub

Erro_UpDownBaixaDe_DownClick:

    Select Case Err

        Case 47852
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169457)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaDe_UpClick

    lErro = Data_Up_Down_Click(BaixaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47853

    Exit Sub

Erro_UpDownBaixaDe_UpClick:

    Select Case Err

        Case 47853
            BaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169458)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownBaixaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaAte_DownClick

    lErro = Data_Up_Down_Click(BaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47854

    Exit Sub

Erro_UpDownBaixaAte_DownClick:

    Select Case Err

        Case 47854
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169459)

    End Select

    Exit Sub

End Sub

Private Sub UpDownBaixaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaAte_UpClick

    lErro = Data_Up_Down_Click(BaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47855

    Exit Sub

Erro_UpDownBaixaAte_UpClick:

    Select Case Err

        Case 47855
            BaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169460)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDigitacaoBaixaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDigitacaoBaixaAte_DownClick

    lErro = Data_Up_Down_Click(DigitacaoBaixaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47850

    Exit Sub

Erro_UpDownDigitacaoBaixaAte_DownClick:

    Select Case Err

        Case 47850
            DigitacaoBaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169461)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDigitacaoBaixaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDigitacaoBaixaAte_UpClick

    lErro = Data_Up_Down_Click(DigitacaoBaixaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47851

    Exit Sub

Erro_UpDownDigitacaoBaixaAte_UpClick:

    Select Case Err

        Case 47851
            DigitacaoBaixaAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169462)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_JUROS_REC
    Set Form_Load_Ocx = Me
    Caption = "Juros Recebidos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpJurosRec"
    
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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

