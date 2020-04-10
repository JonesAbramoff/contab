VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpExtratoConcOcx 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   ScaleHeight     =   3555
   ScaleWidth      =   6600
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   810
      Left            =   135
      TabIndex        =   11
      Top             =   2310
      Width           =   3990
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   315
         Left            =   1665
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoDe 
         Height          =   315
         Left            =   510
         TabIndex        =   13
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   315
         Left            =   3585
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoAte 
         Height          =   315
         Left            =   2445
         TabIndex        =   15
         Top             =   315
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
         Left            =   2040
         TabIndex        =   17
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label5 
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
         Left            =   150
         TabIndex        =   16
         Top             =   375
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Conta Corrente"
      Height          =   1155
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4005
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
         Left            =   195
         TabIndex        =   10
         Top             =   285
         Width           =   900
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
         Left            =   180
         TabIndex        =   9
         Top             =   712
         Width           =   1050
      End
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   1245
         TabIndex        =   8
         Top             =   675
         Width           =   2550
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
      Left            =   4425
      Picture         =   "RelOpExtratoConcOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   930
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpExtratoConcOcx.ctx":0102
      Left            =   945
      List            =   "RelOpExtratoConcOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   375
      Width           =   3060
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpExtratoConcOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpExtratoConcOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpExtratoConcOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpExtratoConcOcx.ctx":0940
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
      Height          =   255
      Left            =   225
      TabIndex        =   18
      Top             =   420
      Width           =   690
   End
End
Attribute VB_Name = "RelOpExtratoConcOcx"
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

    If Not (gobjRelatorio Is Nothing) Then Error 59739
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 59740
            
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 59740
        
        Case 59739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168805)

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
    If lErro <> SUCESSO Then Error 59741
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 59742
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 59741, 59742
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168806)

    End Select

    Exit Sub
   
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro = PreencheComboContas()
    If lErro <> SUCESSO Then Error 59743
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 59744
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 59743, 59744
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168807)

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
    If lErro <> SUCESSO Then Error 59745

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

        Case 59745
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168808)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 59746

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 59747

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 59748
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 59749
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 59746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 59747, 59748, 59749
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168809)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 59750

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 59751

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 59750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 59751

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168810)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 59752

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 59752

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168811)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCheckContas As String
Dim sConta As String

On Error GoTo Erro_PreencherRelOp
    
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sCheckContas, sConta)
    If lErro <> SUCESSO Then gError 59753

'    If Len(EmissaoDe.ClipText) = 0 Then Error 54749
'    If Len(EmissaoAte.ClipText) = 0 Then Error 54750
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 59754
                    
    'Preenche a Conta Corrente
    lErro = objRelOpcoes.IncluirParametro("TCONTA", sConta)
    If lErro <> AD_BOOL_TRUE Then gError 59755
    
    lErro = objRelOpcoes.IncluirParametro("TCONTACORRENTE", ContaCorrente.Text)
    If lErro <> AD_BOOL_TRUE Then gError 59756

    'Preenche com a Opcao Conta Corrente(Todas Contas ou uma Cnta)
    lErro = objRelOpcoes.IncluirParametro("TTODCONTAS", sCheckContas)
    If lErro <> AD_BOOL_TRUE Then gError 59757
           
    If Len(EmissaoDe.ClipText) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", CStr(DATA_NULA))
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", EmissaoDe.Text)
    End If
    If lErro <> AD_BOOL_TRUE Then gError 59758

    If Len(EmissaoAte.ClipText) = 0 Then
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", CStr(DATA_NULA))
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", EmissaoAte.Text)
    End If
    If lErro <> AD_BOOL_TRUE Then gError 59759
        
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCheckContas, sConta)
    If lErro <> SUCESSO Then gError 59760

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 59753, 59754, 59755, 59756, 59757, 59758, 59759, 59760
                
'        Case 54749
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
'            EmissaoDe.SetFocus
'
'        Case 54750
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
'            EmissaoAte.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168812)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCheckContas As String, sConta As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente e Cobrador

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
        
    'Se a opção para todos os Clientes estiver selecionada
    If TodasCtas.Value = True Then
        sCheckContas = "Todas"
        sConta = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'TEm que indicar o tipo do Cliente
        If ContaCorrente.Text = "" Then Error 59761
        sCheckContas = "Uma"
        sConta = ContaCorrente.Text
    
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(EmissaoDe.ClipText) <> "" And Trim(EmissaoAte.ClipText) <> "" Then
    
         If CDate(EmissaoDe.Text) > CDate(EmissaoAte.Text) Then Error 59762
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                                
        Case 59761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
            ContaCorrente.SetFocus
            
        Case 59762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", Err)
            EmissaoDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168813)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCheckContas As String, sConta As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
    
    sExpressao = ""
    
    'Se a opção para apenas um cliente estiver selecionada
    If sCheckContas = "Uma" Then

        sExpressao = "Conta = " & Forprint_ConvInt(Codigo_Extrai(sConta))

    End If
        
    If Trim(EmissaoDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(EmissaoDe.Text))

    End If
    
    If Trim(EmissaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(EmissaoAte.Text))

    End If

'    If giFilialEmpresa <> EMPRESA_TODA Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "CodFilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
'
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168814)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sConta As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 59763
    
    'pega  Tipo cliente e Exibe
    lErro = objRelOpcoes.ObterParametro("TTODCONTAS", sParam)
    If lErro <> SUCESSO Then gError 59764
                   
    If sParam = "Todas" Then
    
        Call TodasCtas_Click
    
    Else
        'se é apenas uma entaoo exibe esta
        lErro = objRelOpcoes.ObterParametro("TCONTA", sConta)
        If lErro <> SUCESSO Then gError 59765
                            
        ApenasCta.Value = True
        ContaCorrente.Enabled = True
        
        If sConta = "" Then
            ContaCorrente.ListIndex = -1
        Else
            ContaCorrente.Text = sConta
        End If
    End If
           
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DEMINIC", sParam)
    If lErro <> SUCESSO Then gError 59766

    Call DateParaMasked(EmissaoDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DEMFIM", sParam)
    If lErro <> SUCESSO Then gError 59767
    
    Call DateParaMasked(EmissaoAte, CDate(sParam))
           
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 59763, 59764, 59765, 59766, 59767
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168815)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 59768

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro <> SUCESSO Then Error 59769

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 59768 'Tratado na rotina chamada
    
        Case 59769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, ContaCorrente.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168816)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoAte)

End Sub

Private Sub EmissaoDe_GotFocus()
        
    Call MaskEdBox_TrataGotFocus(EmissaoDe)

End Sub

Private Sub TodasCtas_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TodasCtas_Click
    
    'Limpa e desabilita a ComboTipo
    ContaCorrente.ListIndex = -1
    ContaCorrente.Enabled = False
    TodasCtas.Value = True
    
    Exit Sub

Erro_TodasCtas_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168817)

    End Select

    Exit Sub
    
End Sub

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    'defina todos os tipos
    Call TodasCtas_Click
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168818)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub ApenasCta_Click()

Dim lErro As Long

On Error GoTo Erro_OptionUmTipo_Click
    
    'Limpa Combo Tipo e Abilita
    ContaCorrente.ListIndex = -1
    ContaCorrente.Enabled = True
    ContaCorrente.SetFocus
    
    Exit Sub

Erro_OptionUmTipo_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168819)

    End Select

    Exit Sub
    
End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    If Len(EmissaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then Error 59770

    End If

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True


    Select Case Err

        Case 59770

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168820)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    If Len(EmissaoDe.ClipText) > 0 Then

        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then Error 59771

    End If

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True


    Select Case Err

        Case 59771

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168821)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub
    
Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 59772

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case Err

        Case 59772
            EmissaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168822)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 59773

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case Err

        Case 59773
            EmissaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168823)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 59774

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case Err

        Case 59774
            EmissaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168824)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 59775

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case Err

        Case 59775
            EmissaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168825)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EXTRATOTES
    Set Form_Load_Ocx = Me
    Caption = "Extrato Conciliado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpExtratoConc"
    
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

