VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOPRelEmbarqueOcx 
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LockControls    =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   6630
   Begin VB.Frame Frame3 
      Caption         =   "Clientes a excluir"
      Height          =   1995
      Left            =   135
      TabIndex        =   18
      Top             =   1845
      Width           =   6405
      Begin VB.ListBox ListClientes 
         Height          =   1410
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   360
         Width           =   3750
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   555
         Left            =   4680
         Picture         =   "RelOPRelEmbarqueOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   405
         Width           =   1530
      End
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   555
         Left            =   4680
         Picture         =   "RelOPRelEmbarqueOcx.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1170
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   780
      Left            =   135
      TabIndex        =   13
      Top             =   900
      Width           =   4245
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1485
         TabIndex        =   14
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
         Left            =   525
         TabIndex        =   2
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
         Left            =   3750
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   2775
         TabIndex        =   3
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
      Begin VB.Label dFim 
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
         Left            =   2385
         TabIndex        =   17
         Top             =   345
         Width           =   360
      End
      Begin VB.Label dIni 
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
         Height          =   240
         Left            =   135
         TabIndex        =   16
         Top             =   315
         Width           =   345
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
      Left            =   4725
      Picture         =   "RelOPRelEmbarqueOcx.ctx":21FC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOPRelEmbarqueOcx.ctx":22FE
      Left            =   990
      List            =   "RelOPRelEmbarqueOcx.ctx":2300
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2685
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4410
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "RelOPRelEmbarqueOcx.ctx":2302
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOPRelEmbarqueOcx.ctx":2480
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOPRelEmbarqueOcx.ctx":29B2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOPRelEmbarqueOcx.ctx":2B3C
         Style           =   1  'Graphical
         TabIndex        =   8
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
      Left            =   180
      TabIndex        =   12
      Top             =   405
      Width           =   615
   End
End
Attribute VB_Name = "RelOPRelEmbarqueOcx"
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

Dim giCliente As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1


Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    
    lErro = CarregaList_Clientes
    If lErro <> SUCESSO Then gError 87500
    
    Call Define_Padrao
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 87500
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172548)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long
'??? Voltar a função para correções

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    giCliente = 1
          
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172549)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela
Dim lErro As Long
Dim sParam As String
Dim objCliente As New ClassCliente
Dim iIndice As Integer
Dim iIndiceRel As Integer
Dim sListCount As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 87455
   
    'pega Cliente e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTE", sParam)
    If lErro Then gError 87456
    
'    Cliente.Text = sParam
'    Call Cliente_Validate(bSGECancelDummy)
        
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 87457

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 87458

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
    
    'Maristela (Inicio)
    'Limpa a Lista
    For iIndice = 0 To ListClientes.ListCount - 1
        ListClientes.Selected(iIndice) = False
    Next
    'Obtem o numero de Clientes selecionados na Lista
    lErro = objRelOpcoes.ObterParametro("NLISTCOUNT", sListCount)
    If lErro <> SUCESSO Then gError 90519
    'Percorre toda a Lista
    For iIndice = 0 To ListClientes.ListCount - 1
        'Percorre todos os Clientes que foram slecionados
        For iIndiceRel = 1 To StrParaInt(sListCount)
            lErro = objRelOpcoes.ObterParametro("NLIST" & SEPARADOR & iIndiceRel, sParam)
            If lErro <> SUCESSO Then gError 90520
            'Se o cliente não foi excluido
            If sParam = Codigo_Extrai(ListClientes.List(iIndice)) Then
                'Marca os Clientes que foram gravados
                ListClientes.Selected(iIndice) = True
            End If
        Next
    Next
    'Maristela (Fim)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 87455 To 87458
        
        Case 90519, 90520

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172550)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoCliente = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 87459
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 87460

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 87459
        
        Case 87460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172551)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 87461
    
    Call Limpa_ListClientes
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
'    DescCliente.Caption = ""
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 87461
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172552)

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
Dim iIndice As Integer
Dim sCliente As String
Dim sListCount As String
Dim iNCliente As Integer

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 87463
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 87464
         
'    lErro = objRelOpcoes.IncluirParametro("NCLIENTE", sCliente)
'    If lErro <> AD_BOOL_TRUE Then gError 87465
'
'    lErro = objRelOpcoes.IncluirParametro("TCLIENTE", Cliente.Text)
'    If lErro <> AD_BOOL_TRUE Then gError 87466

    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 87467
    
    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 87468
    
    'Maristela (Inicio)
     iNCliente = 1
    'Percorre toda a Lista
    For iIndice = 0 To ListClientes.ListCount - 1
        If ListClientes.Selected(iIndice) = True Then
            sCliente = Codigo_Extrai(ListClientes.List(iIndice))
            'Inclui todos os Clientes que foram slecionados
            lErro = objRelOpcoes.IncluirParametro("NLIST" & SEPARADOR & iNCliente, sCliente)
            If lErro <> AD_BOOL_TRUE Then gError 90517
            iNCliente = iNCliente + 1
        End If
    Next
    sListCount = iNCliente - 1
    'Inclui o numero de Clientes selecionados na Lista
    lErro = objRelOpcoes.IncluirParametro("NLISTCOUNT", sListCount)
    If lErro <> AD_BOOL_TRUE Then gError 90518
    'Maristela (Fim)
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 87469
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 87463 To 87469
        
        Case 90517, 90518
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172553)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 87470

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 87471

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 87472
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 87470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 87471, 87472

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172554)

    End Select

    Exit Sub

End Sub


Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim colCliente As New Collection

On Error GoTo Erro_BotaoExecutar_Click

    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 87491
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 87492
    
    lErro = RetiraNomes_Sel(colCliente)
    If lErro <> SUCESSO Then gError 87510

    lErro = Atualiza_RelEmbarqueCli(colCliente)
    If lErro <> SUCESSO Then gError 87511

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 87473

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 87473
        
        Case 87491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
        
        Case 87492
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
        
        Case 87510, 87511
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172555)

    End Select

    Exit Sub

End Sub


Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 87474

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 87475

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 87476

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 87477
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 87474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 87475, 87476, 87477

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172556)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório
'???Verificar função com Fernando

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

'   If sCliente <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(CLng(sCliente))

'   If sCliente_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_F))
'
'    End If
    
'    If Trim(DataInicial.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))
'
'    End If
'
'    If Trim(DataFinal.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))
'
'    End If
'
'    If sExpressao <> "" Then sExpressao = sExpressao & " E "
'    sExpressao = sExpressao & "NORDENACAO = " & Forprint_ConvInt(CInt(sOrdenacao))
'
'    If sExpressao <> "" Then sExpressao = sExpressao & " E "
'    sExpressao = sExpressao & "NDEVOLUCOES = " & Forprint_ConvInt(CInt(Devolucoes.Value))
     
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172557)

    End Select

    Exit Function

End Function


Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'critica Cliente Inicial e Final
    
'    If Cliente.Text <> "" Then William
'        sCliente = CStr(LCodigo_Extrai(Cliente.Text))
'    Else
'        sCliente = ""
'    End If
        
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 87479
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function


Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                             
         Case 87479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172558)

    End Select

    Exit Function

End Function


'Private Sub Cliente_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objCliente As New ClassCliente
'
'On Error GoTo Erro_Cliente_Validate
'
'    If Len(Trim(Cliente.Text)) > 0 Then
'
'        'Tenta ler o Cliente (NomeReduzido ou Código)
'        lErro = TP_Cliente_Le2(Cliente, objCliente, 0)
'        If lErro <> SUCESSO Then gError 87480
'
'        'Preenche a Razão Social do Cliente
'        If Len(Trim(objCliente.sRazaoSocial)) > 0 Then
'            DescCliente.Caption = objCliente.sRazaoSocial
'        End If
'
'    End If
'
'    giCliente = 1
'
'    Exit Sub
'
'Erro_Cliente_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 87480
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172559)
'
'    End Select
'
'End Sub

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
        If lErro <> SUCESSO Then gError 87481

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 87481

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172560)

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
        If lErro <> SUCESSO Then gError 87482

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 87482

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172561)

    End Select

    Exit Sub

End Sub

'Private Sub LabelCliente_Click()
'
'Dim objCliente As New ClassCliente
'Dim colSelecao As Collection
'
'    giCliente = 1
'
'    If Len(Trim(Cliente.Text)) > 0 Then
'        'Preenche com o cliente da tela
'        objCliente.lCodigo = LCodigo_Extrai(Cliente.Text)
'    End If
'
'    'Chama Tela ClientesLista
'    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)
'
'End Sub


'Private Sub objEventoCliente_evSelecao(obj1 As Object)
'
'Dim objCliente As ClassCliente
'
'    Set objCliente = obj1
'
'    'Preenche campo Cliente
'    If giCliente = 1 Then
'
'        Cliente.Text = CStr(objCliente.lCodigo)
'        Call Cliente_Validate(bSGECancelDummy)
'
'    End If
'
'    Me.Show
'
'    Exit Sub
'
'End Sub


Private Sub UpDown1_DownClick()
'??? voltar a sub e incluir código de erro

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 87483

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 87483
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172562)

    End Select

    Exit Sub

End Sub


Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 87484

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 87484
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172563)

    End Select

    Exit Sub

End Sub


Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 87485

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 87485
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172564)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 87486

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 87486
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172565)

    End Select

    Exit Sub

End Sub

Function CarregaList_Clientes() As Long

Dim lErro As Long
Dim colCodigoClientes As New Collection
Dim objCliente As ClassCliente

On Error GoTo Erro_CarregaList_Clientes
    
    lErro = Clientes_Le_Todos(colCodigoClientes)
    If lErro <> SUCESSO Then gError 87498

    'preenche cada ComboBox País com os objetos da colecao colCodigoDescricao
    For Each objCliente In colCodigoClientes
        ListClientes.AddItem CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sNomeReduzido
    Next

    CarregaList_Clientes = SUCESSO

    Exit Function

Erro_CarregaList_Clientes:

    CarregaList_Clientes = gErr

    Select Case gErr

        Case 87498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172566)

    End Select

    Exit Function

End Function

Public Function RetiraNomes_Sel(colCliente As Collection) As Long
'Retira da combo todos os nomes que não estão selecionados

Dim iIndice As Integer
Dim lCodCliente As Long

    For iIndice = 0 To ListClientes.ListCount - 1
        If ListClientes.Selected(iIndice) = False Then
            lCodCliente = LCodigo_Extrai(ListClientes.List(iIndice))
            colCliente.Add lCodCliente
        End If
    Next
    
End Function


Public Function Atualiza_RelEmbarqueCli(colCliente As Collection) As Long
'Atualiza o campo único (CodCliente) da tabela RelEmbarqueCli

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim lCodCliente As Long

On Error GoTo Erro_Atualiza_RelEmbarqueCli

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 87501
    Next
    
    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 87502
    
    'Le todos os registros da tabela
    lErro = Comando_ExecutarPos(alComando(1), "SELECT CodCliente FROM RelEmbarqueCli", 0, lCodCliente)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87503
    
    'Posiciona ponteiro no primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87504
    
    If lErro = AD_SQL_SUCESSO Then
    
        'Caso sejam encontrados registros
        Do While lErro = AD_SQL_SUCESSO
        
            'Apaga os registro existentes
            lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM RelEmbarqueCli", alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 87506
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87512
            
        Loop
            'Para cada objeto dentro da coleção
            For iIndice = 1 To colCliente.Count
                lCodCliente = colCliente(iIndice)
                    
                'Inserir registros
                lErro = Comando_Executar(alComando(1), "INSERT INTO RelEmbarqueCli (CodCliente) VALUES (?)", lCodCliente)
                If lErro <> AD_SQL_SUCESSO Then gError 87507
            Next
        
        'Caso a tabela já esteja vazia
    Else
        
        'Para cada objeto dentro da coleção
        For iIndice = 1 To colCliente.Count
            lCodCliente = colCliente(iIndice)
            
            'Inserir registros
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelEmbarqueCli (CodCliente) VALUES (?)", lCodCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 87508
        Next
        
    End If

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 87509

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
Atualiza_RelEmbarqueCli = SUCESSO

    Exit Function
    
Erro_Atualiza_RelEmbarqueCli:

    Atualiza_RelEmbarqueCli = gErr
    
    Select Case gErr
    
        Case 87501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 87502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 87503, 87504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELEMBARQUECLI", gErr)
            
        Case 87505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOK_RELEMBARQUECLI", gErr)
        
        Case 87506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELEMBARQUECLI", gErr)
            
        Case 87507, 87508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELEMBARQUECLI", gErr)

        Case 87509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172567)
            
    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    'Fecha Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function


Private Sub BotaoMarcar_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListClientes.ListCount - 1
        ListClientes.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoDesmarcar_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListClientes.ListCount - 1
        ListClientes.Selected(iIndice) = False
    Next

End Sub

Sub Limpa_ListClientes()

Dim iIndice As Integer

    For iIndice = 0 To ListClientes.ListCount - 1
        ListClientes.Selected(iIndice) = False
    Next

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_FAT_CLIENTE
    Set Form_Load_Ocx = Me
    Caption = "Relação de Embarque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelEmbarque"
    
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
    
'    If KeyCode = KEYCODE_BROWSER Then
'
'        If Me.ActiveControl Is Cliente Then
'            Call LabelCliente_Click
'        End If
'
'    End If

End Sub


'Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelCliente, Source, X, Y)
'End Sub
'
'Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
'End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


Function Clientes_Le_Todos(colClientes As Collection) As Long
'Preenche colClientes com objetos da classe ClassCliente percorrendo a tabela de Clientes

Dim lErro As Long
Dim lComando As Long
Dim objCliente As ClassCliente
Dim tCliente As typeCliente

On Error GoTo Erro_Clientes_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87494
    
    With tCliente
    
        .sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        .sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
        .sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
        .sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
        .sCgc = String(STRING_CGC, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sObservacao2 = String(STRING_CLIENTE_OBSERVACAO, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        
        lErro = Comando_Executar(lComando, "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco,ProxCodFilial, CodPadraoCobranca FROM Clientes ORDER BY NomeReduzido", _
        .lCodigo, .sRazaoSocial, .sNomeReduzido, .iTipo, .sObservacao, .dLimiteCredito, .iCondicaoPagto, .dDesconto, .iCodMensagem, .iTabelaPreco, .iProxCodFilial, .iCodPadraoCobranca)
    
        If lErro <> AD_SQL_SUCESSO Then gError 87495
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87496
        
        Do While lErro <> AD_SQL_SEM_DADOS
            
            Set objCliente = New ClassCliente
            
            objCliente.lCodigo = .lCodigo
            objCliente.sRazaoSocial = .sRazaoSocial
            objCliente.sNomeReduzido = .sNomeReduzido
            objCliente.iTipo = .iTipo
            objCliente.sObservacao = .sObservacao
            objCliente.dLimiteCredito = .dLimiteCredito
            objCliente.iCondicaoPagto = .iCondicaoPagto
            objCliente.dDesconto = .dDesconto
            objCliente.iCodMensagem = .iCodMensagem
            objCliente.iTabelaPreco = .iTabelaPreco
            objCliente.iProxCodFilial = .iProxCodFilial
            objCliente.iCodPadraoCobranca = .iCodPadraoCobranca
            
            colClientes.Add objCliente
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87497
    
        Loop
    
    End With
        
    lErro = Comando_Fechar(lComando)
    
    Clientes_Le_Todos = SUCESSO

    Exit Function

Erro_Clientes_Le_Todos:

    Clientes_Le_Todos = gErr

    Select Case gErr

        Case 87494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 87495, 87496, 87497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172568)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


