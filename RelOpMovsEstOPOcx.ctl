VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpMovsEstOPOcx 
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   KeyPreview      =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   7455
   Begin VB.Frame Frame1 
      Caption         =   "Data"
      Height          =   1305
      Left            =   4665
      TabIndex        =   14
      Top             =   1905
      Width           =   2625
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   2115
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   1110
         TabIndex        =   6
         Top             =   315
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   2115
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   795
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   1095
         TabIndex        =   7
         Top             =   795
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   705
         TabIndex        =   23
         Top             =   345
         Width           =   315
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   22
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4935
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "RelOpMovsEstOPOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "RelOpMovsEstOPOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpMovsEstOPOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpMovsEstOPOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordem de Produção"
      Height          =   1305
      Left            =   135
      TabIndex        =   18
      Top             =   1905
      Width           =   4380
      Begin VB.TextBox OpFinal 
         Height          =   300
         Left            =   1140
         TabIndex        =   5
         Top             =   825
         Width           =   2055
      End
      Begin VB.TextBox OpInicial 
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label LabelOpFinal 
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
         Left            =   735
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   885
         Width           =   360
      End
      Begin VB.Label LabelOpInicial 
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
         Left            =   780
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   780
      Left            =   120
      TabIndex        =   17
      Top             =   975
      Width           =   4380
      Begin VB.OptionButton OpStatus 
         Caption         =   "Encerradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2850
         TabIndex        =   3
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton OpStatus 
         Caption         =   "Abertas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1575
         TabIndex        =   2
         Top             =   315
         Width           =   1215
      End
      Begin VB.OptionButton OpStatus 
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
         Height          =   330
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   300
         Width           =   1110
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
      Left            =   5235
      Picture         =   "RelOpMovsEstOPOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   855
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMovsEstOPOcx.ctx":0A96
      Left            =   1545
      List            =   "RelOpMovsEstOPOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   855
      TabIndex        =   21
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMovsEstOPOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giOp_Inicial As Integer

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoOp = New AdmEvento

    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170130)

    End Select

    Exit Sub

End Sub

Sub Define_Padrao()
'Preenche a tela com as opções padrão de FilialEmpresa

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giOp_Inicial = 1
   
    OpStatus(1).Value = True

    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170131)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 93501

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 93501

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170132)

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
        If lErro <> SUCESSO Then gError 93502

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 93502

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170133)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpInicial_Click

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpInicial.Text

    End If

    Call Chama_Browse_OP(objOp)

    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170134)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpFinal_Click

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpFinal.Text

    End If

    Call Chama_Browse_OP(objOp)

   Exit Sub

Erro_LabelOpFinal_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170135)

    End Select

    Exit Sub

End Sub

Private Sub Chama_Browse_OP(objOp As ClassOrdemDeProducao)

Dim lErro As Long
Dim colSelecao As Collection
Dim iOpStatus As Integer
Dim iIndice As Integer

On Error GoTo Erro_Chama_Browse_OP

   'verifica status selecionado
    For iIndice = 0 To 2
        If OpStatus(iIndice).Value = True Then iOpStatus = iIndice
    Next

    Select Case iOpStatus

        Case 0
            Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOp, objEventoOp)

        Case 1
            Call Chama_Tela("OrdemProdListaModal", colSelecao, objOp, objEventoOp)

        Case 2
            Call Chama_Tela("OrdProdBaixadasListaModal", colSelecao, objOp, objEventoOp)

   End Select

   Exit Sub

Erro_Chama_Browse_OP:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170136)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOp = obj1

    If giOp_Inicial = 1 Then

        OpInicial.Text = objOp.sCodigo
        
    Else

        OpFinal.Text = objOp.sCodigo

    End If

    Me.Show
    
    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170137)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 38537

    'pega Ordem de Producao Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TOPINIC", sParam)
    If lErro Then gError 38538

    OpInicial.Text = sParam

    'pega Ordem de Producao Final e exibe
    lErro = objRelOpcoes.ObterParametro("TOPFIM", sParam)
    If lErro Then gError 38539

    OpFinal.Text = sParam

    'Pega status e exibe
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then gError 38540

    OpStatus(CInt(sParam)) = True

    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 93508

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 93509

    Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 38537 To 38540
       
       
        Case 93508 To 93509
       
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170138)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoOp = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29893
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 38535

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 38535
        
        Case 29893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170139)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub


Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'ordem de produção inicial não pode ser maior que a final
    If Trim(OpInicial.Text) <> "" And Trim(OpFinal.Text) <> "" Then

        If OpInicial.Text > OpFinal.Text Then gError 38542

    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 93510
    
    End If
    

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 93510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus

        Case 38542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", gErr)
            OpInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170140)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela
    Call Define_Padrao

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
Dim sOp_I As String
Dim sOp_F As String

Dim sStatus As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp
          
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 38545

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 38546

    sOp_I = OpInicial.Text
    lErro = objRelOpcoes.IncluirParametro("TOPINIC", sOp_I)
    If lErro <> AD_BOOL_TRUE Then gError 38547

    sOp_F = OpFinal.Text
    lErro = objRelOpcoes.IncluirParametro("TOPFIM", sOp_F)
    If lErro <> AD_BOOL_TRUE Then gError 38548
  
    'verifica opção selecionada
    For iIndice = 0 To 2
        If OpStatus(iIndice).Value = True Then sStatus = CStr(iIndice)
    Next
    
    If sStatus = "0" Then gobjRelatorio.sNomeTsk = "MovEstOp"
    If sStatus = "1" Then gobjRelatorio.sNomeTsk = "MEstOpAb"
    If sStatus = "2" Then gobjRelatorio.sNomeTsk = "MEstOpBa"

    lErro = objRelOpcoes.IncluirParametro("NSTATUS", sStatus)
    If lErro <> AD_BOOL_TRUE Then gError 38549
  

    'Preenche data Inicial
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 93506
    
    'Preenche  data final
    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 93507


    lErro = Monta_Expressao_Selecao(objRelOpcoes, sOp_I, sOp_F, sStatus)
    If lErro <> SUCESSO Then gError 38550

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

       
        Case 38545 To 38550
        
        
        Case 93506 To 93507
        
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170141)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 38551

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 38552

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 38551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 38552

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170142)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 38553

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 38553

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170143)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 38554

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 38555

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 38556

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 38554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 38555, 38556

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170144)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sOp_I As String, sOp_F As String, sStatus As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sOp_I <> "" Then sExpressao = "OrdemProducao >= " & Forprint_ConvTexto(sOp_I)

    If sOp_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "OrdemProducao <= " & Forprint_ConvTexto(sOp_F)

    End If
     
    
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
    
    
    
''    If sExpressao <> "" Then sExpressao = sExpressao & " E "
''    sExpressao = "NSTATUS = " & Forprint_ConvInt(CInt(sStatus))
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If


    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170145)

    End Select

    Exit Function

End Function

Private Function Valida_OrdProd(sCodigoOP As String) As Long
Dim objOp As New ClassOrdemDeProducao
Dim lErro As Long
Dim iIndice As Integer
Dim iOpStatus As Integer

On Error GoTo Erro_Valida_OrdProd

    objOp.iFilialEmpresa = giFilialEmpresa
    objOp.sCodigo = sCodigoOP
    
    'verifica status selecionado
    For iIndice = 0 To 2
       If OpStatus(iIndice).Value = True Then iOpStatus = iIndice
    Next

    'se a opção de status é "Abertas"
    If iOpStatus = 1 Then

        lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
        If lErro <> SUCESSO And lErro <> 34455 Then Error 38557

        If lErro = 34455 Then Error 38558
    
    End If

    'Se a opção de status é "Encerradas"
    If iOpStatus = 2 Then

        lErro = CF("OPBaixada_Le_SemItens", objOp)
        If lErro <> SUCESSO And lErro <> 34459 Then Error 38559

        If lErro = 34459 Then Error 38560

    End If
    
    'Se a opção de status é "Todas"
    If iOpStatus = 0 Then

        lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
        If lErro <> SUCESSO And lErro <> 34455 Then
            
            Error 38561

        Else
        
            If lErro <> SUCESSO Then

                lErro = CF("OPBaixada_Le_SemItens", objOp)
                If lErro <> SUCESSO And lErro <> 34459 Then Error 38562

                If lErro = 34459 Then Error 38563
            
            End If


        End If

    End If

    Valida_OrdProd = SUCESSO

    Exit Function

Erro_Valida_OrdProd:

    Valida_OrdProd = Err

    Select Case Err

        Case 38557, 38559, 38561, 38562
        
        Case 38558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_ABERTA_INEXISTENTE", Err)
            
        Case 38560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_ENCERRADA_INEXISTENTE", Err)
                              
        Case 38563
           lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170146)

    End Select

    Exit Function

End Function

Private Sub OpInicial_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_OpInicial_Validate

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpInicial.Text)
        If lErro <> SUCESSO Then Error 38564
        
    End If

    Exit Sub

Erro_OpInicial_Validate:

    Cancel = True


    Select Case Err

        Case 38564
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170147)

    End Select

    Exit Sub

End Sub

Private Sub OpFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpFinal_Validate

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpFinal.Text)
        If lErro <> SUCESSO Then Error 38565
    
    End If

    Exit Sub

Erro_OpFinal_Validate:

    Cancel = True


    Select Case Err

        Case 38565
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170148)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_MOVIMENTO_ESTOQUE_OP
    Set Form_Load_Ocx = Me
    Caption = "Relação de Movimentos de Estoque para cada Ordem de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpMovsEstOP"
    
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

Private Sub UpDown1_DownClick()
Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 93503

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 93503
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170149)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37444

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 37444
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170150)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 93504

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 93504
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170151)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 37446

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 37446
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170152)

    End Select

    Exit Sub

End Sub

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
        
        If Me.ActiveControl Is OpInicial Then
            Call LabelOpInicial_Click
        ElseIf Me.ActiveControl Is OpFinal Then
            Call LabelOpFinal_Click
        End If
    
    End If

End Sub

Private Sub LabelOpFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpFinal, Source, X, Y)
End Sub

Private Sub LabelOpFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelOpInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpInicial, Source, X, Y)
End Sub

Private Sub LabelOpInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpInicial, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

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

