VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpLancEstICMSSTOcx 
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   ScaleHeight     =   2370
   ScaleWidth      =   7065
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   720
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   4800
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1845
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   840
         TabIndex        =   15
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   4335
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3330
         TabIndex        =   17
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
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
         Left            =   2925
         TabIndex        =   19
         Top             =   315
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
         Left            =   465
         TabIndex        =   18
         Top             =   285
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      Height          =   675
      Left            =   120
      TabIndex        =   10
      Top             =   1515
      Width           =   4785
      Begin VB.ComboBox EstadoAte 
         Height          =   315
         Left            =   3300
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox EstadoDe 
         Height          =   315
         Left            =   825
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label EstadoAteLabel 
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
         Left            =   2895
         TabIndex        =   12
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label30 
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
         Left            =   465
         TabIndex        =   11
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpLancEstICMSST.ctx":0000
      Left            =   855
      List            =   "RelOpLancEstICMSST.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   2916
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4770
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpLancEstICMSST.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpLancEstICMSST.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpLancEstICMSST.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "RelOpLancEstICMSST.ctx":081A
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
      Left            =   5325
      Picture         =   "RelOpLancEstICMSST.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   900
      Width           =   1575
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
      Left            =   150
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "RelOpLancEstICMSSTOcx"
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

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
                    
    'Carrega a combo Placa UF com os Estados cadastrados no BD
    lErro = Carrega_PlacaUF()
    If lErro <> SUCESSO Then gError 75059
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
                    
        Case 75059
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169667)

    End Select

    Exit Sub

End Sub

Private Function Carrega_PlacaUF() As Long
'Lê as Siglas dos Estados e alimenta a list da Combobox PlacaUF

Dim lErro As Long
Dim colSiglasUF As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_PlacaUF

    Set colSiglasUF = gcolUFs

    'Adiciona na Combo PlacaUF
    For iIndice = 1 To colSiglasUF.Count
        EstadoDe.AddItem colSiglasUF.Item(iIndice)
        EstadoAte.AddItem colSiglasUF.Item(iIndice)
    Next

    Carrega_PlacaUF = SUCESSO

    Exit Function

Erro_Carrega_PlacaUF:

    Carrega_PlacaUF = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169668)

    End Select
    
    Exit Function
    
End Function

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
        If lErro <> SUCESSO Then gError 78091

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 78091

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169669)

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
        If lErro <> SUCESSO Then gError 78092

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 78092

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169670)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 78093

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 78093
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169671)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 78094

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 78094
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169672)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 78095

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 78095
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169673)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 78096

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 78096
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169674)

    End Select

    Exit Sub

End Sub

Private Sub EstadoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstadoDe_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(EstadoDe.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(EstadoDe)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 75388

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 75389

    Exit Sub

Erro_EstadoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 75388

        Case 75389
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, EstadoDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169675)

    End Select

    Exit Sub

End Sub

Private Sub EstadoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstadoAte_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(EstadoAte.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(EstadoAte)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 75386

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 75387

    Exit Sub

Erro_EstadoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 75386

        Case 75387
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, EstadoAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169676)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_PreencherParametrosNaTela

    'Limpa a tela
    Call Limpar_Tela

    'Carrega Opções de Relatório
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 75060
                  
    'pega Estado inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TESTADODE", sParam)
    If lErro Then gError 75062
        
    For iIndice = 0 To EstadoDe.ListCount - 1
        If EstadoDe.List(iIndice) = sParam Then
            EstadoDe.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'pega Estado final e exibe
    lErro = objRelOpcoes.ObterParametro("TESTADOATE", sParam)
    If lErro Then gError 75063
        
    For iIndice = 0 To EstadoAte.ListCount - 1
        If EstadoAte.List(iIndice) = sParam Then
            EstadoAte.ListIndex = iIndice
            Exit For
        End If
    Next
        
    'pega Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro Then gError 78102
    
    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega Data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 78103

    Call DateParaMasked(DataFinal, CDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 75060 To 75063, 78102, 78103
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169677)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
        
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 75064
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 75065

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 75064
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 75065
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169678)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    EstadoDe.Text = ""
    EstadoAte.Text = ""
    
    ComboOpcoes.SetFocus
    
End Sub

Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
        
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 78099
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 78100
    
    'Data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 78101
    End If
    
    'Estado inicial não pode ser maior que o estado final
    If Trim(EstadoDe.Text) <> "" And Trim(EstadoAte.Text) <> "" Then
         If EstadoDe.Text > EstadoAte.Text Then gError 75066
    End If
               
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 75066
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_INICIAL_MAIOR", gErr)
            EstadoDe.SetFocus
        
        Case 78099
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
            
        Case 78100
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case 78101
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
                       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169679)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecucao As Boolean = False) As Long
'preenche o arquivo com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp

    'Critica as datas
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 75067
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 75068
    
    lErro = objRelOpcoes.IncluirParametro("TESTADODE", EstadoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 75069
    
    lErro = objRelOpcoes.IncluirParametro("TESTADOATE", EstadoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 75070
               
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 78104
    
    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 78105
    
    If bExecucao Then
            
        lErro = CF("RelOpLancEstICMSST_Prepara", lNumIntRel, EstadoDe.Text, EstadoAte.Text, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
        If lErro <> SUCESSO Then gError 75072
            
        lErro = Monta_Expressao_Selecao(objRelOpcoes, lNumIntRel)
        If lErro <> SUCESSO Then gError 75072
    
    End If
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 75067 To 75072, 78104, 78105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169680)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, ByVal lNumIntRel As Long) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = "NumIntRel = " & Forprint_ConvLong(lNumIntRel)

    objRelOpcoes.sSelecao = sExpressao

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169681)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 75073

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 75074

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 75073
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 75074

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169682)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 75075

    Call gobjRelatorio.Executar_Prossegue2(Me)
        
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 75075
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169683)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 75076

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 75077

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 75078

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75076
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 75077, 75078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169684)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Lista de Reg. de Entrada/Saída por Estado - ICMS Susbstituição"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpLancEstICMSST"
    
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

Private Sub EstadoAteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(EstadoAteLabel, Source, X, Y)
End Sub

Private Sub EstadoAteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(EstadoAteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

