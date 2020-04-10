VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpEmissaoNFFaturaOcx 
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   ScaleHeight     =   3501.351
   ScaleMode       =   0  'User
   ScaleWidth      =   4650
   Begin VB.Frame Frame1 
      Caption         =   "Opções"
      Height          =   1050
      Left            =   30
      TabIndex        =   14
      Top             =   1950
      Width           =   4530
      Begin VB.OptionButton OptGerarDocs 
         Caption         =   "Apenas gerar .docs"
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
         Left            =   135
         TabIndex        =   17
         Top             =   780
         Width           =   3855
      End
      Begin VB.OptionButton OptAmbos 
         Caption         =   "Executar o relatório e gerar .docs"
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
         Left            =   135
         TabIndex        =   16
         Top             =   510
         Width           =   3855
      End
      Begin VB.OptionButton OptImpNormal 
         Caption         =   "Apenas executar o relatório"
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
         Left            =   135
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   3855
      End
   End
   Begin VB.CheckBox NaoImprimirDataSaida 
      Caption         =   "Não imprimir data de saída nas notas fiscais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   165
      TabIndex        =   13
      Top             =   3135
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3465
      ScaleHeight     =   495
      ScaleMode       =   0  'User
      ScaleWidth      =   1065
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   30
      Width           =   1125
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpEmissaoNFFaturaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpEmissaoNFFaturaOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
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
      Height          =   552
      Left            =   1425
      Picture         =   "RelOpEmissaoNFFaturaOcx.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   30
      Width           =   1845
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Nota Fiscal"
      Height          =   1320
      Left            =   30
      TabIndex        =   5
      Top             =   645
      Width           =   4530
      Begin VB.ComboBox TipoFormulario 
         Height          =   315
         ItemData        =   "RelOpEmissaoNFFaturaOcx.ctx":07B2
         Left            =   1680
         List            =   "RelOpEmissaoNFFaturaOcx.ctx":07C8
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   -165
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   675
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   975
      End
      Begin MSMask.MaskEdBox NFiscalInicial 
         Height          =   300
         Left            =   675
         TabIndex        =   2
         Top             =   915
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFiscalFinal 
         Height          =   300
         Left            =   2130
         TabIndex        =   3
         Top             =   915
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label DescMod 
         BorderStyle     =   1  'Fixed Single
         Height          =   645
         Left            =   1650
         TabIndex        =   18
         Top             =   225
         Width           =   2805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Formulário:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   -105
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label6 
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
         Left            =   1755
         TabIndex        =   8
         Top             =   975
         Width           =   360
      End
      Begin VB.Label Label14 
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
         Left            =   300
         TabIndex        =   7
         Top             =   975
         Width           =   300
      End
      Begin VB.Label LabelSerie 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   285
         Width           =   510
      End
   End
End
Attribute VB_Name = "RelOpEmissaoNFFaturaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

'Variavel global
Private glNumAcessosTimer As Long

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    If gsNomeEmpresa = "PDL Sistemas Ltda." Then NaoImprimirDataSaida.Visible = True
    
    Set objEventoSerie = New AdmEvento
    
    TipoFormulario.ListIndex = 0
    
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 122647
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 122624 'Tratado na Rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168459)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
  
    Set objEventoSerie = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes, Optional vParam As Variant) As Long

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim bEncontrou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 122625
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    If Not IsMissing(vParam) Then
        
        'Se passou a Série
        If vParam <> "" Then
                    
            objSerie.iFilialEmpresa = giFilialEmpresa
            
            lErro = CF("Serie_FilialEmpresa_Customiza", objSerie)
            If lErro <> SUCESSO Then gError 126937
            
            objSerie.sSerie = CStr(vParam)
            
            'Lê a Serie no BD
            lErro = CF("Serie_Le", objSerie)
            If lErro <> SUCESSO And lErro <> 22202 Then gError 122626
            
            'Se não encontrou Erro
            If lErro = 22202 Then gError 122627
                       
            bEncontrou = False
            
            For iIndice = 0 To TipoFormulario.ListCount - 1
            
                If TipoFormulario.ItemData(iIndice) = objSerie.iTipoFormulario Then
                    
                    TipoFormulario.ListIndex = iIndice
                    bEncontrou = True
                    Exit For
                End If
                
            Next
            
            If bEncontrou = False Then gError 122628
        
        End If
            
        For iIndice = 0 To Serie.ListCount - 1
            
            If Trim(Serie.List(iIndice)) = Trim(objSerie.sSerie) Then
                
                Serie.ListIndex = iIndice
                Exit For
            End If
        
        Next
        
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 122625
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 122626
        
        Case 122627
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
        
        Case 122628
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULARIO_IMCOMPATIVEL", gErr, objSerie.sSerie)
        
        Case 126937
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168460)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros() As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long

On Error GoTo Erro_Critica_Parametros
          
    'Verifica se a Série Foi Preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 122629
    
    'Verifica se a Nota Inicial Foi Preenchida
    If Len(Trim(NFiscalInicial.Text)) = 0 Then gError 122630
    
    'Verifica se a Nota Final foi preechida
    If Len(Trim(NFiscalFinal.Text)) = 0 Then gError 122631
      
    'Verifica se o numero da Nota Fiscal inicial é maior que o da final
    If Len(Trim(NFiscalInicial.ClipText)) > 0 And Len(Trim(NFiscalFinal.ClipText)) > 0 Then
    
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then gError 122632
    
    End If
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr

        Case 122632
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
           
        Case 122629
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 122630
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DE_NAO_PREENCHIDO", gErr)
        
        Case 122631
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168461)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 122633
    
    TipoFormulario.ListIndex = 0
    
    Serie.ListIndex = -1
    Serie.SetFocus
    TipoFormulario.ListIndex = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 122633 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168462)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then gError 122634
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 122635
    
    lErro = objRelOpcoes.IncluirParametro("NBORDERO", "0")
    If lErro <> AD_BOOL_TRUE Then gError 122636
    
    lErro = objRelOpcoes.IncluirParametro("NNFISCALINIC", NFiscalInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 122636

    lErro = objRelOpcoes.IncluirParametro("NNFISCALFIM", NFiscalFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 122637
   
    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then gError 122638
    
    lErro = objRelOpcoes.IncluirParametro("TSEMDATASAIDA", IIf(NaoImprimirDataSaida.Value, "S", "N"))
    If lErro <> AD_BOOL_TRUE Then gError 122638
    
    '################################################
    'Inserido por Wagner 24/10/2005
    If bExecutando Then
    
        lErro = CF("RelEmissaoNF_Prepara", StrParaLong(NFiscalInicial.Text), StrParaLong(NFiscalFinal.Text), Serie.Text, giFilialEmpresa, lNumIntRel)
        If lErro <> SUCESSO Then gError 140570
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 140571
    
    End If
    '################################################
   
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 122634 To 122638, 140570, 140571 'Tratado na Rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168463)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim lFaixaFinal As Long
Dim vbMsgRes As VbMsgBoxResult
Dim sDanfe As String
Dim colNF As New Collection
Dim objNF As ClassNFiscal
Dim objRelatorio As AdmRelatorio
Dim iFilialEmpresa As Integer
Dim bJaConfigImpr As Boolean

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 122648

    objSerie.sSerie = Serie.Text
    objSerie.iFilialEmpresa = giFilialEmpresa

    lErro = CF("Serie_FilialEmpresa_Customiza", objSerie)
    If lErro <> SUCESSO Then gError 126938

    'Lock na Tabela Série para a Impreessão
    lErro = CF("Serie_Lock_ImpressaoNFiscal", objSerie)
    If lErro <> SUCESSO And lErro <> 60387 Then gError 122649

    'Se não encontrou a Série --> ERRO
    If lErro = 60387 Then gError 122650

    'Dá Mensagem ao usuário caso seja Reimpressão
    If CLng(NFiscalInicial.Text) < objSerie.lProxNumNFiscalImpressa Then

        'Verifica se a Faixa Final também não é menor que a que está no BD
        If CLng(NFiscalFinal.Text) < objSerie.lProxNumNFiscalImpressa Then
            lFaixaFinal = CLng(NFiscalFinal.Text)
        Else
            lFaixaFinal = objSerie.lProxNumNFiscalImpressa - 1
        End If

        'Avisa Reimpressao de Nota e Pede Confirmação
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NFISCAL_REIMPRESSA", CLng(NFiscalInicial.Text), lFaixaFinal)
        If vbMsgRes = vbNo Then gError 122651

    End If

    'Altera campo Imprimindo = RELATORIO_NFISCAL_IMPRIMINDO
    lErro = CF("Serie_Altera_Imprimindo", objSerie)
    If lErro <> SUCESSO And lErro <> 61008 Then gError 122652
    Me.Enabled = False

    If Len(Trim(objSerie.sNomeTsk)) > 0 Then

        gobjRelatorio.sNomeTsk = objSerie.sNomeTsk

    Else
        
        Select Case TipoFormulario.ItemData(TipoFormulario.ListIndex)
            
            Case TIPO_FORMULARIO_NFISCAL_FATURA
                gobjRelatorio.sNomeTsk = "nffat"
                
            Case TIPO_FORMULARIO_NFISCAL_FATURA_SERVICO
                gobjRelatorio.sNomeTsk = "nffatser"
                
            Case TIPO_FORMULARIO_NFISCAL_FATURA_FRETE
                gobjRelatorio.sNomeTsk = "nffretef"
                
            Case TIPO_FORMULARIO_NFISCAL
                gobjRelatorio.sNomeTsk = "nfiscal"
                
            Case TIPO_FORMULARIO_NFISCAL_SERVICO
                gobjRelatorio.sNomeTsk = "nfserv"
                
            Case TIPO_FORMULARIO_NFISCAL_FRETE
                gobjRelatorio.sNomeTsk = "nffrete"
        
        End Select
                               
        'If TipoFormulario.ItemData(TipoFormulario.ListIndex) = TIPO_FORMULARIO_NFISCAL_FATURA Then
'            gobjRelatorio.sNomeTsk = "nffat"
        'ElseIf TipoFormulario.ItemData(TipoFormulario.ListIndex) = TIPO_FORMULARIO_NFISCAL_FATURA_SERVICO Then
'             gobjRelatorio.sNomeTsk = "nffatser"
        'ElseIf TipoFormulario.ItemData(TipoFormulario.ListIndex) = TIPO_FORMULARIO_NFISCAL_FATURA_FRETE Then
'            gobjRelatorio.sNomeTsk = "nffretef"
        'End If

    End If
       
    bJaConfigImpr = False
    If Not OptGerarDocs.Value Then
        bJaConfigImpr = True
        lErro = gobjRelatorio.Executar_Prossegue
        If lErro <> SUCESSO And lErro <> 7072 Then gError 122653
    End If
    
    If Not OptImpNormal.Value Then
    
        iFilialEmpresa = giFilialEmpresa
        If giFilialEmpresa > 50 Then giFilialEmpresa = giFilialEmpresa - 50
        lErro = CF("NFe_Le_Faixa", giFilialEmpresa, Serie.Text, StrParaLong(NFiscalInicial.Text), StrParaLong(NFiscalFinal.Text), colNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        For Each objNF In colNF
            
            lErro = CF("NFe_Obtem_Nome_Danfe", objNF, sDanfe)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Set objRelatorio = New AdmRelatorio
        
            If bJaConfigImpr Then
                objRelatorio.bConfiguraImpressora = False
            Else
                objRelatorio.bConfiguraImpressora = True
                bJaConfigImpr = True
            End If
            lErro = objRelatorio.ExecutarDireto("Emissão das Notas Fiscais Fatura", "", 2, gobjRelatorio.sNomeTsk, "NNFISCALINIC", objNF.lNumNotaFiscal, "NNFISCALFIM", objNF.lNumNotaFiscal, "TSERIE", objNF.sSerie, "NBORDERO", "0", "TTO_EMAIL", "", "TSUBJECT", "", "TALIASATTACH", "", "TMAILARQ", sDanfe)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Next
        giFilialEmpresa = iFilialEmpresa
    
    End If

    'Cancelou o relatório
    If lErro = 7072 Then gError 122654
    
    If gobjFAT.lCodCliGov <> 0 Then
        gobjRelatorio.sNomeTsk = "danfeter"
        
        lErro = gobjRelatorio.Executar_Prossegue
        If lErro <> SUCESSO And lErro <> 7072 Then gError 122653
    
        'Cancelou o relatório
        If lErro = 7072 Then gError 122654
    End If

    Timer1.Interval = INTERVALO_MONITORAMENTO_IMPRESSAO_NF

    Exit Sub

Erro_BotaoExecutar_Click:

    If iFilialEmpresa <> 0 Then giFilialEmpresa = iFilialEmpresa

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 122648, 122649, 122651, 122652, 122653, 126938

        Case 122650
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)

        Case 122654 'Cancelou o relatório
            Unload Me

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168464)

    End Select

    'Faz Unlock da Tabela
    lErro = CF("Serie_Unlock_ImpressaoNF", objSerie)

    Exit Sub

End Sub

Private Sub LabelSerie_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieListaModal
    Call Chama_Tela("SerieListaModal", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Sub NFiscalInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sNumero As String

On Error GoTo Erro_NFiscalInicial_Validate

    If Len(Trim(NFiscalInicial.Text)) > 0 Then
        sNumero = NFiscalInicial.Text
    End If

    lErro = Critica_Numero(sNumero)
    If lErro <> SUCESSO Then gError 122656

    Exit Sub

Erro_NFiscalInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 122656

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168465)

    End Select

    Exit Sub

End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sNumero As String

On Error GoTo Erro_NFiscalFinal_Validate

    If Len(Trim(NFiscalFinal.Text)) > 0 Then
        sNumero = NFiscalFinal.Text
    End If

    lErro = Critica_Numero(sNumero)
    If lErro <> SUCESSO Then gError 122657

    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 122657

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168466)

    End Select

    Exit Sub

End Sub

'Private Function Carrega_Serie(iTipoFormulario As Integer) As Long
Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    'lErro = CF("Series_Le_TipoFormulario", colSerie, iTipoFormulario)
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 122639
    
    Serie.Clear
    
    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
    
    For iIndice = 1 To colSerie.Count
        
        If colSerie.Item(iIndice).lProxNumNFiscal > colSerie.Item(iIndice).lProxNumNFiscalImpressa Then
            Serie.ListIndex = iIndice - 1
            Exit For
        End If
    Next
    
    Carrega_Serie = SUCESSO
    
    Exit Function
    
Erro_Carrega_Serie:

    Carrega_Serie = gErr
    
    Select Case gErr
    
        Case 122639 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168467)
            
    End Select
    
    Exit Function

End Function

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie, iIndice As Integer

    Set objSerie = obj1

    'Coloca a Série na Tela
    For iIndice = 0 To Serie.ListCount - 1
        
        If Trim(Serie.List(iIndice)) = Trim(objSerie.sSerie) Then
            
            Serie.ListIndex = iIndice
            Exit For
        
        End If
    
    Next
    
    Call Serie_Validate(bSGECancelDummy)

    Exit Sub

End Sub

Private Sub Serie_Click()

Dim lErro As Long

On Error GoTo Erro_Serie_Click
    
    'Traz os números default
    lErro = Traz_Numeros_Default()
    If lErro <> SUCESSO Then gError 122640
        
    Exit Sub
    
Erro_Serie_Click:

    Select Case gErr
    
        Case 122640 'Tratado na Rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168468)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim lNumNotaUltima As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Serie_Validate

    'Verifica se a Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
        
    'Verifica se é uma Serie selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 122641
    
    'Se não encontrou na lista da ComboBox
    If lErro <> SUCESSO Then
        
        'Traz os números default
        lErro = Traz_Numeros_Default()
        If lErro <> SUCESSO Then gError 122642
    
    End If

    Exit Sub
    
Erro_Serie_Validate:

    Cancel = True


    Select Case gErr
    
        Case 122641, 122642
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168469)
    
    End Select
    
    Exit Sub

End Sub

Private Function Traz_Numeros_Default() As Long

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim objNFiscal As New ClassNFiscal
Dim objModDocFis As New ClassModelosDocFiscais

On Error GoTo Erro_Traz_Numeros_Default

    If Serie.ListIndex = -1 Then Exit Function

    objSerie.sSerie = Serie.Text

    'Tenta ler a série no BD
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then gError 122658

    If lErro = 22202 Then gError 122659

    'Coloca número default de NFiscalInicial na tela
    If objSerie.lProxNumNFiscalImpressa > 0 Then
        NFiscalInicial.Text = objSerie.lProxNumNFiscalImpressa
    Else
        NFiscalInicial.Text = ""
    End If
    objNFiscal.sSerie = objSerie.sSerie
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    lErro = CF("NFiscal_FilialEmpresa_Customiza", objNFiscal)
    If lErro <> SUCESSO Then gError 126939

    'Le a Ultima Nota Cadastrada
    lErro = CF("NFiscal_Le_UltimaCadastrada", objNFiscal)
    If lErro <> SUCESSO And lErro <> 60431 Then gError 122660

    'Coloca nº default de NF final na tela
    If lErro = 60431 Then
        NFiscalFinal.Text = ""
    Else
        NFiscalFinal.Text = objNFiscal.lNumNotaFiscal
    End If
    
    objModDocFis.iCodigo = objSerie.iModDocFis
    
    lErro = CF("ModelosDocFiscais_Le", objModDocFis)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    DescMod.Caption = objModDocFis.sDescricao

    Traz_Numeros_Default = SUCESSO

    Exit Function

Erro_Traz_Numeros_Default:

    Traz_Numeros_Default = gErr

    Select Case gErr

        Case 122658, 122660, 126939
            Serie.SetFocus

        Case 122659
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)
            Serie.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168470)

    End Select

    Exit Function

End Function

Private Sub Timer1_Timer()

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim lNumNotasImprimindo As Long

On Error GoTo Erro_Timer1_Timer

    objSerie.sSerie = Serie.Text
    objSerie.iFilialEmpresa = giFilialEmpresa

    lErro = CF("Serie_FilialEmpresa_Customiza", objSerie)
    If lErro <> SUCESSO Then gError 126940

    lNumNotasImprimindo = CLng(NFiscalFinal.Text) - CLng(NFiscalInicial.Text) + 1
    glNumAcessosTimer = glNumAcessosTimer + 1
    
    'Se não ultrapassou o tempo máximo de impressão
    If (glNumAcessosTimer * INTERVALO_MONITORAMENTO_IMPRESSAO_NF) <= (TEMPO_MAX_IMPRESSAO_UMA_NF * lNumNotasImprimindo) Then
    
        'Verifica se já Terminou a Impressão
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And 22202 Then gError 122643
       
        'Se a Série não está cadastrada --> ERRO
        If lErro = 22202 Then gError 122644
            
        'Se terminou a Impressão
        If objSerie.iImprimindo = RELATORIO_NF_NAO_IMPRIMINDO Then
                   
           Timer1.Interval = 0
           
            'Chama a Tela de Controle de Impressão de Notas Fiscais
            Call Chama_Tela("RelOpControleImprNF", objSerie, CLng(NFiscalInicial.Text), CLng(NFiscalFinal.Text))
          
            Unload Me
    
        End If
    
    Else
    
        'zera o timer
        Timer1.Interval = 0
        
        'Coloca iImprimindo = 0
        lErro = CF("Serie_Altera_Nao_Imprimindo", objSerie)
        If lErro <> SUCESSO And lErro <> 61025 Then gError 122646
                
        'Não encontrou a Série
        If lErro = 61025 Then gError 122645
        
        'Chama a Tela de Controle de Impressão das Notas Fiscais
        Call Chama_Tela("RelOpControleImprNF", objSerie, CLng(NFiscalInicial.Text), CLng(NFiscalFinal.Text))
        
        Unload Me
    
    End If
    
    Exit Sub
    
Erro_Timer1_Timer:
      
    Select Case gErr
    
        Case 122646, 126940
        
        Case 122643 'Tratado na Rotina chamada
        
        Case 122644, 122645
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168471)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub TipoFormulario_Click()

Dim lErro As Long

On Error GoTo Erro_TipoFormulario_Click

    If TipoFormulario.ListIndex <> -1 Then
    
'        lErro = Carrega_Serie(TipoFormulario.ItemData(TipoFormulario.ListIndex))
'        If lErro <> SUCESSO Then gError 122647

    End If
    
    Exit Sub
    
Erro_TipoFormulario_Click:

    Select Case gErr
        
        Case 122647
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168472)
    
    End Select
    
    Exit Sub

End Sub

'William
'copiada da tela RelOpNotasFiscais
Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero

    If Len(Trim(sNumero)) > 0 Then

        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then gError 122661

        If CLng(sNumero) < 0 Then gError 122662

    End If

    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = gErr

    Select Case gErr

        Case 122661

        Case 122662
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sNumero)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168473)

    End Select

    Exit Function

End Function




'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_NF
    Set Form_Load_Ocx = Me
    Caption = "Emissão das Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpEmissaoNFiscal"
    
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
        
        If Me.ActiveControl Is Serie Then
            Call LabelSerie_Click
        End If
    
    End If

End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LabelSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerie, Source, X, Y)
End Sub

Private Sub LabelSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerie, Button, Shift, X, Y)
End Sub






