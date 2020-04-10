VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpEmissaoNFiscalOcx 
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2910
   ScaleMode       =   0  'User
   ScaleWidth      =   4665
   Begin VB.Frame FrameNF 
      Caption         =   "Nota Fiscal"
      Height          =   1725
      Left            =   240
      TabIndex        =   4
      Top             =   990
      Width           =   4215
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   825
         Width           =   975
      End
      Begin VB.ComboBox TipoFormulario 
         Height          =   315
         ItemData        =   "RelOpEmissaoNFiscalOcx.ctx":0000
         Left            =   1680
         List            =   "RelOpEmissaoNFiscalOcx.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2385
      End
      Begin MSMask.MaskEdBox NFiscalInicial 
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   1290
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFiscalFinal 
         Height          =   300
         Left            =   3135
         TabIndex        =   8
         Top             =   1290
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         Left            =   1125
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   885
         Width           =   510
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
         Left            =   1305
         TabIndex        =   11
         Top             =   1343
         Width           =   300
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
         Left            =   2760
         TabIndex        =   10
         Top             =   1343
         Width           =   360
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
         Top             =   420
         Width           =   1380
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2370
      Top             =   180
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3300
      ScaleHeight     =   495
      ScaleMode       =   0  'User
      ScaleWidth      =   1065
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   1125
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpEmissaoNFiscalOcx.ctx":0047
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpEmissaoNFiscalOcx.ctx":0579
         Style           =   1  'Graphical
         TabIndex        =   3
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
      Height          =   555
      Left            =   960
      Picture         =   "RelOpEmissaoNFiscalOcx.ctx":06F7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   1845
   End
End
Attribute VB_Name = "RelOpEmissaoNFiscalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variavel global
Private glNumAcessosTimer As Long

Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento
        
    TipoFormulario.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 38153 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168474)

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

    If Not (gobjRelatorio Is Nothing) Then Error 64131
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    If Not IsMissing(vParam) Then
        
        'Se passou a Série
        If vParam <> "" Then
                    
            objSerie.iFilialEmpresa = giFilialEmpresa
            
            lErro = CF("Serie_FilialEmpresa_Customiza", objSerie)
            If lErro <> SUCESSO Then Error 20861
            
            objSerie.sSerie = CStr(vParam)
            
            'Lê a Serie no BD
            lErro = CF("Serie_Le", objSerie)
            If lErro <> SUCESSO And lErro <> 22202 Then Error 64128
            
            'Se não encontrou Erro
            If lErro = 22202 Then Error 64129
                       
            bEncontrou = False
            
            For iIndice = 0 To TipoFormulario.ListCount - 1
            
                If TipoFormulario.ItemData(iIndice) = objSerie.iTipoFormulario Then
                    
                    TipoFormulario.ListIndex = iIndice
                    bEncontrou = True
                    Exit For
                End If
                
            Next
            
            If bEncontrou = False Then Error 64130
        
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

    Trata_Parametros = Err

    Select Case Err
        
        Case 20861, 64128
        
        Case 64131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case 64129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)
        
        Case 64130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULARIO_IMCOMPATIVEL", Err, objSerie.sSerie)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168475)

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
      
    'Verifica se a Série está cadastrada
    If Len(Trim(Serie.Text)) = 0 Then Error 60375
    
    'Verifica se a Nota Inicial está Cadastrada
    If Len(Trim(NFiscalInicial.Text)) = 0 Then Error 60376
    
    'Verifica se a Nota Final está Cadastrada
    If Len(Trim(NFiscalFinal.Text)) = 0 Then Error 60377
      
    'Verifica se o numero da Nota Fiscal inicial é maior que o da final
    If Len(Trim(NFiscalInicial.ClipText)) > 0 And Len(Trim(NFiscalFinal.ClipText)) > 0 Then
    
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then Error 38159
    
    End If
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = Err

    Select Case Err

        Case 38159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)
        
        Case 60375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)
        
        Case 60376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DE_NAO_PREENCHIDO", Err)
        
        Case 60377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_NAO_PREENCHIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168476)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47122
    
    TipoFormulario.ListIndex = 0
    
    Serie.ListIndex = -1
    Serie.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47122 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168477)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then Error 38162
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 38163
    
    lErro = objRelOpcoes.IncluirParametro("NBORDERO", "0")
    If lErro <> AD_BOOL_TRUE Then Error 38164
    
    lErro = objRelOpcoes.IncluirParametro("NNFISCALINIC", NFiscalInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38164

    lErro = objRelOpcoes.IncluirParametro("NNFISCALFIM", NFiscalFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38165
   
    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38166
   
'    gobjRelatorio.sNomeTsk = "NFiscal"
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 38162 To 38166 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168478)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim vbMsgRes As VbMsgBoxResult
Dim lFaixaFinal As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 60380
    
    objSerie.sSerie = Serie.Text
    objSerie.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("Serie_FilialEmpresa_Customiza", objSerie)
    If lErro <> SUCESSO Then Error 20862
    
    'Faz o Lock na tabela de Série para a Impressão da Nota Fiscal
    lErro = CF("Serie_Lock_ImpressaoNFiscal", objSerie)
    If lErro <> SUCESSO And lErro <> 60387 Then Error 60381
    
    'Se a Série não está Cadastrada --> ERRO
    If lErro = 60387 Then Error 60384
    
    'Dá Mensagem ao usuário caso seja Reimpressão
    If CLng(NFiscalInicial.Text) < objSerie.lProxNumNFiscalImpressa Then
        
        'Verifica se a Faixa Final tambem não é menor que a que está no BD
        If CLng(NFiscalFinal.Text) < objSerie.lProxNumNFiscalImpressa Then
            lFaixaFinal = CLng(NFiscalFinal.Text)
        Else
            lFaixaFinal = objSerie.lProxNumNFiscalImpressa - 1
        End If
        
        'Avisa Reimpressao de Nota e Pede Confirmação
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NFISCAL_REIMPRESSA", CLng(NFiscalInicial.Text), lFaixaFinal)
        If vbMsgRes = vbNo Then Error 60422
    
    End If
    
    'Altera campo Imprimindo = RELATORIO_NFISCAL_IMPRIMINDO
    lErro = CF("Serie_Altera_Imprimindo", objSerie)
    If lErro <> SUCESSO And lErro <> 61008 Then Error 61010
    
    Me.Enabled = False
    
    If TipoFormulario.ItemData(TipoFormulario.ListIndex) = TIPO_FORMULARIO_NFISCAL Then
        gobjRelatorio.sNomeTsk = "nfiscal"
    ElseIf TipoFormulario.ItemData(TipoFormulario.ListIndex) = TIPO_FORMULARIO_NFISCAL_SERVICO Then
        gobjRelatorio.sNomeTsk = "nfserv"
    ElseIf TipoFormulario.ItemData(TipoFormulario.ListIndex) = TIPO_FORMULARIO_NFISCAL_FRETE Then
        gobjRelatorio.sNomeTsk = "nffrete"
    End If
    
    lErro = gobjRelatorio.Executar_Prossegue
    If lErro <> SUCESSO And lErro <> 7072 Then Error 61435
    
    'Cancelou o relatório
    If lErro = 7072 Then Error 61436
    
    Timer1.Interval = INTERVALO_MONITORAMENTO_IMPRESSAO_NF
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 20862, 60380, 60381, 60422, 61010, 61435 'Tratado na Rotina chamada
        
        Case 60384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)
        
        Case 61436
            Unload Me
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168479)

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

On Error GoTo Erro_NFiscalInicial_Validate
    
    lErro = Critica_Numero(NFiscalInicial.Text)
    If lErro <> SUCESSO Then Error 38174
              
    Exit Sub

Erro_NFiscalInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 38174
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168480)
            
    End Select
    
    Exit Sub

End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalFinal_Validate
     
    lErro = Critica_Numero(NFiscalFinal.Text)
    If lErro <> SUCESSO Then Error 38175
        
    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 38175
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168481)
            
    End Select
    
    Exit Sub

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero
         
    If Len(Trim(sNumero)) > 0 Then
        
        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then Error 38176
 
    End If
 
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = Err

    Select Case Err
                  
        Case 38176 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168482)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie(iTipoFormulario As Integer) As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le_TipoFormulario", colSerie, iTipoFormulario)
    If lErro <> SUCESSO Then Error 38178
    
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

    Carrega_Serie = Err
    
    Select Case Err
    
        Case 38178 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168483)
            
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
    If lErro <> SUCESSO Then Error 60464
        
    Exit Sub
    
Erro_Serie_Click:

    Select Case Err
    
        Case 60464 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168484)
    
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
    If lErro <> SUCESSO And lErro <> 12253 Then Error 38179
    
    'Se não encontrou na lista da ComboBox
    If lErro <> SUCESSO Then
        
        'Traz os números default
        lErro = Traz_Numeros_Default()
        If lErro <> SUCESSO Then Error 60464
    
    End If
    
    Exit Sub
    
Erro_Serie_Validate:

    Cancel = True


    Select Case Err
    
        Case 38179, 60464 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168485)
    
    End Select
    
    Exit Sub

End Sub

Private Function Traz_Numeros_Default() As Long

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Traz_Numeros_Default

    If Serie.ListIndex = -1 Then Exit Function
    
    objSerie.sSerie = Serie.List(Serie.ListIndex)

    'Tenta ler a série no BD
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then Error 60366
    
    If lErro = 22202 Then Error 60367
        
    'Coloca número default de NFiscalInicial na tela
    If objSerie.lProxNumNFiscalImpressa > 0 Then NFiscalInicial.Text = objSerie.lProxNumNFiscalImpressa
    
    objNFiscal.sSerie = objSerie.sSerie
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("NFiscal_FilialEmpresa_Customiza", objNFiscal)
    If lErro <> SUCESSO Then Error 20863
    
    'Le a Ultima Nota Cadastrada
    lErro = CF("NFiscal_Le_UltimaCadastrada", objNFiscal)
    If lErro <> SUCESSO And lErro <> 60431 Then Error 60374
    
    'Coloca nº default de NF final na tela
    If lErro = 60431 Then
        NFiscalFinal.Text = ""
    Else
        If objNFiscal.lNumNotaFiscal > 0 Then NFiscalFinal.Text = objNFiscal.lNumNotaFiscal
    End If

    Traz_Numeros_Default = SUCESSO
    
    Exit Function
    
Erro_Traz_Numeros_Default:
    
    Traz_Numeros_Default = Err
    
    Select Case Err
    
        Case 20863, 38179, 60366, 60374 'Tratado na Rotina chamada
            Serie.SetFocus
       
        Case 60367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)
            Serie.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168486)
    
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
    If lErro <> SUCESSO Then Error 20864
    
    lNumNotasImprimindo = CLng(NFiscalFinal.Text) - CLng(NFiscalInicial.Text) + 1
    glNumAcessosTimer = glNumAcessosTimer + 1
    
    'Se não ultrapassou o tempo máximo de impressão
    If (glNumAcessosTimer * INTERVALO_MONITORAMENTO_IMPRESSAO_NF) <= (TEMPO_MAX_IMPRESSAO_UMA_NF * lNumNotasImprimindo) Then
    
        'Verifica se já Terminou a Impressão
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And 22202 Then Error 60382
    
        'Se não existe a série no BD
        If lErro = 22202 Then Error 60393
    
        'Se terminou a impressão
        If objSerie.iImprimindo = RELATORIO_NF_NAO_IMPRIMINDO Then
        
            Timer1.Interval = 0
        
            'Chama a Tela de Controle de Impressão das Notas Fiscais
            Call Chama_Tela("RelOpControleImprNF", objSerie, CLng(NFiscalInicial.Text), CLng(NFiscalFinal.Text))
        
            Unload Me
        
        End If
    
    Else
        'zera o timer
        Timer1.Interval = 0
        
        'Coloca iImprimindo = 0
        lErro = CF("Serie_Altera_Nao_Imprimindo", objSerie)
        If lErro <> SUCESSO And lErro <> 61025 Then Error 61030
        
        If lErro = 61025 Then Error 61038
        
        'Chama a Tela de Controle de Impressão das Notas Fiscais
        Call Chama_Tela("RelOpControleImprNF", objSerie, CLng(NFiscalInicial.Text), CLng(NFiscalFinal.Text))
        
        Unload Me
        
    End If
    
    Exit Sub
    
Erro_Timer1_Timer:
      
    'zera o timer
    Timer1.Interval = 0
    
    Select Case Err
        
        Case 20864, 60382, 61030 'Tratado na Rotina chamada
                
        Case 60393, 61038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168487)
    
    End Select
    
    'Chama a Tela de Controle de Impressão das Notas Fiscais
    Call Chama_Tela("RelOpControleImprNF", objSerie, CLng(NFiscalInicial.Text), CLng(NFiscalFinal.Text))
    Unload Me
    
    Exit Sub
    
End Sub

Private Sub TipoFormulario_Click()

Dim lErro As Long

On Error GoTo Erro_TipoFormulario_Click

    If (TipoFormulario.ListIndex <> -1) Then
    
        lErro = Carrega_Serie(TipoFormulario.ItemData(TipoFormulario.ListIndex))
        If lErro <> SUCESSO Then Error 64060

    End If
    
    Exit Sub
    
Erro_TipoFormulario_Click:

    Select Case Err
        
        Case 64060
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168488)
    
    End Select
    
    Exit Sub

End Sub

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

Private Sub LabelSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerie, Source, X, Y)
End Sub

Private Sub LabelSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerie, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

