VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpCoinfoLog 
   Appearance      =   0  'Flat
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   6660
   Begin VB.Frame Frame4 
      Caption         =   "Importação"
      Height          =   2700
      Left            =   90
      TabIndex        =   22
      Top             =   3705
      Width           =   6360
      Begin VB.ListBox PrimeiraImportList 
         Columns         =   3
         Height          =   960
         ItemData        =   "RelOpCoinfoLog.ctx":0000
         Left            =   315
         List            =   "RelOpCoinfoLog.ctx":0010
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   2745
         Visible         =   0   'False
         Width           =   5805
      End
      Begin VB.CheckBox PrimeiraImport 
         Caption         =   "Primeira importação (Somente Suporte - o uso indevido causa danos ao sistema)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   300
         TabIndex        =   25
         Top             =   2685
         Visible         =   0   'False
         Width           =   5940
      End
      Begin VB.ListBox ImportarList 
         Columns         =   3
         Height          =   2085
         ItemData        =   "RelOpCoinfoLog.ctx":00E8
         Left            =   330
         List            =   "RelOpCoinfoLog.ctx":0104
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   510
         Visible         =   0   'False
         Width           =   5790
      End
      Begin VB.CheckBox Importar 
         Caption         =   "Importar\Exportar e Atualizar dados antes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   330
         TabIndex        =   23
         Top             =   225
         Width           =   4290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log de Atualização"
      Height          =   2580
      Left            =   90
      TabIndex        =   8
      Top             =   975
      Width           =   4500
      Begin VB.CheckBox SoErros 
         Caption         =   "Exibe somente os erros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   135
         TabIndex        =   21
         Top             =   2040
         Width           =   2790
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tabelas"
         Height          =   825
         Left            =   105
         TabIndex        =   16
         Top             =   330
         Width           =   4140
         Begin VB.ComboBox Tabelas 
            Height          =   315
            ItemData        =   "RelOpCoinfoLog.ctx":01FF
            Left            =   1740
            List            =   "RelOpCoinfoLog.ctx":0209
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   255
            Width           =   2310
         End
         Begin VB.CheckBox Todas 
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
            Height          =   252
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   30
         End
         Begin VB.Label Label4 
            Caption         =   "Tabela:"
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
            Left            =   1080
            TabIndex        =   19
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data de Atualização"
         Height          =   780
         Left            =   105
         TabIndex        =   9
         Top             =   1275
         Width           =   4140
         Begin MSComCtl2.UpDown UpDownDtIni 
            Height          =   315
            Left            =   1545
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicial 
            Height          =   300
            Left            =   570
            TabIndex        =   11
            Top             =   315
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDtFim 
            Height          =   315
            Left            =   3450
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFinal 
            Height          =   300
            Left            =   2490
            TabIndex        =   13
            Top             =   315
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2115
            TabIndex        =   15
            Top             =   360
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   14
            Top             =   330
            Width           =   345
         End
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCoinfoLog.ctx":0225
      Left            =   915
      List            =   "RelOpCoinfoLog.ctx":0227
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   2505
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
      Left            =   4680
      Picture         =   "RelOpCoinfoLog.ctx":0229
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpCoinfoLog.ctx":032B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpCoinfoLog.ctx":04A9
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpCoinfoLog.ctx":09DB
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpCoinfoLog.ctx":0B65
         Style           =   1  'Graphical
         TabIndex        =   3
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
      Left            =   240
      TabIndex        =   7
      Top             =   420
      Width           =   615
   End
End
Attribute VB_Name = "RelOpCoinfoLog"
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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Todas.Value = vbChecked
    Tabelas.Enabled = False

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182607)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 182608
      
    lErro = objRelOpcoes.ObterParametro("NTABELA", sParam)
    If lErro Then gError 182609
    
    If StrParaInt(sParam) = 0 Then
        Todas.Value = vbChecked
        Tabelas.ListIndex = -1
    Else
        Todas.Value = vbUnchecked
        Call Combo_Seleciona_ItemData(Tabelas, StrParaInt(sParam))
    End If
         
    lErro = objRelOpcoes.ObterParametro("NIMPORTAR", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        Importar.Value = vbChecked
    Else
        Importar.Value = vbUnchecked
    End If
    Call Importar_Click
    
    lErro = objRelOpcoes.ObterParametro("NPRIMEIRAIMPORT", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        PrimeiraImport.Value = vbChecked
    Else
        PrimeiraImport.Value = vbUnchecked
    End If
    Call PrimeiraImport_Click
    
    lErro = objRelOpcoes.ObterParametro("NGERCOMIRET", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        PrimeiraImportList.Selected(2) = True
    Else
        PrimeiraImportList.Selected(2) = False
    End If

    lErro = objRelOpcoes.ObterParametro("NATUCLIRET", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        PrimeiraImportList.Selected(1) = True
    Else
        PrimeiraImportList.Selected(1) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NIMPNVLRET", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        PrimeiraImportList.Selected(0) = True
    Else
        PrimeiraImportList.Selected(0) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NCONSFATSIGAV", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        PrimeiraImportList.Selected(3) = True
    Else
        PrimeiraImportList.Selected(3) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NCONTABVOU", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(0) = True
    Else
        ImportarList.Selected(0) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NGEROVER", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(1) = True
    Else
        ImportarList.Selected(1) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NGERMOVEST", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(2) = True
    Else
        ImportarList.Selected(2) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NIMPARQNOVOS", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(3) = True
    Else
        ImportarList.Selected(3) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NCONTABFAT", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(4) = True
    Else
        ImportarList.Selected(4) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NCONTABNF", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(5) = True
    Else
        ImportarList.Selected(5) = False
    End If
    
    
    lErro = objRelOpcoes.ObterParametro("NSOEXP", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(6) = True
    Else
        ImportarList.Selected(6) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NGERARBOL", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        ImportarList.Selected(7) = True
    Else
        ImportarList.Selected(7) = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NERROS", sParam)
    If lErro Then gError 182610
    
    If StrParaInt(sParam) = MARCADO Then
        SoErros.Value = vbChecked
    Else
        SoErros.Value = vbUnchecked
    End If
           
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 182611

    Call DateParaMasked(DataInicial, StrParaDate(sParam))
 
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 182612

    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 182608 To 182612

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182613)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
  
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 182614
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 182615
  
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
                
        Case 182614
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 182615
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182616)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros(iTabela As Integer, iImportar As Integer, iSoErros As Integer, iPrimeiraImport As Integer, ByVal objFiltro As ClassFiltroImportCoinfo) As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
             
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 182617
    
    End If
    
    If Todas.Value = vbChecked Then
        iTabela = 0
    Else
        iTabela = Tabelas.ItemData(Tabelas.ListIndex)
    End If
    
    If Importar.Value = vbChecked Then
        iImportar = MARCADO
        objFiltro.iContabilizarVouchers = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(0))
        objFiltro.iGerarOver = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(1))
        objFiltro.iGerarMovEst = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(2))
        objFiltro.iImportArqsNovos = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(3))
        objFiltro.iContabilizarFaturas = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(4))
        objFiltro.iContabilizarNFs = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(5))
        objFiltro.iSoArqExport = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(6))
        objFiltro.iGerarBol = Conv_Boolean_em_Integer_Marcado(ImportarList.Selected(7))
    Else
        iImportar = DESMARCADO
    End If
    
    If PrimeiraImport.Value = vbChecked Then
        iPrimeiraImport = MARCADO
        objFiltro.iImportarNVLRetroativo = Conv_Boolean_em_Integer_Marcado(PrimeiraImportList.Selected(0))
        objFiltro.iAtualizarClientesRetroativo = Conv_Boolean_em_Integer_Marcado(PrimeiraImportList.Selected(1))
        objFiltro.iGerarComissaoRetroativo = Conv_Boolean_em_Integer_Marcado(PrimeiraImportList.Selected(2))
        objFiltro.iConsiderarFatSigav = Conv_Boolean_em_Integer_Marcado(PrimeiraImportList.Selected(3))
    Else
        iPrimeiraImport = DESMARCADO
    End If
    
    If SoErros.Value = vbChecked Then
        iSoErros = MARCADO
    Else
        iSoErros = DESMARCADO
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 182617
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182618)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 182619
    
    Todas.Value = vbChecked
    Tabelas.Enabled = False
    Importar.Value = vbUnchecked
    SoErros.Value = vbUnchecked
    
    If ComboOpcoes.Visible Then
        ComboOpcoes.Text = ""
        ComboOpcoes.SetFocus
    End If
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 182619
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182620)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iTabela As Integer
Dim iImportar As Integer
Dim iSoErros As Integer
Dim sNomeArqParam As String
Dim iPrimeiraImport As Integer
Dim objFiltro As New ClassFiltroImportCoinfo

On Error GoTo Erro_PreencherRelOp

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Formata_E_Critica_Parametros(iTabela, iImportar, iSoErros, iPrimeiraImport, objFiltro)
    If lErro <> SUCESSO Then gError 182621
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 182622
    
    lErro = objRelOpcoes.IncluirParametro("NTABELA", CStr(iTabela))
    If lErro <> AD_BOOL_TRUE Then gError 182623

    lErro = objRelOpcoes.IncluirParametro("NIMPORTAR", CStr(iImportar))
    If lErro <> AD_BOOL_TRUE Then gError 182624
   
    lErro = objRelOpcoes.IncluirParametro("NPRIMEIRAIMPORT", CStr(iPrimeiraImport))
    If lErro <> AD_BOOL_TRUE Then gError 182624
   
    lErro = objRelOpcoes.IncluirParametro("NERROS", CStr(iSoErros))
    If lErro <> AD_BOOL_TRUE Then gError 182624
   
    lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(StrParaDate(DataInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 182625

    lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(StrParaDate(DataFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 182626
    
    lErro = objRelOpcoes.IncluirParametro("NGERCOMIRET", Conv_Boolean_em_String_Marcado(PrimeiraImportList.Selected(2)))
    If lErro <> AD_BOOL_TRUE Then gError 182624

    lErro = objRelOpcoes.IncluirParametro("NATUCLIRET", Conv_Boolean_em_String_Marcado(PrimeiraImportList.Selected(1)))
    If lErro <> AD_BOOL_TRUE Then gError 182624

    lErro = objRelOpcoes.IncluirParametro("NIMPNVLRET", Conv_Boolean_em_String_Marcado(PrimeiraImportList.Selected(0)))
    If lErro <> AD_BOOL_TRUE Then gError 182624

    lErro = objRelOpcoes.IncluirParametro("NCONSFATSIGAV", Conv_Boolean_em_String_Marcado(PrimeiraImportList.Selected(3)))
    If lErro <> AD_BOOL_TRUE Then gError 182624

    lErro = objRelOpcoes.IncluirParametro("NCONTABVOU", Conv_Boolean_em_String_Marcado(ImportarList.Selected(0)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = objRelOpcoes.IncluirParametro("NGEROVER", Conv_Boolean_em_String_Marcado(ImportarList.Selected(1)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = objRelOpcoes.IncluirParametro("NGERMOVEST", Conv_Boolean_em_String_Marcado(ImportarList.Selected(2)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = objRelOpcoes.IncluirParametro("NIMPARQNOVOS", Conv_Boolean_em_String_Marcado(ImportarList.Selected(3)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = objRelOpcoes.IncluirParametro("NCONTABFAT", Conv_Boolean_em_String_Marcado(ImportarList.Selected(4)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = objRelOpcoes.IncluirParametro("NCONTABNF", Conv_Boolean_em_String_Marcado(ImportarList.Selected(5)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = objRelOpcoes.IncluirParametro("NSOEXP", Conv_Boolean_em_String_Marcado(ImportarList.Selected(6)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = objRelOpcoes.IncluirParametro("NGERARBOL", Conv_Boolean_em_String_Marcado(ImportarList.Selected(7)))
    If lErro <> AD_BOOL_TRUE Then gError 182624
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 182627
    
    If bExecutando Then
    
        If iImportar = MARCADO Then
        
            lErro = Sistema_Preparar_Batch(sNomeArqParam)
            If lErro <> SUCESSO Then gError 182703
        
            lErro = CF("Rotina_Importa_Dados_Coinfo", sNomeArqParam, objFiltro)
            If lErro <> SUCESSO Then gError 182628
        
        End If
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    GL_objMDIForm.MousePointer = vbDefault

    PreencherRelOp = gErr

    Select Case gErr

        Case 182621 To 182628, 182703

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182629)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 182630

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 182631

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 182632
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 182630
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 182631, 182632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182633)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 182634

    If Importar.Value = vbUnchecked Then
        Call gobjRelatorio.Executar_Prossegue2(Me)
    End If

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 182634

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182635)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 182636

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 182637

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 182638

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 182639
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 182636
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 182637 To 182639

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182640)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182641)

    End Select

    Exit Function

End Function

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 182642

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 182642

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182643)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 182644

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 182644

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182645)

    End Select

    Exit Sub

End Sub

Private Sub Todas_Click()

    If Todas.Value = vbChecked Then
        Tabelas.ListIndex = -1
        Tabelas.Enabled = False
    Else
        Tabelas.Enabled = True
    End If

End Sub

Private Sub Todas_Change()

    If Todas.Value = vbChecked Then
        Tabelas.ListIndex = -1
        Tabelas.Enabled = False
    Else
        Tabelas.Enabled = True
    End If

End Sub

Private Sub UpDownDtIni_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtIni_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 182646

    Exit Sub

Erro_UpDownDtIni_DownClick:

    Select Case gErr

        Case 182646
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182647)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDtIni_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtIni_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 182648

    Exit Sub

Erro_UpDownDtIni_UpClick:

    Select Case gErr

        Case 182648
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182649)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDtFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtFim_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 182650

    Exit Sub

Erro_UpDownDtFim_DownClick:

    Select Case gErr

        Case 182650
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182651)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDtFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtFim_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 182652

    Exit Sub

Erro_UpDownDtFim_UpClick:

    Select Case gErr

        Case 182652
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182653)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_NOTAS_REC
    Set Form_Load_Ocx = Me
    Caption = "Log de Atualização de Dados da Coinfo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpCoinfoLog"
    
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
        
    
    End If

End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Importar_Click()

Dim iIndice As Integer

    If Importar.Value = vbChecked Then
        ImportarList.Visible = True
        PrimeiraImport.Visible = True
    Else
        ImportarList.Visible = False
        PrimeiraImport.Visible = False
        PrimeiraImportList.Visible = False
        PrimeiraImport.Value = vbUnchecked

        For iIndice = 0 To ImportarList.ListCount - 1
            ImportarList.Selected(iIndice) = False
        Next
        For iIndice = 0 To PrimeiraImportList.ListCount - 1
            PrimeiraImportList.Selected(iIndice) = False
        Next
    End If
End Sub

Private Sub PrimeiraImport_Click()

Dim iIndice As Integer

    If PrimeiraImport.Value = vbChecked Then
        PrimeiraImportList.Visible = True
    Else
        PrimeiraImportList.Visible = False
        
        For iIndice = 0 To PrimeiraImportList.ListCount - 1
            PrimeiraImportList.Selected(iIndice) = False
        Next
        
    End If
    
End Sub

Private Function Conv_Boolean_em_String_Marcado(ByVal bFlag As Boolean) As String
    If bFlag Then
        Conv_Boolean_em_String_Marcado = CStr(MARCADO)
    Else
        Conv_Boolean_em_String_Marcado = CStr(DESMARCADO)
    End If
End Function

Private Function Conv_Boolean_em_Integer_Marcado(ByVal bFlag As Boolean) As Integer
    If bFlag Then
        Conv_Boolean_em_Integer_Marcado = MARCADO
    Else
        Conv_Boolean_em_Integer_Marcado = DESMARCADO
    End If
End Function
