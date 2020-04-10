VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpResVendasUFTA 
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ScaleHeight     =   4005
   ScaleWidth      =   6015
   Begin VB.Frame Frame2 
      Caption         =   "Estados"
      Height          =   945
      Left            =   150
      TabIndex        =   28
      Top             =   2895
      Width           =   5640
      Begin VB.ComboBox EstadoInicial 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   360
         Width           =   2220
      End
      Begin VB.ComboBox EstadoFinal 
         Height          =   315
         Left            =   3270
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label LabelEstadosAte 
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
         Left            =   2865
         TabIndex        =   32
         Top             =   420
         Width           =   360
      End
      Begin VB.Label LabelEstadosDe 
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
         Left            =   135
         TabIndex        =   31
         Top             =   405
         Width           =   315
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
      Left            =   3840
      Picture         =   "RelOpResVendasUFTA.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   930
      Width           =   1815
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   1215
      Left            =   165
      TabIndex        =   19
      Top             =   660
      Width           =   2745
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1875
         TabIndex        =   20
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
         Left            =   915
         TabIndex        =   21
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
         Left            =   1875
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   900
         TabIndex        =   23
         Top             =   735
         Width           =   960
         _ExtentX        =   1693
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
         Height          =   240
         Left            =   525
         TabIndex        =   25
         Top             =   315
         Width           =   345
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
         Left            =   495
         TabIndex        =   24
         Top             =   795
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpResVendasUFTA.ctx":0102
      Left            =   870
      List            =   "RelOpResVendasUFTA.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   270
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   180
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpResVendasUFTA.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpResVendasUFTA.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpResVendasUFTA.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpResVendasUFTA.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Nota Fiscal"
      Height          =   825
      Left            =   165
      TabIndex        =   6
      Top             =   1980
      Width           =   5505
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   690
         TabIndex        =   7
         Top             =   300
         Width           =   765
      End
      Begin MSMask.MaskEdBox NFiscalInicial 
         Height          =   300
         Left            =   2460
         TabIndex        =   8
         Top             =   330
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NFiscalFinal 
         Height          =   300
         Left            =   4170
         TabIndex        =   9
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         Height          =   195
         Left            =   3750
         TabIndex        =   12
         Top             =   375
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2070
         TabIndex        =   11
         Top             =   375
         Width           =   315
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.CheckBox CheckAnalitico 
      Caption         =   "Exibe Nota a Nota"
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
      Left            =   3165
      TabIndex        =   5
      Top             =   4530
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame FramePedido 
      Caption         =   "Natureza de Operação"
      Height          =   1215
      Left            =   75
      TabIndex        =   0
      Top             =   4050
      Visible         =   0   'False
      Width           =   2745
      Begin MSMask.MaskEdBox NaturezaInicial 
         Height          =   300
         Left            =   1020
         TabIndex        =   1
         Top             =   270
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NaturezaFinal 
         Height          =   300
         Left            =   1020
         TabIndex        =   2
         Top             =   750
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelNatFinal 
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
         Left            =   585
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   4
         Top             =   795
         Width           =   360
      End
      Begin VB.Label LabelNatInicial 
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
         Left            =   675
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   3
         Top             =   300
         Width           =   315
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
      Left            =   165
      TabIndex        =   27
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpResVendasUFTA"
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
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giNaturezaInicial As Integer
Dim giEstadoInicial As Integer
Dim giEstadoFinal As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoEstado As New Collection
Dim objEstados As ClassEstado

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento
    Set objEventoNatureza = New AdmEvento
            
    'Carrega a combo série
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then Error 38107
        
    lErro = CF("Estados_Le_Todos", colCodigoEstado)
    If lErro <> SUCESSO Then Error 38107

    'preenche cada ComboBox País com os objetos da colecao colCodigoDescricao
    For Each objEstados In colCodigoEstado

        EstadoInicial.AddItem CStr(objEstados.sSigla) & SEPARADOR & objEstados.sNome
        EstadoFinal.AddItem CStr(objEstados.sSigla) & SEPARADOR & objEstados.sNome
        
    Next
    
    'define Exibir Titulo a Titulo como Padrao
    CheckAnalitico.Value = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 38107

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173151)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 38108
    
    'pega Estado inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TESTADOINIC", sParam)
    If lErro Then Error 38108
    
    EstadoInicial.Text = sParam
    
    'pega  Estado final e exibe
    lErro = objRelOpcoes.ObterParametro("TESTADOFIM", sParam)
    If lErro Then Error 38108
    
    EstadoFinal.Text = sParam
              
    'pega Nota Fiscal inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNFISCALINIC", sParam)
    If lErro <> SUCESSO Then Error 38109
    
    NFiscalInicial.Text = sParam
         
    'pega Nota Fiscal final e exibe
    lErro = objRelOpcoes.ObterParametro("NNFISCALFIM", sParam)
    If lErro <> SUCESSO Then Error 38110
    
    NFiscalFinal.Text = sParam
           
    'pega série e exibe
    lErro = objRelOpcoes.ObterParametro("TSERIE", sParam)
    If lErro <> SUCESSO Then Error 38111

    Serie.Text = sParam
           
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 38113

    DataInicial.PromptInclude = False
    DataInicial.Text = sParam
    DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 38114

    DataFinal.PromptInclude = False
    DataFinal.Text = sParam
    DataFinal.PromptInclude = True
    
    lErro = objRelOpcoes.ObterParametro("NEXIBTIT", sParam)
    If lErro <> SUCESSO Then Error 38145
    
    CheckAnalitico.Value = CInt(sParam)
           
    'pega parâmetro Natureza Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNATUREZAINIC", sParam)
    If lErro Then Error 38146
    
    NaturezaInicial.Text = sParam
    
    'pega parâmetro Natureza Final e exibe
    lErro = objRelOpcoes.ObterParametro("NNATUREZAFIM", sParam)
    If lErro Then Error 38147
    
    NaturezaFinal.Text = sParam
           
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 38108, 38109, 38110, 38111, 38113, 38114, 38145, 38146, 38147

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173152)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29884
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 38105
 
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 38105
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173153)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoNatureza = Nothing
    Set objEventoSerie = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sEstados_I As String, sEstados_F As String

On Error GoTo Erro_Formata_E_Critica_Parametros
  
   'Data Inicial não pode ser maior que a Data Final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 38117
    End If
    
    'Verifica se o numero da Nota Fiscal inicial é maior que o da final
    If Len(Trim(NFiscalInicial.ClipText)) > 0 And Len(Trim(NFiscalFinal.ClipText)) > 0 Then
    
        If CLng(NFiscalInicial.Text) > CLng(NFiscalFinal.Text) Then Error 43216
    
    End If
    
    'Natureza inicial não pode ser maior que a Natureza final
    If Trim(NaturezaInicial.Text) <> "" And Trim(NaturezaFinal.Text) <> "" Then
    
         If CLng(NaturezaInicial.Text) > CLng(NaturezaFinal.Text) Then Error 38148
         
    End If
    
    If EstadoInicial.Text <> "" Then
        sEstados_I = CStr(SCodigo_Extrai(EstadoInicial.Text))
    Else
        sEstados_I = ""
    End If
    
    If EstadoFinal.Text <> "" Then
        sEstados_F = CStr(SCodigo_Extrai(EstadoFinal.Text))
    Else
        sEstados_F = ""
    End If
    
    If sEstados_I <> "" And sEstados_F <> "" Then
        
        If sEstados_I > sEstados_F Then Error 38105
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err

        Case 38105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_INICIAL_MAIOR", Err)
            EstadoInicial.SetFocus
        
        Case 38117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus
        
        Case 43216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)
            NFiscalInicial.SetFocus
            
        Case 38148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INICIAL_MAIOR", Err)
            NaturezaInicial.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173154)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 43199
                
    Serie.Text = ""
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 43199
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173155)

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
Dim sEstados_I As String
Dim sEstados_F As String

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then Error 38122
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 38123
  
    lErro = objRelOpcoes.IncluirParametro("TESTADOINIC", EstadoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38124

    lErro = objRelOpcoes.IncluirParametro("TESTADOFIM", EstadoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38124
           
    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38124

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38125
        
    lErro = objRelOpcoes.IncluirParametro("NNFISCALINIC", NFiscalInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38126

    lErro = objRelOpcoes.IncluirParametro("NNFISCALFIM", NFiscalFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38127
   
    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38128
    
    'Preenche com o Exibir Nota a Nota
    lErro = objRelOpcoes.IncluirParametro("NEXIBTIT", CStr(CheckAnalitico.Value))
    If lErro <> AD_BOOL_TRUE Then Error 38144
    
    lErro = objRelOpcoes.IncluirParametro("NNATUREZAINIC", NaturezaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38149
    
    lErro = objRelOpcoes.IncluirParametro("NNATUREZAFIM", NaturezaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38150

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 38131

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 38122 To 38128
        
        Case 38131, 38144, 38149, 38150

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173156)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 38132

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 38133

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 43200
    
        ComboOpcoes.Text = ""
        Serie.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 38132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 38133, 43200, 43236

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173157)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 38134
    
'    If CheckAnalitico.Value = vbChecked Then
'        gobjRelatorio.sNomeTsk = "ResVenda"
'    Else
'        gobjRelatorio.sNomeTsk = "ResVend1"
'    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 38134

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173158)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 38135

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 38136

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 38137

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 43201
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 38135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 38136, 38137

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173159)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim sEstadoInc As String
Dim sEstadoFin As String, sEstados_I As String, sEstados_F As String

On Error GoTo Erro_Monta_Expressao_Selecao

    If Trim(NFiscalInicial.Text) <> "" Then sExpressao = "NotaFiscal >= " & Forprint_ConvLong(NFiscalInicial.Text)

    If Trim(NFiscalFinal.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NotaFiscal <= " & Forprint_ConvLong(NFiscalFinal.Text)

    End If
    
    If Trim(Serie.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Serie = " & Forprint_ConvTexto(Serie.Text)

    End If
    
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
        
    If Trim(NaturezaInicial.Text) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NaturezaOp >= " & Forprint_ConvTexto(NaturezaInicial.Text)

    End If
    
    If Trim(NaturezaFinal.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NaturezaOp <= " & Forprint_ConvTexto(NaturezaFinal.Text)

    End If

    If EstadoInicial.Text <> "" Then
        sEstados_I = CStr(SCodigo_Extrai(EstadoInicial.Text))
    Else
        sEstados_I = ""
    End If
    
    If EstadoFinal.Text <> "" Then
        sEstados_F = CStr(SCodigo_Extrai(EstadoFinal.Text))
    Else
        sEstados_F = ""
    End If
    
    sEstadoInc = SCodigo_Extrai(sEstados_I)
    sEstadoFin = SCodigo_Extrai(sEstados_F)

    If sEstados_I <> "" Then
   
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Estado >= " & Forprint_ConvTexto(sEstadoInc)
        
    End If

    If sEstados_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Estado <= " & Forprint_ConvTexto(sEstadoFin)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173160)

    End Select

    Exit Function

End Function

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

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
    If lErro <> SUCESSO Then Error 38112
              
    Exit Sub

Erro_NFiscalInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 38112
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173161)
            
    End Select
    
    Exit Sub

End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalFinal_Validate
     
    lErro = Critica_Numero(NFiscalFinal.Text)
    If lErro <> SUCESSO Then Error 38115
        
    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 38115
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173162)
            
    End Select
    
    Exit Sub

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero
         
    If Len(Trim(sNumero)) > 0 Then
        
        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then Error 38118
 
        If CLng(sNumero) < 0 Then Error 38119
        
    End If
 
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = Err

    Select Case Err
                  
        Case 38118
            
        Case 38119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173163)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then Error 38129
    
    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
    
    Carrega_Serie = SUCESSO
    
    Exit Function
    
Erro_Carrega_Serie:

    Carrega_Serie = Err
    
    Select Case Err
    
        Case 38129
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173164)
            
    End Select
    
    Exit Function

End Function

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série na Tela
    Serie.Text = objSerie.sSerie
    
    Call Serie_Validate(bSGECancelDummy)

    Exit Sub

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    'Verifica se a Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
        
    'Verifica se é uma Serie selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 38130
    
    If lErro = 12253 Then Error 54904
    
    Exit Sub
    
Erro_Serie_Validate:

    Cancel = True


    Select Case Err
    
        Case 38130
       
        Case 54904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173165)
    
    End Select
    
    Exit Sub

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 38138

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 38138

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173166)

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
        If lErro <> SUCESSO Then Error 38139

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 38139

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173167)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 38140

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 38140
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173168)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 38141

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 38141
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173169)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 38142

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 38142
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173170)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 38143

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 38143
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173171)

    End Select

    Exit Sub

End Sub

Private Sub LabelNatInicial_Click()

Dim lErro As Long
Dim ObjNatureza As ClassNaturezaOp
Dim colSelecao As Collection


On Error GoTo Erro_LabelNatInicial_Click

    giNaturezaInicial = 1

    If Len(Trim(NaturezaInicial.Text)) <> 0 Then
    
        lErro = Long_Critica(NaturezaInicial.Text)
        If lErro <> SUCESSO Then Error 54905
        
        Set ObjNatureza = New ClassNaturezaOp
        ObjNatureza.sCodigo = NaturezaInicial.Text

    End If

    Call Chama_Tela("NaturezaOperacaoLista", colSelecao, ObjNatureza, objEventoNatureza)
    
    Exit Sub

Erro_LabelNatInicial_Click:

    Select Case Err
    
        Case 54905

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173172)

    End Select

    Exit Sub

End Sub

Private Sub LabelNatFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim ObjNatureza As ClassNaturezaOp

On Error GoTo Erro_LabelNatFinal_Click

    giNaturezaInicial = 0

    If Len(Trim(NaturezaFinal.Text)) <> 0 Then
    
        lErro = Long_Critica(NaturezaFinal.Text)
        If lErro <> SUCESSO Then Error 54906

        Set ObjNatureza = New ClassNaturezaOp
        ObjNatureza.sCodigo = NaturezaFinal.Text

    End If

    Call Chama_Tela("NaturezaOperacaoLista", colSelecao, ObjNatureza, objEventoNatureza)
   
   Exit Sub

Erro_LabelNatFinal_Click:

    Select Case Err
    
        Case 54906

        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173173)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim ObjNatureza As New ClassNaturezaOp

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set ObjNatureza = obj1

    If giNaturezaInicial = 1 Then

        NaturezaInicial.Text = ObjNatureza.sCodigo
        
    Else

        NaturezaFinal.Text = ObjNatureza.sCodigo

    End If

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173174)

    End Select

    Exit Sub

End Sub

Private Sub NaturezaInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjNatureza As New ClassNaturezaOp

On Error GoTo Erro_NaturezaInicial_Validate

    giNaturezaInicial = 1
    
    If Len(Trim(NaturezaInicial.Text)) > 0 Then
        
        lErro = Long_Critica(NaturezaInicial.Text)
        If lErro <> SUCESSO Then Error 38326
    
        ObjNatureza.sCodigo = NaturezaInicial.Text
        
        lErro = CF("NaturezaOperacao_Le", ObjNatureza)
        If lErro <> SUCESSO And lErro <> 17958 Then Error 54907
            
        'Natureza não está cadastrada
        If lErro <> SUCESSO Then Error 54908
        
    End If
       
    Exit Sub

Erro_NaturezaInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 38326
        
        Case 54907

        Case 54908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, ObjNatureza.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173175)

    End Select

    Exit Sub

End Sub

Private Sub NaturezaFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjNatureza As New ClassNaturezaOp

On Error GoTo Erro_NaturezaFinal_Validate

    giNaturezaInicial = 1
    
    If Len(Trim(NaturezaFinal.Text)) > 0 Then
        
        lErro = Long_Critica(NaturezaFinal.Text)
        If lErro <> SUCESSO Then Error 38326
    
        ObjNatureza.sCodigo = NaturezaFinal.Text
        
        lErro = CF("NaturezaOperacao_Le", ObjNatureza)
        If lErro <> SUCESSO And lErro <> 17958 Then Error 54907
            
        'Natureza não está cadastrada
        If lErro <> SUCESSO Then Error 54908
        
    End If
       
    Exit Sub

Erro_NaturezaFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 38326
        
        Case 54907

        Case 54908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, ObjNatureza.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173176)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_RESUMO_VENDAS
    Set Form_Load_Ocx = Me
    Caption = "Resumo de Vendas por Estado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpResVendasUFTA"
    
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
        ElseIf Me.ActiveControl Is NaturezaInicial Then
            Call LabelNatInicial_Click
        ElseIf Me.ActiveControl Is NaturezaFinal Then
            Call LabelNatFinal_Click
        End If
    
    End If

End Sub


Private Sub LabelNatInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNatInicial, Source, X, Y)
End Sub

Private Sub LabelNatInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNatInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelNatFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNatFinal, Source, X, Y)
End Sub

Private Sub LabelNatFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNatFinal, Button, Shift, X, Y)
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

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


Private Sub LabelEstadosAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEstadosAte, Source, X, Y)
End Sub

Private Sub LabelEstadosAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEstadosAte, Button, Shift, X, Y)
End Sub

Private Sub LabelEstadosDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEstadosDe, Source, X, Y)
End Sub

Private Sub LabelEstadosDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEstadosDe, Button, Shift, X, Y)
End Sub

