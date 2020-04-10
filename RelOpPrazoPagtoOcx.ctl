VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPrazoPagtoOcx 
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LockControls    =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   6480
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPrazoPagtoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPrazoPagtoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPrazoPagtoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPrazoPagtoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPrazoPagtoOcx.ctx":0994
      Left            =   870
      List            =   "RelOpPrazoPagtoOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   195
      Width           =   2916
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
      Left            =   4320
      Picture         =   "RelOpPrazoPagtoOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   720
      Left            =   165
      TabIndex        =   34
      Top             =   795
      Width           =   3735
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1470
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   3270
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   330
         Left            =   2310
         TabIndex        =   2
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIni 
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
         Left            =   120
         TabIndex        =   38
         Top             =   300
         Width           =   390
      End
      Begin VB.Label dFim 
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
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Intervalos de Dias"
      Height          =   2640
      Left            =   225
      TabIndex        =   20
      Top             =   1680
      Width           =   2820
      Begin VB.TextBox DiaFinal 
         Height          =   300
         Index           =   1
         Left            =   2190
         TabIndex        =   3
         Top             =   255
         Width           =   405
      End
      Begin VB.TextBox DiaFinal 
         Height          =   300
         Index           =   3
         Left            =   2190
         TabIndex        =   5
         Top             =   1012
         Width           =   405
      End
      Begin VB.TextBox DiaFinal 
         Height          =   300
         Index           =   4
         Left            =   2190
         TabIndex        =   6
         Top             =   1387
         Width           =   405
      End
      Begin VB.TextBox DiaFinal 
         Height          =   300
         Index           =   6
         Left            =   2190
         TabIndex        =   8
         Top             =   2137
         Width           =   405
      End
      Begin VB.TextBox DiaFinal 
         Height          =   300
         Index           =   2
         Left            =   2190
         TabIndex        =   4
         Top             =   637
         Width           =   405
      End
      Begin VB.TextBox DiaFinal 
         Height          =   300
         Index           =   5
         Left            =   2190
         TabIndex        =   7
         Top             =   1747
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "1o. Período "
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
         Left            =   255
         TabIndex        =   33
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "2o. Período "
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
         Left            =   255
         TabIndex        =   32
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "3o. Período"
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
         Left            =   255
         TabIndex        =   31
         Top             =   1050
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "4o. Período "
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
         Left            =   255
         TabIndex        =   30
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "5o. Período"
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
         Left            =   255
         TabIndex        =   29
         Top             =   1770
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "6o. Período"
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
         Left            =   255
         TabIndex        =   28
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label11 
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
         Height          =   270
         Left            =   1725
         TabIndex        =   27
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   1710
         TabIndex        =   26
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label17 
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
         Height          =   255
         Left            =   1725
         TabIndex        =   25
         Top             =   1410
         Width           =   375
      End
      Begin VB.Label Label21 
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
         Height          =   255
         Left            =   1710
         TabIndex        =   24
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   1710
         TabIndex        =   23
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label19 
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
         Height          =   255
         Left            =   1710
         TabIndex        =   22
         Top             =   1770
         Width           =   375
      End
      Begin VB.Label LabelPer2 
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
         Left            =   2835
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame FrameFiliais 
      Caption         =   "Filiais"
      Height          =   1305
      Left            =   3240
      TabIndex        =   17
      Top             =   1680
      Width           =   3015
      Begin VB.ComboBox FilialEmpresaFinal 
         Height          =   315
         Left            =   705
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   2040
      End
      Begin VB.ComboBox FilialEmpresaInicial 
         Height          =   315
         Left            =   705
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label9 
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
         Left            =   285
         TabIndex        =   19
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label13 
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
         Left            =   330
         TabIndex        =   18
         Top             =   390
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
      Left            =   195
      TabIndex        =   40
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label23 
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
      Left            =   2760
      TabIndex        =   39
      Top             =   4425
      Width           =   375
   End
End
Attribute VB_Name = "RelOpPrazoPagtoOcx"
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
    
    If giTipoVersao <> VERSAO_LIGHT Then

        'Preenche as combos de filial Empresa guardando no itemData o codigo
        lErro = Carrega_FilialEmpresa()
        If lErro <> SUCESSO Then Error 48583
    
    Else
    
        FrameFiliais.Visible = False
    
    End If
    
   Call Define_Padrao
                  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 48583

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171312)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
          
'    Devolucoes.Value = 0
    
'    LinhaDetalhe.Value = 0
    
    If giFilialEmpresa <> EMPRESA_TODA Then

        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
    
        FilialEmpresaInicial.Enabled = False
        FilialEmpresaFinal.Enabled = False
        
    Else
    
        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
    
    End If
   
           
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = Err

    Select Case Err
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171313)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 37689
    
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 37690

    Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
   ' DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 37691

    Call DateParaMasked(DataFinal, CDate(sParam))
     'DataFinal.PromptInclude = False
    ''DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
    
    lErro = objRelOpcoes.ObterParametro("NDIAFINAL1", sParam)
    If lErro <> SUCESSO Then Error 37693
    
    DiaFinal(1).Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NDIAFINAL2", sParam)
    If lErro <> SUCESSO Then Error 37695
    
    DiaFinal(2).Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NDIAFINAL3", sParam)
    If lErro <> SUCESSO Then Error 37697
    
    DiaFinal(3).Text = sParam
        
    lErro = objRelOpcoes.ObterParametro("NDIAFINAL4", sParam)
    If lErro <> SUCESSO Then Error 37699
    
    DiaFinal(4).Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NDIAFINAL5", sParam)
    If lErro <> SUCESSO Then Error 37701
    
    DiaFinal(5).Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NDIAFINAL6", sParam)
    If lErro <> SUCESSO Then Error 37703
    
    DiaFinal(6).Text = sParam
    
'    'pega parametro de devolução e exibe
'    lErro = objRelOpcoes.ObterParametro("NDEVOLUCAO", sParam)
'    If lErro <> SUCESSO Then Error 37704
'
'    If sParam <> "" Then Devolucoes.Value = CInt(sParam)
    
'    'pega parametro de linha de detalhe e exibe
'    lErro = objRelOpcoes.ObterParametro("NDETALHE", sParam)
'    If lErro <> SUCESSO Then Error 37705
'
'    LinhaDetalhe.Value = CInt(sParam)
                    
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        'Preenche em Branco
        FilialEmpresaInicial.ListIndex = -1
        FilialEmpresaFinal.ListIndex = -1
        
        'desabilita a combo
        FilialEmpresaInicial.Enabled = False
        FilialEmpresaFinal.Enabled = False
        
    Else
        
        'pega parâmetro FilialEmpresa Inicial
        lErro = objRelOpcoes.ObterParametro("NFILIALINIC", sParam)
        If lErro <> SUCESSO Then Error 48581
         
        FilialEmpresaInicial.Text = sParam
        Call FilialEmpresaInicial_Validate(bSGECancelDummy)
             
        'pega parâmetro FilialEmpresa Final
        lErro = objRelOpcoes.ObterParametro("NFILIALFIM", sParam)
        If lErro <> SUCESSO Then Error 48582
        
        FilialEmpresaFinal.Text = sParam
        Call FilialEmpresaFinal_Validate(bSGECancelDummy)
         
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 37689 To 37705
        
        Case 48581, 48582
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171314)

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
    If lErro <> SUCESSO Then Error 37687
     
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 37687
        
        Case 29884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171315)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 43192
    
    ComboOpcoes.Text = ""
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 43232
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 43192, 43232
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171316)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, lCodigoAuto As Long) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lCodigo As Long
Dim sFilial_I As String
Dim sFilial_F As String

On Error GoTo Erro_PreencherRelOp
       
    lErro = Critica_Parametros(sFilial_I, sFilial_F)
    If lErro <> SUCESSO Then Error 37709
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 37710
   
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37711

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37712
    
    lErro = objRelOpcoes.IncluirParametro("NDIAFINAL1", DiaFinal(1).Text)
    If lErro <> AD_BOOL_TRUE Then Error 37719
    
    lErro = objRelOpcoes.IncluirParametro("NDIAFINAL2", DiaFinal(2).Text)
    If lErro <> AD_BOOL_TRUE Then Error 37720
    
    lErro = objRelOpcoes.IncluirParametro("NDIAFINAL3", DiaFinal(3).Text)
    If lErro <> AD_BOOL_TRUE Then Error 37721
    
    lErro = objRelOpcoes.IncluirParametro("NDIAFINAL4", DiaFinal(4).Text)
    If lErro <> AD_BOOL_TRUE Then Error 37722
    
    lErro = objRelOpcoes.IncluirParametro("NDIAFINAL5", DiaFinal(5).Text)
    If lErro <> AD_BOOL_TRUE Then Error 37723
    
    lErro = objRelOpcoes.IncluirParametro("NDIAFINAL6", DiaFinal(6).Text)
    If lErro <> AD_BOOL_TRUE Then Error 37724
    
'    lErro = objRelOpcoes.IncluirParametro("NDEVOLUCAO", CStr(Devolucoes.Value))
'    If lErro <> AD_BOOL_TRUE Then Error 37725
'
'    lErro = objRelOpcoes.IncluirParametro("NDETALHE", CStr(LinhaDetalhe.Value))
'    If lErro <> AD_BOOL_TRUE Then Error 37726
       
    lErro = objRelOpcoes.IncluirParametro("NFILIALINIC", sFilial_I)
    If lErro <> AD_BOOL_TRUE Then Error 48579
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALINIC", FilialEmpresaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54842

    lErro = objRelOpcoes.IncluirParametro("NFILIALFIM", sFilial_F)
    If lErro <> AD_BOOL_TRUE Then Error 48580
    
    lErro = objRelOpcoes.IncluirParametro("TFILIALFIM", FilialEmpresaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54843
    
    lErro = objRelOpcoes.IncluirParametro("NCODIGO", CStr(lCodigoAuto))
    If lErro <> AD_BOOL_TRUE Then Error 37724
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFilial_I, sFilial_F)
    If lErro <> SUCESSO Then Error 37727
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 37709 To 37727
        
        Case 48579, 48580, 54842, 54843
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171317)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 37728

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 37729

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 43193
    
        ComboOpcoes.Text = ""
        
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then Error 43233
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 37728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 37729, 43193, 43233

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171318)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim lNumAuto As Long
Dim ColPrazos As New Collection

On Error GoTo Erro_BotaoExecutar_Click
    
    lErro = CF("Config_ObterAutomatico","FatConfig", NUM_PROX_RELPRAZOPAGTO, "RelFatPrazoPag", "Codigo", lNumAuto)
    If lErro <> SUCESSO Then Error 48573
    
    lErro = PreencherRelOp(gobjRelOpcoes, lNumAuto)
    If lErro <> SUCESSO Then Error 37730
    
    lErro = Move_Prazo_Memoria(ColPrazos)
    If lErro <> SUCESSO Then Error 48574
    
    lErro = CF("RelFatPrazoPag_Grava",lNumAuto, ColPrazos)
    If lErro <> SUCESSO Then Error 48575
    
    If giFilialEmpresa <> EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "EstPrPag"
    If giFilialEmpresa = EMPRESA_TODA Then gobjRelatorio.sNomeTsk = "EstPrPaE"

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 37730, 48573, 48574, 48575

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171319)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 37731

    lErro = PreencherRelOp(gobjRelOpcoes, 0)
    If lErro Then Error 37732

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 37733

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 43194
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 37731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 37732, 37733, 43194

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171320)

    End Select

    Exit Sub

End Sub


Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFilial_I As String, sFilial_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

    If Trim(DataInicial.ClipText) <> "" Then
    
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
        
    If sFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa >= " & Forprint_ConvInt(CInt(sFilial_I))

    End If
    
    If sFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(CInt(sFilial_F))

    End If
    
'    If sExpressao <> "" Then sExpressao = sExpressao & " E "
'    sExpressao = sExpressao & "NDEVOLUCOES " & Forprint_ConvInt(CInt(Devolucoes.Value))
'
'    If sExpressao <> "" Then sExpressao = sExpressao & " E "
'    sExpressao = sExpressao & "NDETALHE = " & Forprint_ConvInt(CInt(LinhaDetalhe.Value))
     
    If sExpressao <> "" Then
        
        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171321)

    End Select

    Exit Function

End Function

Private Function Critica_Parametros(sFilial_I As String, sFilial_F As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iVazio As Integer
Dim iPreenchido As Integer

On Error GoTo Erro_Critica_Parametros
         
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 37734
    
    End If
    
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        sFilial_I = ""
        sFilial_F = ""
        
    Else
    
        'critica FilialEmpresa Inicial e Final
        If FilialEmpresaInicial.ListIndex <> -1 Then
            sFilial_I = CStr(FilialEmpresaInicial.ItemData(FilialEmpresaInicial.ListIndex))
        Else
            sFilial_I = ""
        End If
        
        If FilialEmpresaFinal.ListIndex <> -1 Then
            sFilial_F = CStr(FilialEmpresaFinal.ItemData(FilialEmpresaFinal.ListIndex))
        Else
            sFilial_F = ""
        End If
                
        If sFilial_I <> "" And sFilial_F <> "" Then
            
            If CInt(sFilial_I) > CInt(sFilial_F) Then Error 48578
            
        End If
    
    End If
    
    'verifica se os períodos anteriores foram preenchidos
    For iIndice = 1 To 6
        If Trim(DiaFinal(iIndice).Text) = "" Then
            iVazio = iIndice
            Exit For
        End If
    Next
    
    For iIndice = 6 To 1 Step -1
        If Trim(DiaFinal(iIndice).Text) <> "" Then
            iPreenchido = iIndice
            Exit For
        End If
    Next
    
    If iVazio <> 0 Then
    
        If iVazio < iPreenchido Then Error 48565
    
    End If
    
    For iIndice = 1 To 5
        
        If Trim(DiaFinal(iIndice).Text) <> "" And Trim(DiaFinal(iIndice + 1)) <> "" Then
            If CLng(DiaFinal(iIndice).Text) >= CLng(DiaFinal(iIndice + 1)) Then Error 48566
        End If
        
    Next
    
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = Err

    Select Case Err
                 
        Case 37734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus

        Case 48565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_PREENCHIDO_INCORRETO", Err)
            
        Case 48566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_ANTERIOR_MENOR", Err)
        
        Case 48578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INICIAL_MAIOR", Err)
            FilialEmpresaInicial.SetFocus
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171322)

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
        If lErro <> SUCESSO Then Error 37747

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37747

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171323)

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
        If lErro <> SUCESSO Then Error 37748

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37748

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171324)

    End Select

    Exit Sub

End Sub


Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37749

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 37749
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171325)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37750

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 37750
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171326)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37751

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 37751
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171327)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37752

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 37752
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171328)

    End Select

    Exit Sub

End Sub

Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero
         
    If Len(Trim(sNumero)) > 0 Then
        
        lErro = Inteiro_Critica(sNumero)
        If lErro <> SUCESSO Then Error 37753
 
        If CInt(sNumero) < 0 Then Error 37754
        
    End If
 
    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = Err

    Select Case Err
                  
        Case 37753
            
        Case 37754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171329)

    End Select

    Exit Function

End Function

Private Sub DiaFinal_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiaFinal_Validate

    lErro = Critica_Numero(DiaFinal(Index).Text)
    If lErro <> SUCESSO Then Error 37756

    Exit Sub

Erro_DiaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37756

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171330)

    End Select

    Exit Sub

End Sub

Function Move_Prazo_Memoria(ColPrazos As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Prazo_Memoria

    ColPrazos.Add 0

    For iIndice = 1 To 6
        
        'se o Periodo esta vazio ----> não preenche mais a colecao
        If DiaFinal(iIndice).Text = "" Then Exit For
       
        ColPrazos.Add CLng(DiaFinal(iIndice).Text)
    Next
            
    ColPrazos.Add 1000
    
    Move_Prazo_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Prazo_Memoria:
    
    Move_Prazo_Memoria = Err
    
        Select Case Err
        
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171331)

    End Select

    Exit Function
    
End Function

Private Sub FilialEmpresaInicial_Validate(Cancel As Boolean)
'Busca a filial com código digitado na lista FilialEmpresa

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_FilialEmpresaInicial_Validate
    
    'se uma opcao da lista estiver selecionada, OK
    If FilialEmpresaInicial.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(FilialEmpresaInicial.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(FilialEmpresaInicial, iCodigo)
    If lErro <> SUCESSO Then Error 48576
        
    Exit Sub

Erro_FilialEmpresaInicial_Validate:

    Cancel = True


    Select Case Err

        Case 48576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171332)

    End Select

    Exit Sub

End Sub

Private Sub FilialEmpresaFinal_Validate(Cancel As Boolean)
'Busca a filial com código digitado na lista FilialEmpresa

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_FilialEmpresaFinal_Validate
    
    'se uma opcao da lista estiver selecionada, OK
    If FilialEmpresaFinal.ListIndex <> -1 Then Exit Sub
    
    If Len(Trim(FilialEmpresaFinal.Text)) = 0 Then Exit Sub
    
    lErro = Combo_Seleciona(FilialEmpresaFinal, iCodigo)
    If lErro <> SUCESSO Then Error 48577
    
    Exit Sub

Erro_FilialEmpresaFinal_Validate:

    Cancel = True


    Select Case Err

        Case 48577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171333)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialEmpresa() As Long
'Carrega as Combos FilialEmpresaInicial e FilialEmpresaFinal

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
    lErro = CF("Cod_Nomes_Le","FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 48584
    
    'preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            FilialEmpresaInicial.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresaInicial.ItemData(FilialEmpresaInicial.NewIndex) = objCodigoNome.iCodigo
    
            FilialEmpresaFinal.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresaFinal.ItemData(FilialEmpresaFinal.NewIndex) = objCodigoNome.iCodigo
        End If
    
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = Err

    Select Case Err

        Case 48584

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171334)

    End Select

    Exit Function

End Function


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRAZO_PAGTO
    Set Form_Load_Ocx = Me
    Caption = "Faturamento por Prazo de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPrazoPagto"
    
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub LabelPer2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPer2, Source, X, Y)
End Sub

Private Sub LabelPer2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPer2, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

