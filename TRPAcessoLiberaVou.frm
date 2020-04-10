VERSION 5.00
Begin VB.Form TRPAcessoLiberaVou 
   Caption         =   "Atenção"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Fatura"
      Height          =   630
      Left            =   3735
      TabIndex        =   20
      Top             =   1320
      Width           =   2715
      Begin VB.Label Label3 
         Caption         =   "Número:"
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   240
         Width           =   750
      End
      Begin VB.Label NumeroFat 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   945
         TabIndex        =   21
         Top             =   180
         Width           =   1440
      End
   End
   Begin VB.CommandButton BotaoContinuar 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   1260
   End
   Begin VB.CommandButton BotaoConsultar 
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1305
      TabIndex        =   1
      Top             =   3240
      Width           =   1260
   End
   Begin VB.CommandButton BotaoCancelarFatura 
      Caption         =   "Cancelar Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5220
      TabIndex        =   4
      Top             =   3240
      Width           =   1260
   End
   Begin VB.CommandButton BotaoOCR 
      Caption         =   "Acertar via OCR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2595
      TabIndex        =   2
      Top             =   3240
      Width           =   1260
   End
   Begin VB.CommandButton BotaoLiberar 
      Caption         =   "Liberar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3930
      TabIndex        =   3
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voucher"
      Height          =   600
      Left            =   60
      TabIndex        =   13
      Top             =   645
      Width           =   6390
      Begin VB.Label NumeroVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   4620
         TabIndex        =   19
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label SerieVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2370
         TabIndex        =   18
         Top             =   180
         Width           =   480
      End
      Begin VB.Label TipoVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   645
         TabIndex        =   17
         Top             =   180
         Width           =   480
      End
      Begin VB.Label LabelNumVou2 
         Caption         =   "Número:"
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
         Left            =   3870
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Index           =   2
         Left            =   1815
         TabIndex        =   15
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
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
         Left            =   165
         TabIndex        =   14
         Top             =   225
         Width           =   435
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Vigência"
      Height          =   630
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   3585
      Begin VB.Label VigenciaDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   645
         TabIndex        =   12
         Top             =   195
         Width           =   1110
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Index           =   45
         Left            =   255
         TabIndex        =   11
         Top             =   255
         Width           =   360
      End
      Begin VB.Label VigenciaAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   2370
         TabIndex        =   10
         Top             =   195
         Width           =   1110
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Index           =   46
         Left            =   1995
         TabIndex        =   9
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"TRPAcessoLiberaVou.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   5
      Left            =   150
      TabIndex        =   25
      Top             =   2430
      Width           =   6270
   End
   Begin VB.Label Faturado 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4680
      TabIndex        =   24
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Faturado:"
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
      Index           =   4
      Left            =   3825
      TabIndex        =   23
      Top             =   2085
      Width           =   840
   End
   Begin VB.Label EmVigencia 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2430
      TabIndex        =   7
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Em vigência:"
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
      Index           =   3
      Left            =   1290
      TabIndex        =   6
      Top             =   2070
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Algumas funcionalidades da tela se encontram bloqueadas porque o voucher já foi faturado ou já entrou em vigência."
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
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   180
      Width           =   6270
   End
End
Attribute VB_Name = "TRPAcessoLiberaVou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjVou As ClassTRPVouchers

Function Trata_Parametros(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjVou = objVou
    
    TipoVou.Caption = objVou.sTipVou
    SerieVou.Caption = objVou.sSerie
    NumeroVou.Caption = CStr(objVou.lNumVou)
    VigenciaDe.Caption = Format(objVou.dtDataVigenciaDe, "dd/mm/yyyy")
    VigenciaAte.Caption = Format(objVou.dtDataVigenciaAte, "dd/mm/yyyy")
    NumeroFat.Caption = CStr(objVou.lNumFat)
    
    If Date <= objVou.dtDataVigenciaDe Then
        EmVigencia.ForeColor = vbBlack
        EmVigencia.Caption = "NÃO"
    Else
        EmVigencia.ForeColor = vbRed
        EmVigencia.Caption = "SIM"
    End If
    
    If objVou.lNumFat = 0 Then
        Faturado.ForeColor = vbBlack
        Faturado.Caption = "NÃO"
        BotaoLiberar.Enabled = True
        BotaoCancelarFatura.Enabled = False
    Else
        Faturado.ForeColor = vbRed
        Faturado.Caption = "SIM"
        BotaoLiberar.Enabled = True 'False
        BotaoCancelarFatura.Enabled = True
    End If
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159203)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoConsultar_Click()
Dim objVou As New ClassTRPVouchers
    objVou.sSerie = gobjVou.sSerie
    objVou.sTipVou = gobjVou.sTipVou
    objVou.lNumVou = gobjVou.lNumVou
    Call Chama_Tela_Modal("TRPVoucher", objVou)
End Sub

Private Sub BotaoContinuar_Click()
    Unload Me
End Sub

Private Sub BotaoLiberar_Click()

Dim objFlag As New AdmGenerico
Dim objUsu As New AdmGenerico
Dim lErro As Long

On Error GoTo Erro_BotaoLiberar_Click

    Load TRPLoginAcesso

    lErro = TRPLoginAcesso.Trata_Parametros(objFlag, objUsu)
    If lErro <> SUCESSO Then gError 129290

    TRPLoginAcesso.Show vbModal

    If objFlag.vVariavel = False Then gError 129291
    
    gobjVou.dtDataLibManut = Date
    gobjVou.dHoraLibManut = CDbl(Time)
    gobjVou.sUsuarioLibManut = objUsu.vVariavel
    
    Unload Me
    
    Exit Sub

Erro_BotaoLiberar_Click:
        
End Sub

Private Sub BotaoOCR_Click()
Dim objOcr As New ClassTRPOcorrencias
    objOcr.sSerie = gobjVou.sSerie
    objOcr.sTipoDoc = gobjVou.sTipVou
    objOcr.lNumVou = gobjVou.lNumVou
    objOcr.dtDataEmissao = gdtDataAtual
    Unload Me
    Call Chama_Tela("TRPOcorrencias", objOcr)
End Sub

Private Sub BotaoCancelarFatura_Click()
Dim objFat As New ClassFaturaTRP
    objFat.lNumFat = gobjVou.lNumFat
    Unload Me
    Call Chama_Tela("TRPCancelarFatura", objFat)
End Sub
