VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl NFeConfig 
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   ScaleHeight     =   8730
   ScaleWidth      =   7800
   Begin VB.CommandButton BotaoDirXml 
      Caption         =   "..."
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
      Left            =   6000
      TabIndex        =   21
      Top             =   810
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Caption         =   "NFCe (Consumidor)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   180
      TabIndex        =   14
      Top             =   1365
      Width           =   7275
      Begin VB.ComboBox versaoNFE 
         Height          =   315
         ItemData        =   "NFeConfig.ctx":0000
         Left            =   1275
         List            =   "NFeConfig.ctx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1215
         Width           =   975
      End
      Begin VB.CommandButton BotaoTesteImpressora 
         Caption         =   "Testar"
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
         Left            =   6045
         TabIndex        =   50
         Top             =   750
         Width           =   930
      End
      Begin VB.ComboBox ModeloImpressora 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "NFeConfig.ctx":001E
         Left            =   1290
         List            =   "NFeConfig.ctx":0041
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   795
         Width           =   3075
      End
      Begin VB.TextBox PortaImpressora 
         Height          =   330
         Left            =   5115
         MaxLength       =   20
         TabIndex        =   46
         ToolTipText     =   "Senha definida pelo contribuinte no software de ativação com 8 a 32 caracteres"
         Top             =   765
         Width           =   735
      End
      Begin VB.Frame Frame5 
         Caption         =   "Configuração de Email"
         Height          =   2385
         Left            =   150
         TabIndex        =   36
         Top             =   2535
         Width           =   6870
         Begin VB.TextBox SMTPSenha 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1110
            PasswordChar    =   "*"
            TabIndex        =   38
            Top             =   1980
            Width           =   2805
         End
         Begin VB.CheckBox NFCeEnviarEmail 
            Caption         =   "Enviar para o cliente um e-mail com o QRCode"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   510
            TabIndex        =   37
            Top             =   330
            Width           =   4410
         End
         Begin MSMask.MaskEdBox SMTP 
            Height          =   315
            Left            =   1110
            TabIndex        =   39
            Top             =   720
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SMTPUsu 
            Height          =   315
            Left            =   1110
            TabIndex        =   40
            Top             =   1545
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Porta 
            Height          =   315
            Left            =   1110
            TabIndex        =   41
            Top             =   1125
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Porta:"
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
            Left            =   135
            TabIndex        =   45
            Top             =   1170
            Width           =   915
         End
         Begin VB.Label LabelSMTPSenha 
            Alignment       =   1  'Right Justify
            Caption         =   "Senha:"
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
            Left            =   150
            TabIndex        =   44
            Top             =   2025
            Width           =   915
         End
         Begin VB.Label LabelSMTPUsu 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuário:"
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
            Left            =   150
            TabIndex        =   43
            Top             =   1575
            Width           =   915
         End
         Begin VB.Label LabelSMTP 
            Alignment       =   1  'Right Justify
            Caption         =   "SMTP:"
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
            Left            =   150
            TabIndex        =   42
            Top             =   750
            Width           =   915
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Codigo de Segurança"
         Height          =   690
         Left            =   135
         TabIndex        =   29
         Top             =   1785
         Width           =   6885
         Begin VB.TextBox NFCECSC 
            Height          =   315
            Left            =   2295
            MaxLength       =   36
            TabIndex        =   32
            Top             =   270
            Width           =   3555
         End
         Begin VB.TextBox idNFCECSC 
            Height          =   330
            Left            =   540
            MaxLength       =   6
            TabIndex        =   30
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Left            =   1575
            TabIndex        =   33
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Label10 
            Caption         =   "Id:"
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
            Left            =   225
            TabIndex        =   31
            Top             =   330
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contingencia Offline"
         Height          =   1170
         Left            =   150
         TabIndex        =   22
         Top             =   4950
         Width           =   6855
         Begin VB.TextBox xJust 
            Height          =   315
            Left            =   900
            MaxLength       =   255
            TabIndex        =   27
            Top             =   720
            Width           =   5745
         End
         Begin VB.CheckBox EmContingencia 
            Caption         =   "Ativada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   23
            Top             =   330
            Width           =   1095
         End
         Begin MSComCtl2.UpDown UpDownDataEmContingencia 
            Height          =   315
            Left            =   2820
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmContingencia 
            Height          =   315
            Left            =   1830
            TabIndex        =   25
            Top             =   315
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraContingencia 
            Height          =   300
            Left            =   3750
            TabIndex        =   34
            Top             =   315
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            Height          =   195
            Index           =   17
            Left            =   3225
            TabIndex        =   35
            Top             =   345
            Width           =   480
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
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
            Left            =   195
            TabIndex        =   28
            Top             =   780
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Em:"
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
            Index           =   0
            Left            =   1455
            TabIndex        =   26
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.TextBox NFCeSerie 
         Height          =   330
         Left            =   810
         MaxLength       =   3
         TabIndex        =   16
         Top             =   315
         Width           =   600
      End
      Begin VB.ComboBox NFCeAmbiente 
         Height          =   315
         ItemData        =   "NFeConfig.ctx":00C6
         Left            =   4905
         List            =   "NFeConfig.ctx":00D0
         TabIndex        =   15
         Top             =   300
         Width           =   2205
      End
      Begin MSMask.MaskEdBox NFCeProxNumNFiscal 
         Height          =   300
         Left            =   3120
         TabIndex        =   17
         Top             =   345
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "versão:"
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
         Left            =   555
         TabIndex        =   52
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label13 
         Caption         =   "Porta:"
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
         Left            =   4515
         TabIndex        =   49
         Top             =   825
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Impressora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   48
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Próximo Número:"
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
         Left            =   1650
         TabIndex        =   20
         Top             =   375
         Width           =   1425
      End
      Begin VB.Label Label7 
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
         Height          =   240
         Left            =   195
         TabIndex        =   19
         Top             =   345
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ambiente:"
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
         Left            =   3960
         TabIndex        =   18
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "NFe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   7590
      Width           =   7305
      Begin VB.ComboBox NFeAmbiente 
         Height          =   315
         ItemData        =   "NFeConfig.ctx":00EB
         Left            =   4935
         List            =   "NFeConfig.ctx":00F5
         TabIndex        =   12
         Top             =   285
         Width           =   2205
      End
      Begin VB.TextBox NFeSerie 
         Height          =   330
         Left            =   810
         MaxLength       =   6
         TabIndex        =   11
         Top             =   285
         Width           =   600
      End
      Begin MSMask.MaskEdBox NFeProxNumNFiscal 
         Height          =   300
         Left            =   3135
         TabIndex        =   8
         Top             =   285
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ambiente:"
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
         Left            =   3990
         TabIndex        =   13
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Height          =   240
         Left            =   195
         TabIndex        =   10
         Top             =   315
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Próximo Número:"
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
         Left            =   1665
         TabIndex        =   9
         Top             =   315
         Width           =   1425
      End
   End
   Begin VB.TextBox CertificadoA1A3 
      Height          =   315
      Left            =   1515
      MaxLength       =   80
      TabIndex        =   5
      Top             =   300
      Width           =   4515
   End
   Begin VB.TextBox DirArqXml 
      Height          =   330
      Left            =   1500
      MaxLength       =   200
      TabIndex        =   3
      Top             =   840
      Width           =   4500
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6555
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "NFeConfig.ctx":0110
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "NFeConfig.ctx":028E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Certificado:"
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
      TabIndex        =   6
      Top             =   360
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "Diretório XMLs:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   885
      Width           =   1425
   End
End
Attribute VB_Name = "NFeConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer

Private gbCarregandoTela As Boolean

Public Function NFeConfig_Grava() As Long

Dim objConfiguracaoNFe As New ClassConfiguracaoNFe, lErro As Long

On Error GoTo Erro_NFeConfig_Grava

    'grava na memória
    Call Move_Tela_Memoria(objConfiguracaoNFe)
    
    'grava no bd
    lErro = CF_ECF("ConfiguracaoNFe_Grava", objConfiguracaoNFe)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjNFeInfo.CopiaConfig(objConfiguracaoNFe)
    
    NFeConfig_Grava = SUCESSO
    
    Exit Function
    
Erro_NFeConfig_Grava:
    
    NFeConfig_Grava = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144089)
    
    End Select
    
    Exit Function

End Function

Private Sub Move_Tela_Memoria(objConfiguracaoNFe As ClassConfiguracaoNFe)

    With objConfiguracaoNFe
        .iNFCeAmbiente = NFCeAmbiente.ItemData(NFCeAmbiente.ListIndex)
        .iNFeAmbiente = NFeAmbiente.ItemData(NFeAmbiente.ListIndex)
        .lNFCeProximoNum = StrParaLong(Trim(NFCeProxNumNFiscal.Text))
        .lNFeProximoNum = StrParaLong(Trim(NFeProxNumNFiscal.Text))
        .sCertificadoA1A3 = Trim(CertificadoA1A3.Text)
        .sDirArqXml = Trim(DirArqXml.Text)
        .sNFCeSerie = Trim(NFCeSerie.Text)
        .sNFeSerie = Trim(NFeSerie.Text)
        .iEmContingencia = EmContingencia.Value
        .dtContingenciaDataEntrada = MaskedParaDate(DataEmContingencia)
        .dContingenciaHoraEntrada = CDbl(StrParaDate(HoraContingencia.Text))
        .sidNFCECSC = Trim(idNFCECSC.Text)
        .sNFCECSC = Trim(NFCECSC.Text)
        .sContigenciaxJust = Trim(xJust.Text)
        If ModeloImpressora.ListIndex <> -1 Then
            .iModeloImpressora = ModeloImpressora.ItemData(ModeloImpressora.ListIndex)
        Else
            .iModeloImpressora = 0
        End If
        .sPortaImpressora = Trim(PortaImpressora.Text)
        .iNFCeEnviarEmail = NFCeEnviarEmail.Value
        .sSMTP = SMTP.Text
        .sSMTPUsu = SMTPUsu.Text
        .sSMTPSenha = SMTPSenha.Text
        .lSMTPPorta = StrParaLong(Porta.Text)
        .iVersaoNFe = versaoNFE.ItemData(versaoNFE.ListIndex)
        
        '??? nao é lido da tela
        .lNFCeProximoLote = gobjNFeInfo.lNFCeProximoLote
        .lNFeProximoLote = gobjNFeInfo.lNFeProximoLote
        .sDirXsd = gobjNFeInfo.sDirXsd
        .iNFCeImprimir = gobjNFeInfo.iNFCeImprimir
        .iNFDescricaoProd = gobjNFeInfo.iNFDescricaoProd
        .iFocaTipoVenda = gobjNFeInfo.iFocaTipoVenda
    
    End With
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objObject As Object, bSaiuContingencia As Boolean

On Error GoTo Erro_Gravar_Registro

    If gobjNFeInfo.iEmContingencia <> 0 And EmContingencia.Value = vbUnchecked Then
        bSaiuContingencia = True
    Else
        bSaiuContingencia = False
    End If
    
    'grava as configurações no arquivo e na memória
    lErro = NFeConfig_Grava()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If Len(Trim(gobjLojaECF.sFTPURL)) > 0 And bSaiuContingencia Then
        
        Set objObject = gobjLojaECF
            
        lErro = CF_ECF("Rotina_FTP_Envio_Caixa", objObject, 2)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If giCodModeloECF = IMPRESSORA_NFCE And gobjNFeInfo.iEmContingencia <> 0 And gsUF = "SP" Then
        giCodModeloECF = IMPRESSORA_SAT_2_5_15
    End If
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144091)
    
    End Select

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 109485
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144093)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    gbCarregandoTela = True
    
    Call Combo_Seleciona_ItemData(NFCeAmbiente, gobjNFeInfo.iNFCeAmbiente)
    Call Combo_Seleciona_ItemData(NFeAmbiente, gobjNFeInfo.iNFeAmbiente)
    DirArqXml.Text = gobjNFeInfo.sDirArqXml
    NFCeProxNumNFiscal.Text = CStr(gobjNFeInfo.lNFCeProximoNum)
    NFeProxNumNFiscal.Text = CStr(gobjNFeInfo.lNFeProximoNum)
    CertificadoA1A3.Text = gobjNFeInfo.sCertificadoA1A3
    NFCeSerie.Text = gobjNFeInfo.sNFCeSerie
    NFeSerie.Text = gobjNFeInfo.sNFeSerie
    EmContingencia.Value = IIf(gobjNFeInfo.iEmContingencia <> 0, vbChecked, vbUnchecked)
    Call DateParaMasked(DataEmContingencia, gobjNFeInfo.dtContingenciaDataEntrada)
    HoraContingencia.PromptInclude = False
    If gobjNFeInfo.dtContingenciaDataEntrada <> DATA_NULA Then HoraContingencia.Text = Format(CDate(gobjNFeInfo.dContingenciaHoraEntrada), "hh:mm:ss")
    HoraContingencia.PromptInclude = True
    
    Call Combo_Seleciona_ItemData(versaoNFE, gobjNFeInfo.iVersaoNFe)

    
    xJust.Text = gobjNFeInfo.sContigenciaxJust
    idNFCECSC.Text = gobjNFeInfo.sidNFCECSC
    NFCECSC.Text = gobjNFeInfo.sNFCECSC
    PortaImpressora.Text = gobjNFeInfo.sPortaImpressora
    NFCeEnviarEmail.Value = IIf(gobjNFeInfo.iNFCeEnviarEmail <> 0, vbChecked, vbUnchecked)
    SMTP.Text = gobjNFeInfo.sSMTP
    SMTPUsu.Text = gobjNFeInfo.sSMTPUsu
    SMTPSenha.Text = gobjNFeInfo.sSMTPSenha
    
    If gobjNFeInfo.lSMTPPorta <> 0 Then
        Porta.PromptInclude = False
        Porta.Text = CStr(gobjNFeInfo.lSMTPPorta)
        Porta.PromptInclude = True
    End If
       
    Call Combo_Seleciona_ItemData(ModeloImpressora, gobjNFeInfo.iModeloImpressora)
    
    gbCarregandoTela = False
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144094)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144095)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Configurações NFe"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFeConfig"

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

Private Sub BotaoTesteImpressora_Click()

Dim lErro As Long, bRelAberto As Boolean
Dim sMensagem As String
Dim objTela As Object

On Error GoTo Erro_BotaoTesteImpressora_Click

    Set objTela = Me
    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_TESTE_IMPRESSAO, objTela)
    bRelAberto = True
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Teste Impressao")
    If lErro <> SUCESSO Then gError 133088
    
    sMensagem = "123456789012345678901234567890123456789012345678"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
    
    sMensagem = "Teste de impressão"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
            
    'guilhotina
    lErro = AFRAC_AcionarGuilhotina("P")
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acionar Guilhotina")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    sMensagem = "123456789012345678901234567890123456789012345678"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
    
    sMensagem = "Teste de impressão"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
            
    'guilhotina
    lErro = AFRAC_AcionarGuilhotina("P")
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acionar Guilhotina")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    sMensagem = "123456789012345678901234567890123456789012345678"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
    
    sMensagem = "Teste de impressão"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
            
    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    bRelAberto = False
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fechar Teste Impressao")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub
    
Erro_BotaoTesteImpressora_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    If bRelAberto Then Call AFRAC_FecharRelatorioGerencial(objTela)

    Exit Sub

End Sub

Private Sub CertificadoA1A3_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmContingencia_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DataEmContingencia_Validate
    
    If Len(Trim(DataEmContingencia.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataEmContingencia.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
        
    Exit Sub
    
Erro_DataEmContingencia_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160203)

    End Select

    Exit Sub
    
End Sub

Private Sub DirArqXml_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EmContingencia_Click()

    If Not gbCarregandoTela Then
    
        If EmContingencia.Value = vbChecked Then
        
            Call DateParaMasked(DataEmContingencia, Date)
            HoraContingencia.PromptInclude = False
            HoraContingencia.Text = Format(CDate(Now), "hh:mm:ss")
            HoraContingencia.PromptInclude = True
            xJust.Text = "Sem acesso a internet"
        
        Else
        
            Call DateParaMasked(DataEmContingencia, DATA_NULA)
            HoraContingencia.PromptInclude = False
            HoraContingencia.Text = ""
            HoraContingencia.PromptInclude = True
            xJust.Text = ""
        
        End If
    
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'**** fim do trecho a ser copiado *****

Public Function objParent() As Object

    Set objParent = Parent
    
End Function

Private Sub BotaoDirXml_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoDirXml_Click

    szTitle = "Localização física dos arquivos .xml"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        DirArqXml.Text = sBuffer
        Call DirArqXml_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoDirXml_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Private Sub DirArqXml_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_DirArqXml_Validate

    If Len(Trim(DirArqXml.Text)) = 0 Then Exit Sub
    
    If right(DirArqXml.Text, 1) <> "\" And right(DirArqXml.Text, 1) <> "/" Then
        iPos = InStr(1, DirArqXml.Text, "/")
        If iPos = 0 Then
            DirArqXml.Text = DirArqXml.Text & "\"
        Else
            DirArqXml.Text = DirArqXml.Text & "/"
        End If
    End If

    If Len(Trim(Dir(DirArqXml.Text, vbDirectory))) = 0 Then gError 192327

    Exit Sub

Erro_DirArqXml_Validate:

    Cancel = True

    Select Case gErr

        Case 192327, 76
            Call Rotina_ErroECF(vbOKOnly, ERRO_DIRETORIO_INVALIDO, gErr, DirArqXml.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 192328)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmContingencia_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmContingencia_DownClick

    lErro = Data_Up_Down_Click(DataEmContingencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call DataEmContingencia_Validate(False)
    
    Exit Sub

Erro_UpDownDataEmContingencia_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160204)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmContingencia_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmContingencia_UpClick

    lErro = Data_Up_Down_Click(DataEmContingencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call DataEmContingencia_Validate(False)
    
    Exit Sub

Erro_UpDownDataEmContingencia_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160205)

    End Select

    Exit Sub

End Sub

Public Sub HoraContingencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HoraContingencia, iAlterado)

End Sub

Public Sub DataEmContingencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmContingencia, iAlterado)

End Sub

Public Sub HoraContingencia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraContingencia_Validate

    'Verifica se a hora de Contingencia foi digitada
    If Len(Trim(HoraContingencia.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraContingencia.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_HoraContingencia_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 157193)

    End Select

    Exit Sub

End Sub

Private Sub xJust_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub xJust_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_xJust_Validate

    If Len(Trim(xJust.Text)) = 0 Then Exit Sub
    
    If Len(Trim(xJust.Text)) < 15 Then gError 192327

    Exit Sub

Erro_xJust_Validate:

    Cancel = True

    Select Case gErr

        Case 192327
            Call Rotina_ErroECF(vbOKOnly, ERRO_XJUST_INVALIDO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 192328)

    End Select

    Exit Sub

End Sub
