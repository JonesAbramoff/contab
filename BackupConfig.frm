VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form BackupConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração do Backup"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "BackupConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoBkpTeste 
      Caption         =   "Fazer backup e restaurar na empresa de testes imediatamente"
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
      Left            =   600
      TabIndex        =   18
      Top             =   5535
      Width           =   3390
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   2
      Left            =   135
      TabIndex        =   21
      Top             =   390
      Visible         =   0   'False
      Width           =   4500
      Begin VB.Frame FrameFTP 
         Caption         =   "FTP"
         Enabled         =   0   'False
         Height          =   4170
         Left            =   45
         TabIndex        =   22
         Top             =   -15
         Width           =   4395
         Begin VB.Frame Frame5 
            Caption         =   "Diretório de retorno do backup"
            Height          =   975
            Left            =   150
            TabIndex        =   46
            Top             =   3105
            Width           =   4110
            Begin VB.CommandButton BotaoFTPDownload 
               Caption         =   "Download"
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
               Left            =   1290
               TabIndex        =   14
               Top             =   570
               Width           =   1620
            End
            Begin VB.TextBox DirDownload 
               Height          =   285
               Left            =   45
               TabIndex        =   13
               Top             =   270
               Width           =   3525
            End
            Begin VB.CommandButton BotaoProcurarFTP 
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
               Left            =   3525
               TabIndex        =   47
               Top             =   210
               Width           =   510
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Dados para conexão"
            Height          =   1605
            Left            =   150
            TabIndex        =   42
            Top             =   180
            Width           =   4110
            Begin VB.TextBox FTPDiretorio 
               Height          =   300
               Left            =   825
               MaxLength       =   50
               TabIndex        =   11
               Top             =   1200
               Width           =   3105
            End
            Begin VB.TextBox FTPURL 
               Height          =   300
               Left            =   825
               MaxLength       =   255
               TabIndex        =   8
               Top             =   240
               Width           =   3135
            End
            Begin VB.TextBox FTPUsername 
               Height          =   300
               Left            =   825
               MaxLength       =   50
               TabIndex        =   9
               Top             =   570
               Width           =   2295
            End
            Begin VB.TextBox FTPPassword 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   825
               MaxLength       =   50
               PasswordChar    =   "*"
               TabIndex        =   10
               Top             =   885
               Width           =   2295
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Dir.:"
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
               Left            =   390
               TabIndex        =   48
               Top             =   1260
               Width           =   375
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "URL:"
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
               Left            =   315
               TabIndex        =   45
               Top             =   285
               Width           =   450
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   45
               TabIndex        =   44
               Top             =   615
               Width           =   720
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   150
               TabIndex        =   43
               Top             =   945
               Width           =   615
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Conexão"
            Height          =   1275
            Left            =   150
            TabIndex        =   23
            Top             =   1785
            Width           =   4110
            Begin VB.CommandButton BotaoFTP 
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
               Height          =   300
               Left            =   1305
               TabIndex        =   12
               Top             =   870
               Width           =   1620
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Status:"
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
               Left            =   330
               TabIndex        =   27
               Top             =   570
               Width           =   615
            End
            Begin VB.Label FTPStatus 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1005
               TabIndex        =   26
               Top             =   525
               Width           =   2940
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Comando:"
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
               Left            =   90
               TabIndex        =   25
               Top             =   210
               Width           =   855
            End
            Begin VB.Label FTPComando 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1020
               TabIndex        =   24
               Top             =   165
               Width           =   2925
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   1
      Left            =   150
      TabIndex        =   20
      Top             =   390
      Width           =   4455
      Begin VB.Frame Frame 
         Caption         =   "Extras"
         Height          =   855
         Left            =   45
         TabIndex        =   41
         Top             =   2280
         Width           =   4380
         Begin VB.CheckBox optFTP 
            Caption         =   "Transferir arquivos compactados por FTP"
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
            Left            =   45
            TabIndex        =   7
            Top             =   525
            Width           =   3975
         End
         Begin VB.CheckBox optZip 
            Caption         =   "Compactar arquivos"
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
            Left            =   45
            TabIndex        =   6
            Top             =   240
            Value           =   1  'Checked
            Width           =   3330
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Acompanhamento"
         Height          =   990
         Left            =   45
         TabIndex        =   36
         Top             =   3150
         Width           =   4380
         Begin VB.Label ProxBackup 
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1590
            TabIndex        =   40
            Top             =   600
            Width           =   2520
         End
         Begin VB.Label UltBackup 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nenhum backup realizado"
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
            Height          =   330
            Left            =   1590
            TabIndex        =   39
            Top             =   240
            Width           =   2520
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Próximo Backup:"
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
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Top             =   630
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Último Backup:"
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
            Index           =   3
            Left            =   240
            TabIndex        =   37
            Top             =   285
            Width           =   1290
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Diretório onde o backup será gerado"
         Height          =   945
         Left            =   45
         TabIndex        =   34
         Top             =   1320
         Width           =   4380
         Begin VB.CheckBox optNomeComData 
            Caption         =   "Incluir a data no nome do arquivo"
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
            Left            =   45
            TabIndex        =   5
            Top             =   645
            Value           =   1  'Checked
            Width           =   3330
         End
         Begin VB.TextBox NomeDiretorio 
            Height          =   285
            Left            =   45
            TabIndex        =   4
            Top             =   315
            Width           =   3735
         End
         Begin VB.CommandButton BotaoProcurar 
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
            Left            =   3780
            TabIndex        =   35
            Top             =   285
            Width           =   510
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Programação"
         Height          =   1260
         Index           =   0
         Left            =   30
         TabIndex        =   28
         Top             =   15
         Width           =   4395
         Begin VB.CheckBox optHabilitado 
            Caption         =   "Habilitado"
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
            Left            =   2925
            TabIndex        =   1
            Top             =   255
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1485
            TabIndex        =   0
            Top             =   195
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   2580
            TabIndex        =   29
            Top             =   195
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox RepetirDias 
            Height          =   315
            Left            =   1500
            TabIndex        =   2
            Top             =   525
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Hora 
            Height          =   315
            Left            =   1500
            TabIndex        =   3
            Top             =   885
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   1005
            TabIndex        =   33
            Top             =   930
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "dias"
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
            Index           =   1
            Left            =   2280
            TabIndex        =   32
            Top             =   570
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Repetir a cada"
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
            Index           =   0
            Left            =   165
            TabIndex        =   31
            Top             =   585
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Iniciar em:"
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
            Index           =   2
            Left            =   570
            TabIndex        =   30
            Top             =   255
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4575
      Left            =   75
      TabIndex        =   19
      Top             =   60
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   8070
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Básico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "FTP"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton BotaoBkp 
      Caption         =   "Fazer backup imediatamente"
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
      Left            =   615
      TabIndex        =   17
      Top             =   5145
      Width           =   3390
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   645
      Picture         =   "BackupConfig.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   1380
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2625
      Picture         =   "BackupConfig.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Width           =   1380
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   180
      Top             =   4845
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
End
Attribute VB_Name = "BackupConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim iFrameAtual As Integer

Function Trata_Parametros() As Long
'
End Function

Private Sub BotaoBkpTeste_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoBkpTeste_Click

    If giExeBkp = MARCADO Then gError 211710
    giExeBkp = MARCADO
    
    'If optHabilitado.Value = vbUnchecked Then gError 211712

    lErro = Grava_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Backup_Executa_Direto", 1, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    giExeBkp = DESMARCADO
    
'    lErro = Traz_Config_Tela
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_BACKUP_REALIZADO_COM_SUCESSO")
    
    Exit Sub
    
Erro_BotaoBkpTeste_Click:

    giExeBkp = DESMARCADO

    Select Case gErr
    
        Case 211710
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUP_EM_EXECUCAO", gErr)
            
        Case 211712
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUP_DESABILITADO", gErr)
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209001)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoBkp_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoBkp_Click

    If giExeBkp = MARCADO Then gError 211710
    giExeBkp = MARCADO
    
    If optHabilitado.Value = vbUnchecked Then gError 211712

    lErro = Grava_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Backup_Executa_Direto", 1)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    giExeBkp = DESMARCADO
    
    lErro = Traz_Config_Tela
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_BACKUP_REALIZADO_COM_SUCESSO")
    
    Exit Sub
    
Erro_BotaoBkp_Click:

    giExeBkp = DESMARCADO

    Select Case gErr
    
        Case 211710
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUP_EM_EXECUCAO", gErr)
            
        Case 211712
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUP_DESABILITADO", gErr)
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209001)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoCancela_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoFTPDownload_Click()

Dim lErro As Long
Dim iBKPAtivo As Integer

On Error GoTo Erro_BotaoFTPDownload_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(DirDownload.Text)) = 0 Then gError 211968

    lErro = CF("Backup_Trata_Download", DirDownload.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    GL_objMDIForm.MousePointer = vbDefault
        
    Call Rotina_Aviso(vbOKOnly, "AVISO_DOWNLOAD_ARQUIVO_SUCESSO")
      
    Exit Sub
    
Erro_BotaoFTPDownload_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 211968
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_INFORMADO", gErr)
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211969)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iBKPAtivo As Integer

On Error GoTo Erro_BotaoOK_Click

    lErro = CF("Backup_Verifica_Habilitado", iBKPAtivo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = Grava_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If optHabilitado.Value = vbChecked Then
        If iBKPAtivo = DESMARCADO Then
            Call Rotina_Aviso(vbOKOnly, "AVISO_BACKUP_REINICIAR_CORPORATOR")
        End If
    End If

    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209001)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Grava_Registro() As Long

Dim lErro As Long
Dim objBackup As New ClassBackupConfig

On Error GoTo Erro_Grava_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    objBackup.lCodigo = 1
    objBackup.sDescricao = "Padrão"
    
    lErro = CF("BackupConfig_Le", objBackup)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If optHabilitado.Value = vbChecked Then
        objBackup.iHabilitado = MARCADO
    Else
        objBackup.iHabilitado = DESMARCADO
    End If
    
    If optNomeComData.Value = vbChecked Then
        objBackup.iIncluirDataNomeArq = MARCADO
    Else
        objBackup.iIncluirDataNomeArq = DESMARCADO
    End If
    
    If Len(Trim(Hora.ClipText)) > 0 Then
        objBackup.dHora = CDbl(CDate(Hora.Text))
    Else
        objBackup.dHora = 0
    End If
    objBackup.dtDataInicio = StrParaDate(Data.Text)
    objBackup.iRepetirDias = StrParaInt(RepetirDias.Text)
    objBackup.sDiretorio = NomeDiretorio.Text
        
    If objBackup.iHabilitado = MARCADO And objBackup.dtDataInicio <> DATA_NULA And objBackup.iRepetirDias > 0 And objBackup.dHora <> 0 Then
    
        If objBackup.dtDataUltBkp = DATA_NULA Then
        'Se vai ser feito pela primeira vez
            'Se está previsto para iniciar depois de hoje a próxima é na data de início
            If objBackup.dtDataInicio > Date Then
                objBackup.dtDataProxBkp = objBackup.dtDataInicio
            Else
            'Se a data de início é hoje ou anterior é para fazer hoje se a hora atual é anterior a hora do backup
            'E amanhã se já passou da hora
                If objBackup.dHora > CDbl(Time) Then
                    objBackup.dtDataProxBkp = Date
                Else
                    objBackup.dtDataProxBkp = Date + 1
                End If
            End If
        Else
            If objBackup.dtDataProxBkp = DATA_NULA Then
                objBackup.dtDataProxBkp = objBackup.dtDataUltBkp + objBackup.iRepetirDias
            End If
        End If
        
    Else
        objBackup.dtDataProxBkp = DATA_NULA
    End If
    
    If optZip.Value = vbChecked Then
        objBackup.iCompactar = MARCADO
    Else
        objBackup.iCompactar = DESMARCADO
    End If
    
    If optFTP.Value = vbChecked Then
        objBackup.iTransfFTP = MARCADO
    Else
        objBackup.iTransfFTP = DESMARCADO
    End If
    
    objBackup.sFTPDir = FTPDiretorio.Text
    objBackup.sFTPSenha = FTPPassword.Text
    objBackup.sFTPURL = FTPURL.Text
    objBackup.sFTPUsu = FTPUsername.Text
    objBackup.sDirDownload = DirDownload.Text
    
'    lErro = CF("Backup_Verifica_Habilitado", iBKPAtivo)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
    lErro = CF("BackupConfig_Grava", objBackup)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    If objBackup.dtDataProxBkp <> DATA_NULA Then
'        If iBKPAtivo = DESMARCADO Then
'            Call Rotina_Aviso(vbOKOnly, "AVISO_BACKUP_REINICIAR_CORPORATOR")
'        End If
'    End If
'
'    Unload Me

    GL_objMDIForm.MousePointer = vbDefault

    Grava_Registro = SUCESSO
    
    Exit Function
    
Erro_Grava_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Grava_Registro = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209001)
    
    End Select
    
    Exit Function

End Function

Private Function Traz_Config_Tela() As Long

Dim lErro As Long
Dim objBackup As New ClassBackupConfig

On Error GoTo Erro_Traz_Config_Tela

    objBackup.lCodigo = 1
    objBackup.sDescricao = "Padrão"
    
    lErro = CF("BackupConfig_Le", objBackup)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
    
        If objBackup.iHabilitado = MARCADO Then
            optHabilitado.Value = vbChecked
        Else
            optHabilitado.Value = vbUnchecked
        End If
        
        If objBackup.iIncluirDataNomeArq = MARCADO Then
            optNomeComData.Value = vbChecked
        Else
            optNomeComData.Value = vbUnchecked
        End If
        
        If objBackup.dHora <> 0 Then
            Hora.PromptInclude = False
            Hora.Text = Format(objBackup.dHora, Hora.Format)
            Hora.PromptInclude = True
        End If
        
        Call DateParaMasked(Data, objBackup.dtDataInicio)
        
        RepetirDias.PromptInclude = False
        RepetirDias.Text = CStr(objBackup.iRepetirDias)
        RepetirDias.PromptInclude = True
        
        If objBackup.dtDataUltBkp <> DATA_NULA Then UltBackup.Caption = Format(objBackup.dtDataUltBkp, "dd/mm/yyyy")
        If objBackup.dtDataProxBkp <> DATA_NULA Then
            If objBackup.dtDataProxBkp < Date Or (objBackup.dtDataProxBkp = Date And objBackup.dHora <= CDbl(Time)) Then
                ProxBackup.Caption = "Dentro de alguns minutos"
            Else
                ProxBackup.Caption = Format(objBackup.dtDataProxBkp, "dd/mm/yyyy")
            End If
        End If
        
        NomeDiretorio.Text = objBackup.sDiretorio
        
        If objBackup.iCompactar = MARCADO Then
            optZip.Value = vbChecked
        Else
            optZip.Value = vbUnchecked
        End If
        
        If objBackup.iTransfFTP = MARCADO Then
            optFTP.Value = vbChecked
        Else
            optFTP.Value = vbUnchecked
        End If
        
        FTPDiretorio.Text = objBackup.sFTPDir
        FTPPassword.Text = objBackup.sFTPSenha
        FTPURL.Text = objBackup.sFTPURL
        FTPUsername.Text = objBackup.sFTPUsu
        DirDownload.Text = objBackup.sDirDownload
    
    End If

    Traz_Config_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Config_Tela:

    Traz_Config_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209002)
    
    End Select
    
    Exit Function

End Function

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    lErro = Traz_Config_Tela
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209002)
    
    End Select
    
    Exit Sub

End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorio.Text, "/")
        If iPos = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 209003

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 209003, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209004)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física dos arquivos"
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
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209005)

    End Select

    Exit Sub
  
End Sub

Private Sub optFTP_Click()
    If optFTP.Value = vbChecked Then
        FrameFTP.Enabled = True
    Else
        FrameFTP.Enabled = False
    End If
End Sub

Private Sub optZip_Click()
    If optZip.Value = vbChecked Then
        optFTP.Enabled = True
    Else
        optFTP.Enabled = False
        optFTP.Value = vbUnchecked
    End If
End Sub

Private Sub UpDownData_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209006)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209007)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'verifica se o campo Data está correto

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se o campo Data foi preenchida
    If Len(Data.ClipText) > 0 Then
        
        'Critica a Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209008)

    End Select

    Exit Sub
    
End Sub

Private Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se Hora está preenchida
    If Len(Trim(Hora.ClipText)) <> 0 Then

       'Critica a Hora
       lErro = Hora_Critica(Hora.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209009)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFTP_Click()

Dim lTeste As Long

On Error GoTo Erro_BotaoFTP_Click

    Inet1.AccessType = icUseDefault
    Inet1.URL = FTPURL.Text
    Inet1.UserName = FTPUsername.Text
    Inet1.Password = FTPPassword.Text
    
    FTPStatus.Caption = ""
    
    lTeste = 0
    FTPComando.Caption = "DIR " & FTPDiretorio.Text & "/*.*"
    BotaoFTP.MousePointer = vbHourglass
    Inet1.Execute , "DIR " & FTPDiretorio.Text & "/*.*"
    Do While FTPStatus.Caption <> "Mensagem completada" And FTPStatus.Caption <> "Erro de comunicação" And FTPStatus.Caption <> "Diretorio inexistente" And lTeste < 100
        lTeste = lTeste + 1
        Sleep (1000)
        DoEvents
    Loop
    
    BotaoFTP.MousePointer = vbDefault
    
    If FTPStatus.Caption = "Mensagem completada" Then
        Call Rotina_Aviso(vbOKOnly, "CONEXAO_BEM_SUCEDIDA")
    ElseIf FTPStatus.Caption = "Diretorio inexistente" Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", FTPDiretorio.Text)
    Else
        Call Rotina_Aviso(vbOKOnly, "NAO_CONSEGUIU_ESTABELECER_CONEXAO")
    End If
'
'        FTPStatus.Caption = ""
'        lTeste = 0
'        FTPComando.Caption = "MKDIR " & FTPDiretorio.Text
'        BotaoFTP.MousePointer = vbHourglass
'        Inet1.Execute , "MKDIR " & FTPDiretorio.Text
'        Do While FTPStatus.Caption <> "Mensagem completada" And lTeste < 100
'            lTeste = lTeste + 1
'            Sleep (1000)
'            DoEvents
'        Loop
'
'        BotaoFTP.MousePointer = vbDefault
'
'        If FTPStatus.Caption = "" Then
'            Call Rotina_Aviso(vbOKOnly, "CONEXAO_BEM_SUCEDIDA")
'        Else
'            Call Rotina_Aviso(vbOKOnly, "NAO_CONSEGUIU_ESTABELECER_CONEXAO")
'        End If
'    End If

    If Inet1.StillExecuting Then Inet1.Cancel
    
    Exit Sub
    
Erro_BotaoFTP_Click:

    If Inet1.StillExecuting Then Inet1.Cancel
    BotaoFTP.MousePointer = vbDefault

    Select Case gErr
    
        Case 35754
            Call Rotina_Erro(vbOKOnly, "NAO_CONSEGUIU_ESTABELECER_CONEXAO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211970)
    
    End Select
    
    Exit Sub

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211971)

    End Select

    Exit Sub

End Sub

Private Sub DirDownload_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_DirDownload_Validate

    If Len(Trim(DirDownload.Text)) = 0 Then Exit Sub
    
    If right(DirDownload.Text, 1) <> "\" And right(DirDownload.Text, 1) <> "/" Then
        iPos = InStr(1, DirDownload.Text, "/")
        If iPos = 0 Then
            DirDownload.Text = DirDownload.Text & "\"
        Else
            DirDownload.Text = DirDownload.Text & "/"
        End If
    End If

    If Len(Trim(Dir(DirDownload.Text, vbDirectory))) = 0 Then gError 211972

    Exit Sub

Erro_DirDownload_Validate:

    Cancel = True

    Select Case gErr

        Case 211972, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, DirDownload.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211973)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurarFTP_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurarFTP_Click

    szTitle = "Localização física dos arquivos"
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
       
        DirDownload.Text = sBuffer
        Call DirDownload_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurarFTP_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211974)

    End Select

    Exit Sub
  
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    
Dim sTeste As String
Dim iPos As Integer
Dim sTeste1 As String

On Error GoTo Erro_Inet1_StateChanged

    Select Case State
    
        Case 1
            FTPStatus.Caption = "Pesquisando IP..."
              
        Case icHostResolved
            FTPStatus.Caption = "IP encontrado"
        
        Case icReceivingResponse
            FTPStatus.Caption = "Recebendo mensagem..."
        
        Case icResponseCompleted
            sTeste1 = " "
            sTeste = ""
            Do While Len(sTeste1) > 0
                sTeste1 = Inet1.GetChunk(1000)
                sTeste = sTeste & sTeste1
            Loop
            If left(FTPComando.Caption, 3) = "DIR" And Len(sTeste) < 5 Then
                FTPStatus.Caption = "Diretorio inexistente"
            Else
                FTPStatus.Caption = "Mensagem completada"
            End If
        Case icConnecting
            FTPStatus.Caption = "Conectando..."
            
        Case icConnected
            FTPStatus.Caption = "Conectado"
            
        Case icRequesting
            FTPStatus.Caption = "Enviando pedido ao servidor..."
            
        Case icRequestSent
            FTPStatus.Caption = "Pedido enviado ao servidor"
            
        Case icDisconnecting
            FTPStatus.Caption = "Desconectando..."
            
        Case icDisconnected
            FTPStatus.Caption = "Desconectado"
    
        Case icError
            FTPStatus.Caption = "Erro de comunicação"
    
        Case icResponseReceived
            FTPStatus.Caption = "Mensagem recebida...aguarde"
    
    End Select
    
    Exit Sub
    
Erro_Inet1_StateChanged:

    Select Case gErr
    
    End Select
    
    Exit Sub

End Sub
