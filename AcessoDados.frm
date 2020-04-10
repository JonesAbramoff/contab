VERSION 5.00
Begin VB.Form AcessoDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados de Acesso"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6045
   Begin VB.Frame Frame5 
      Caption         =   "Limites"
      Height          =   2670
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   2745
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Filiais:"
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
         Left            =   720
         TabIndex        =   4
         Top             =   802
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Logs:"
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
         Left            =   810
         TabIndex        =   5
         Top             =   1274
         Width           =   495
      End
      Begin VB.Label LimiteLogs 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1305
         TabIndex        =   6
         Top             =   1215
         Width           =   1230
      End
      Begin VB.Label LimiteFiliais 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1305
         TabIndex        =   7
         Top             =   735
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresas:"
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
         TabIndex        =   8
         Top             =   330
         Width           =   915
      End
      Begin VB.Label LimiteEmpresas 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1305
         TabIndex        =   9
         Top             =   270
         Width           =   765
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Validade De:"
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
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   1740
         Width           =   1125
      End
      Begin VB.Label ValidadeDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1305
         TabIndex        =   11
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label ValidadeAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1305
         TabIndex        =   12
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Validade Até:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   2220
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Módulos Liberados"
      Height          =   3615
      Left            =   3105
      TabIndex        =   1
      Top             =   60
      Width           =   2760
      Begin VB.ListBox Modulos 
         Enabled         =   0   'False
         Height          =   3180
         Left            =   195
         TabIndex        =   2
         Top             =   255
         Width           =   2325
      End
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
      Height          =   555
      Left            =   465
      Picture         =   "AcessoDados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3030
      Width           =   2085
   End
End
Attribute VB_Name = "AcessoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub BotaoOk_Click()
    Unload Me
End Sub

Public Function Trata_Parametros(Optional objDicConfig As ClassDicConfig) As Long

Dim objModulo As AdmModulo, lErro As Long
Dim sCgc As String, sNomeEmpresa As String
Dim iNumeroLogs As Integer, iNumeroEmpresas As Integer
Dim iNumeroFiliais As Integer, colModulosLib As New Collection
Dim dtDataValidade As Date, sTextoSenha As String

On Error GoTo Erro_Trata_Parametros
    
    If objDicConfig Is Nothing Then
        Set objDicConfig = New ClassDicConfig
        
        lErro = DicConfig_Le(objDicConfig)
        If lErro <> SUCESSO Then Error 62407
    
        lErro = Senha_Empresa_Decifra(objDicConfig.sSenha, sCgc, sNomeEmpresa, iNumeroLogs, iNumeroEmpresas, iNumeroFiliais, colModulosLib, dtDataValidade, sTextoSenha)
        
        objDicConfig.iLimiteLogs = iNumeroLogs
        objDicConfig.iLimiteEmpresas = iNumeroEmpresas
        objDicConfig.iLimiteFiliais = iNumeroFiliais
        Set objDicConfig.colModulosLib = colModulosLib
        objDicConfig.dtValidadeAte = dtDataValidade
    
    End If
        
    
    'Coloca na tela os limites de acesso
    LimiteEmpresas = objDicConfig.iLimiteEmpresas
    LimiteFiliais = objDicConfig.iLimiteFiliais
    LimiteLogs = objDicConfig.iLimiteLogs
    ValidadeDe = Format(objDicConfig.dtValidadeDe, "dd/mm/yyyy")
    ValidadeAte = Format(objDicConfig.dtValidadeAte, "dd/mm/yyyy")

    'Coloca na tela os módulos liberados
    For Each objModulo In objDicConfig.colModulosLib
        Modulos.AddItem objModulo.sNome
    Next

    Trata_Parametros = SUCESSO

    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case 62407
        
        Case Else
        
    End Select

    Exit Function

End Function
