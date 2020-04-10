VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Logo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração do Logo"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "Logo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
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
      Left            =   780
      Picture         =   "Logo.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2025
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
      Height          =   555
      Left            =   3450
      Picture         =   "Logo.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2025
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Logo"
      Height          =   615
      Left            =   780
      TabIndex        =   2
      Top             =   1305
      Width           =   4065
      Begin VB.OptionButton OptEmpToda 
         Caption         =   "Empresa Toda"
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
         TabIndex        =   4
         Top             =   255
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton OptFilial 
         Caption         =   "Desta Filial"
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
         Left            =   2205
         TabIndex        =   3
         Top             =   270
         Width           =   1560
      End
   End
   Begin VB.CommandButton BotaoVisualizar 
      Caption         =   "Visualizar"
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
      Left            =   2190
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4545
      TabIndex        =   0
      Top             =   180
      Width           =   300
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2715
      Top             =   2055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Escolhendo Figura para o Produto"
   End
   Begin MSMask.MaskEdBox NomeFigura 
      Height          =   315
      Left            =   765
      TabIndex        =   7
      Top             =   180
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin VB.Image Figura 
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Left            =   765
      Stretch         =   -1  'True
      Top             =   525
      Width           =   4080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Figura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   8
      Top             =   225
      Width           =   600
   End
End
Attribute VB_Name = "Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function Trata_Parametros() As Long
'
End Function

Private Sub BotaoCancela_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_BotaoOK_Click
    
    If OptFilial.Value Then
        iFilialEmpresa = giFilialEmpresa
    Else
        iFilialEmpresa = EMPRESA_TODA
    End If
    
    lErro = CF("Config_Grava", "AdmConfig", "LOCALIZACAO_LOGO", iFilialEmpresa, NomeFigura.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208983)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim sAux As String

On Error GoTo Erro_Form_Load

    sAux = ""
    
    OptFilial.Value = True
    
    lErro = CF("Config_Le", "AdmConfig", "LOCALIZACAO_LOGO", giFilialEmpresa, sAux)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Or sAux = "" Then

        OptEmpToda.Value = True

        'Le a versão do sistema
        lErro = CF("Config_Le", "AdmConfig", "LOCALIZACAO_LOGO", EMPRESA_TODA, sAux)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    NomeFigura.Text = sAux
    
    If sAux <> "" Then Call BotaoVisualizar_Click

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208984)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoVisualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVisualizar_Click

    'verifica se a figura foi preenchida
    If Len(Trim(NomeFigura.Text)) > 0 Then
    
        '?????? fazer um código muito melhor
        'verifica se o arquivo é do tipo imagem
        If GetAttr(NomeFigura.Text) = vbArchive Or GetAttr(NomeFigura.Text) = vbArchive + vbReadOnly Or GetAttr(NomeFigura.Text) = vbNormal Or GetAttr(NomeFigura.Text) = vbNormal + vbReadOnly Or GetAttr(NomeFigura.Text) = 8224 Then
            'coloca a figura na tela
            Figura.Picture = LoadPicture(NomeFigura.Text)
        Else
            gError 208985
        End If
    Else
        Figura.Picture = LoadPicture
    
    End If
    
    Exit Sub
    
Erro_BotaoVisualizar_Click:

    Select Case gErr
    
        Case 53
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_ENCONTRADO", gErr, NomeFigura.Text)
                    
        Case 208985
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_INVALIDO", gErr, NomeFigura.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208986)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProcurar_Click()

On Error GoTo Erro_BotaoProcurar_Click

    ' Set CancelError is True
    CommonDialog.CancelError = True
    ' Set flags
    CommonDialog.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog.Filter = "All Files (*.*)|*.*|Bitmap Files" & _
    "(*.bmp)|*.bmp|Jpg Files (*.jpg)|*.jpg"
    ' Specify default filter
    CommonDialog.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog.ShowOpen

    ' Display name of selected file
     NomeFigura.Text = CommonDialog.FileName
     
     Call BotaoVisualizar_Click
     
    Exit Sub
    
 Call BotaoVisualizar_Click

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub NomeFigura_Validate(Cancel As Boolean)
    Call BotaoVisualizar_Click
End Sub
