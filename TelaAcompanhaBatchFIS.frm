VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form TelaAcompanhaBatchFIS 
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   Icon            =   "TelaAcompanhaBatchFIS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
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
      Left            =   1545
      TabIndex        =   0
      Top             =   1755
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3330
      Top             =   1740
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   210
      TabIndex        =   1
      Top             =   1260
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label SpedStatus1 
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   495
      Width           =   4395
   End
   Begin VB.Label SPEDStatus2 
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   735
      Width           =   4395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Registros Processados:"
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
      Left            =   210
      TabIndex        =   4
      Top             =   165
      Width           =   2985
   End
   Begin VB.Label TotReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3285
      TabIndex        =   3
      Top             =   90
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Processamento"
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
      Left            =   210
      TabIndex        =   2
      Top             =   1005
      Width           =   1305
   End
End
Attribute VB_Name = "TelaAcompanhaBatchFIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objIN86Modelo As ClassIN86Modelos
Public colIN86TiposArquivo As Collection
Public sNomeArqParam As String
Public iRotinaBatch As Integer
Public iCancelaBatch As Integer
Public dValorAtual As Double
Public dValorTotal As Double

Public objEFD As ClassEFDPisCofinsSel
Public objSpedECF As ClassSpedECFSel

'Rotina Sped Fiscal
Public iFilialEmpresa As Integer
Public sDiretorio As String
Public dtDataIni As Date
Public dtDataFim As Date
Public iIncluiRegInv As Integer
Public iMotivoRegInv As Integer
Public dtDataInv As Date
Public iFiltroNatureza As Integer
Public iIncluiRCPE As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    TotReg.Caption = "0"
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174601)

    End Select

    Exit Sub

End Sub
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If iCancelaBatch <> CANCELA_BATCH Then
        iCancelaBatch = CANCELA_BATCH
        Cancel = 1
        iCancelaBatch = 0
        Cancel = 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set colIN86TiposArquivo = Nothing
    Set objIN86Modelo = Nothing

End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub Cancelar_Click()

Dim lErro As Long

On Error GoTo Erro_Cancelar_Click

    iCancelaBatch = CANCELA_BATCH

    Exit Sub

Erro_Cancelar_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174602)

    End Select

    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** CONTROLE TIMER - INÍCIO ***
Private Sub Timer1_Timer()

Dim objIN86 As ClassIN86
Dim lErro As Long
Dim sErro As String
Dim objTela As Object

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 103644
    
    Set gcolFiliais = New Collection
    
    'carrega a coleção global de filiais
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, gcolFiliais)
    If lErro <> SUCESSO Then gError 103646

    GL_lUltimoErro = SUCESSO
    
    Select Case iRotinaBatch
    
        Case ROTINA_ARQIN86_BATCH
            Set objIN86 = New ClassIN86
            lErro = objIN86.Rotina_Gera_IN86(objIN86Modelo, colIN86TiposArquivo)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 103645

        Case ROTINA_SPED_FISCAL
        
            Set objTela = Me
        
            lErro = CF("Gera_Sped_Fiscal", sDiretorio, iFilialEmpresa, dtDataIni, dtDataFim, objTela, iIncluiRegInv, iMotivoRegInv, dtDataInv, iFiltroNatureza, iIncluiRCPE)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 204109
            
        Case ROTINA_SPED_FISCAL_PIS
        
            Set objEFD.objTela = Me
        
            lErro = CF("Gera_Sped_Fiscal_Pis", objEFD)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 204109

        Case ROTINA_ECF
        
            Set objSpedECF.objTela = Me
        
            lErro = CF("Gera_Sped_ECF", objSpedECF)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End Select

    iCancelaBatch = CANCELA_BATCH

    Unload Me

    Exit Sub

Erro_Timer1_Timer:

    If iCancelaBatch <> CANCELA_BATCH Then

'        sErro = "Houve algum tipo de erro. Verifique o arquivo de log de erros configurado em \windows\adm100.ini ."
'        Call MsgBox(sErro, vbOKOnly, "SGE-Forprint")
    
        Call Rotina_ErrosBatch
    
    End If
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 103626, 103644, 103645, 103646, 204109

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174603)

    End Select

    iCancelaBatch = CANCELA_BATCH
    Unload Me

    Exit Sub

End Sub
'*** CONTROLE TIMER - FIM ***

Public Sub SPED_Status(ByVal sStatus1 As String, ByVal sStatus2 As String)
    SpedStatus1.Caption = sStatus1
    SPEDStatus2.Caption = sStatus2
    DoEvents
End Sub

Public Function Processa_Item(Optional ByVal bAtualizaReg As Boolean = True) As Long

Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Processa_Item

    dValorAtual = dValorAtual + 1
    
    If bAtualizaReg Then
        TotReg.Caption = Format(StrParaLong(TotReg.Caption) + 1, "#,##0")
    End If

    If dValorTotal >= dValorAtual Then
        ProgressBar1.Value = CInt((dValorAtual / dValorTotal) * 100)
    End If
    DoEvents
    
    If iCancelaBatch = CANCELA_BATCH Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_SPED_FISCAL")
        If vbMsgRes = vbYes Then gError ERRO_SEM_MENSAGEM
        iCancelaBatch = 0
    End If
    
    Processa_Item = SUCESSO
    
    Exit Function
    
Erro_Processa_Item:
    
    Processa_Item = gErr
    
    Exit Function

End Function

Public Sub Inicia_Processo(ByVal lContador As Long)
    dValorAtual = 0
    iCancelaBatch = 0
    dValorTotal = lContador
    ProgressBar1.Value = 0
    DoEvents
End Sub
