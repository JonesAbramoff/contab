VERSION 5.00
Begin VB.Form TrataAlteracao 
   Caption         =   "Aviso de Alteração"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoCancelar 
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
      Height          =   420
      Left            =   3735
      TabIndex        =   3
      Top             =   1980
      Width           =   1365
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
      Height          =   420
      Left            =   855
      TabIndex        =   2
      Top             =   1980
      Width           =   1365
   End
   Begin VB.CheckBox CheckConfMsg 
      Caption         =   "Não exibir esta mensagem novamente"
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
      Left            =   90
      TabIndex        =   1
      Top             =   1530
      Width           =   3795
   End
   Begin VB.Label LabelMensagem 
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
      Height          =   1320
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5730
   End
End
Attribute VB_Name = "TrataAlteracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long


Dim gObjObjetosBD As ClassObjetoBD
Dim giAvisaSobrePosicao As Integer

Private Sub BotaoCancelar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoCancelar_Click

    giRetornoTela = vbCancel

'??? Jones: Acho que se o usuário cancelou não deve alterar a configuração

    Unload Me
    
Exit Sub

Erro_BotaoCancelar_Click:

    Select Case gErr
    
        Case 80443
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175587)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    giRetornoTela = vbOK
        
    lErro = Configura_TrataAlteracao(gObjObjetosBD)
    If lErro <> SUCESSO Then gError 80442

    Unload Me
    
Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
    
        Case 80442
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175588)

    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(objObjetoBD As ClassObjetoBD) As Long

Dim sMensagem As String

    If (gObjObjetosBD Is Nothing) Then
        Set gObjObjetosBD = New ClassObjetoBD
        gObjObjetosBD.sClasseObjeto = objObjetoBD.sClasseObjeto
    End If
    
    sMensagem = objObjetoBD.sNomeObjetoMSG & " já existe no Banco de Dados, esta gravação vai implicar na alteração de seus dados"
    
    If objObjetoBD.iAvisaSobrePosicao = -1 Then
        CheckConfMsg.Enabled = False
    Else
        CheckConfMsg.Enabled = True
    End If
    
    LabelMensagem = sMensagem

End Function

Public Function Configura_TrataAlteracao(objObjetoBD As ClassObjetoBD) As Long
'**Função responsável em chamar nova função que atualiza ObjetosBD

Dim lErro As Long

On Error GoTo Erro_Configura_TrataAlteracao
    
    'Se check estiver marcado então atualiza no ObjetosBD
    If CheckConfMsg = MARCADO Then

        objObjetoBD.iAvisaSobrePosicao = 0
        
        lErro = CF("ObjetosBD_Atualiza", objObjetoBD)
        If lErro <> SUCESSO And lErro <> 80439 Then gError 80441
        
        If lErro = 80441 Then gError 80435
        
    End If

Configura_TrataAlteracao = SUCESSO

    Exit Function
    
Erro_Configura_TrataAlteracao:

    Configura_TrataAlteracao = gErr
    
    Select Case gErr
    
        Case 80435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEOBJETO_INEXISTENTE", gErr, Error, objObjetoBD.sClasseObjeto)
        
        Case 80439
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175589)
        
    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    TrataAlteracao.top = 0

    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175590)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gObjObjetosBD = Nothing
    
End Sub
