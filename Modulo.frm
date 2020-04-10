VERSION 5.00
Begin VB.Form Modulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "Modulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
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
      Left            =   1980
      Picture         =   "Modulo.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   945
      Width           =   975
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
      Left            =   555
      Picture         =   "Modulo.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   945
      Width           =   975
   End
   Begin VB.ComboBox ComboModulo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   2310
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Módulo:"
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
      Left            =   270
      TabIndex        =   2
      Top             =   360
      Width           =   690
   End
End
Attribute VB_Name = "Modulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objTelaPrincipal As Form

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoOK_Click

    'Seleciona o modulo escolhido no menu principal
    For iIndice = 0 To objTelaPrincipal.ComboModulo.ListCount - 1
        If objTelaPrincipal.ComboModulo.List(iIndice) = ComboModulo.Text Then
            objTelaPrincipal.ComboModulo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162788)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colUsuarioModulo As New Collection
Dim objUsuarioModulo As ClassUsuarioModulo

On Error GoTo Erro_Form_Load

    'le todos os modulos validos para o usuario/empresa/filial passados como parametro e coloca-os em colModulo
    lErro = CF("UsuarioModulo_Le_UsuarioEmpresa",colUsuarioModulo, gsUsuario, glEmpresa, giFilialEmpresa)
    If lErro <> SUCESSO Then Error 44423
    
    For Each objUsuarioModulo In colUsuarioModulo
        ComboModulo.AddItem objUsuarioModulo.sNomeModulo
    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162789)
    
    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros(objTelaPrincipal1 As Form) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If IsMissing(objTelaPrincipal1) Then Error 44424

    For iIndice = 0 To ComboModulo.ListCount - 1
    
        If ComboModulo.List(iIndice) = objTelaPrincipal1.ComboModulo.Text Then
            ComboModulo.ListIndex = iIndice
            Exit For
        End If
        
    Next
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 44424
            lErro = Rotina_Erro(vbOKOnly, "TELA_MODULO_CHAMADA_SEM_PARAMETRO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162790)
    
    End Select
    
    Exit Function

End Function


Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

