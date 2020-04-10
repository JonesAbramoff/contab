VERSION 5.00
Begin VB.Form GrupoUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo de Usuários"
   ClientHeight    =   4515
   ClientLeft      =   1050
   ClientTop       =   3480
   ClientWidth     =   5775
   Icon            =   "GrupoUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2667.611
   ScaleMode       =   0  'User
   ScaleWidth      =   5422.413
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BotaoDesTodos 
      Caption         =   "Desmarcar Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   4245
      Picture         =   "GrupoUsuarios.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1395
   End
   Begin VB.CommandButton BotaoMarTodos 
      Caption         =   "Marca Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   4245
      Picture         =   "GrupoUsuarios.frx":132C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1395
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
      Left            =   1005
      Picture         =   "GrupoUsuarios.frx":2346
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3675
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
      Left            =   3420
      Picture         =   "GrupoUsuarios.frx":24A0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3660
      Width           =   1380
   End
   Begin VB.ListBox GrupoUsuarios 
      Height          =   2985
      Left            =   255
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   465
      Width           =   3945
   End
   Begin VB.Label Label1 
      Caption         =   "Aplicar esse alteração para os grupos de usuários:"
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   75
      Width           =   3945
   End
End
Attribute VB_Name = "GrupoUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gcolGrupoUsu As Collection

Function Trata_Parametros(colGrupoUsu As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gcolGrupoUsu = colGrupoUsu
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161744)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoCancela_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoDesTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To GrupoUsuarios.ListCount - 1
        GrupoUsuarios.Selected(iIndice) = False
    Next
    
End Sub

Private Sub BotaoMarTodos_Click()

Dim iIndice As Integer

    For iIndice = 0 To GrupoUsuarios.ListCount - 1
        GrupoUsuarios.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objGrupoUsu As ClassGrupoUsuarios

On Error GoTo Erro_BotaoOK_Click
    
    'Verificar se teve algum item marcado
    For iIndice = 0 To GrupoUsuarios.ListCount - 1
    
        If GrupoUsuarios.Selected(iIndice) = True Then
        
            Set objGrupoUsu = New ClassGrupoUsuarios
            
            objGrupoUsu.sCodGrupo = GrupoUsuarios.List(iIndice)
            
            gcolGrupoUsu.Add objGrupoUsu
            
        End If
    
    Next

    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case Err
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 161745)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colGrupoUsu As New Collection
Dim objGrupoUsu As ClassGrupoUsuarios

On Error GoTo Erro_Form_Load

    lErro = CF("GrupoUsuarios_Le_Todos", colGrupoUsu)
    If lErro <> SUCESSO Then gError 129292

    For Each objGrupoUsu In colGrupoUsu
        If objGrupoUsu.dtDataValidade = DATA_NULA Or objGrupoUsu.dtDataValidade >= Date Then GrupoUsuarios.AddItem objGrupoUsu.sCodGrupo
    Next

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 129292
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161746)
    
    End Select
    
    Exit Sub

End Sub
