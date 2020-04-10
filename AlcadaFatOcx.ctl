VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AlcadaFatOcx 
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8415
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   8415
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6120
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "AlcadaFatOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "AlcadaFatOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "AlcadaFatOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "AlcadaFatOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Limites"
      Height          =   1425
      Left            =   120
      TabIndex        =   9
      Top             =   1725
      Width           =   5325
      Begin MSMask.MaskEdBox LimiteOperacao 
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   390
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LimiteMensal 
         Height          =   300
         Left            =   2670
         TabIndex        =   2
         Top             =   885
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Limite Mensal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   915
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Limite por Operação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   645
         TabIndex        =   14
         Top             =   405
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuário"
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5325
      Begin VB.ComboBox CodUsuario 
         Height          =   315
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   375
         Width           =   1515
      End
      Begin VB.Label Nome 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1050
         TabIndex        =   13
         Top             =   915
         Width           =   3900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   390
         TabIndex        =   12
         Top             =   945
         Width           =   555
      End
      Begin VB.Label UsuariosLabel 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   420
         Width           =   660
      End
   End
   Begin VB.ListBox UsuariosComAlcada 
      Height          =   1815
      Left            =   5895
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1350
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Usuários com Habilitação"
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
      Left            =   5865
      TabIndex        =   16
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "AlcadaFatOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoUsuario As AdmEvento
Attribute objEventoUsuario.VB_VarHelpID = -1

Dim iAlterado As Integer

Private Sub Adiciona_Lista(objLiberacaoCredito As ClassLiberacaoCredito)

Dim lErro As Long
Dim objLiberacaoCredito1 As New ClassLiberacaoCredito
Dim colLiberacaoCredito As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Adiciona_Lista
    
    'Procura na coleção se este usuário já tem alçada
    For iIndice = 1 To UsuariosComAlcada.ListCount
        If UsuariosComAlcada.List(iIndice - 1) = objLiberacaoCredito.sCodUsuario Then Exit Sub
    Next
    
    'Se ele é novo adiciona-o na lista
    UsuariosComAlcada.AddItem objLiberacaoCredito.sCodUsuario
    
    Exit Sub
    
Erro_Adiciona_Lista:
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142636)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim objLiberacaoCredito1 As ClassLiberacaoCredito
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_BotaoExcluir_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o usuário está preenchido
    If Len(CodUsuario.Text) = 0 Then Error 48133
        
    objLiberacaoCredito.sCodUsuario = CodUsuario.Text
    
    'Verifica se o usuário tem alçada
    lErro = CF("LiberacaoCredito_Le",objLiberacaoCredito)
    If lErro <> SUCESSO And lErro <> 36968 Then Error 48134
    
    If lErro <> SUCESSO Then Error 48135
    
    'Pede a confirmação da exclusão da alçada do usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ALCADAFAT_USUARIO", objLiberacaoCredito.sCodUsuario)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    lErro = CF("LiberacaoCredito_Exclui",objLiberacaoCredito)
    If lErro <> SUCESSO Then Error 48136
    
    Call Limpa_Tela_AlcadaFat
    
    'Remove o usuário da lista de usuários com alçada
    For iIndice = 0 To UsuariosComAlcada.ListCount - 1
        If UsuariosComAlcada.List(iIndice) = objLiberacaoCredito.sCodUsuario Then
            UsuariosComAlcada.RemoveItem iIndice
            Exit For
        End If
    Next
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
        
        Case 48133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", Err)
                
        Case 48134, 48136
        
        Case 48135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALCADA_NAO_CADASTRADA", Err, objLiberacaoCredito.sCodUsuario)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142637)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 48114
    
    Call Limpa_Tela_AlcadaFat
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:
    
    Select Case Err
            
        Case 48114
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142638)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Gravar_Registro
        
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos estão preenchidos
    If Len(CodUsuario.Text) = 0 Then gError 48115
    If Len(LimiteOperacao.Text) = 0 Then gError 48117
    If Len(LimiteMensal.Text) = 0 Then gError 48116

    'Verifica se LimiteOperacao <= LimiteMensal
    If CDbl(LimiteOperacao.Text) > CDbl(LimiteMensal.Text) Then gError 48142
    
    'Transfere os dados da tela para os obj's
    objLiberacaoCredito.sCodUsuario = CodUsuario.Text
    objLiberacaoCredito.dLimiteMensal = CDbl(Format(LimiteMensal.Text, "Fixed"))
    objLiberacaoCredito.dLimiteOperacao = CDbl(Format(LimiteOperacao.Text, "Fixed"))
    
    'Grava a alçada no BD
    lErro = CF("LiberacaoCredito_Grava",objLiberacaoCredito)
    If lErro <> SUCESSO Then gError 48118
    
    'Adiciona na listbox se necessário
     Call Adiciona_Lista(objLiberacaoCredito)
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 48115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", gErr)
                
        Case 48116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITEMENSAL_NAO_INFORMADO", gErr)
            
        Case 48117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITEOPERACAO_NAO_INFORMADO", gErr)
        
        Case 48142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITES", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142639)
            
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 48150
    
    Call Limpa_Tela_AlcadaFat
      
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case Err
        
        Case 48150
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142640)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Limpa_Tela_AlcadaFat()

Dim lErro As Long

    Call Limpa_Tela(Me)
    
    Nome.Caption = ""
    CodUsuario.ListIndex = -1
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
End Sub

Private Sub CodUsuario_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CodUsuario_Click()

Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_CodUsuario_Click
    
    'Verifica se algum codigo está selecionado
    If CodUsuario.ListIndex = -1 Then Exit Sub
    
    'Coloca o código selecionado nos obj's
    objUsuarios.sCodUsuario = CodUsuario.List(CodUsuario.ListIndex)
    objLiberacaoCredito.sCodUsuario = CodUsuario.List(CodUsuario.ListIndex)
    
    'Le o nome do Usário
    lErro = CF("Usuarios_Le",objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then Error 48111
    
    If lErro <> SUCESSO Then Error 48112
    
    'Coloca o nome do usário na tela
    Nome.Caption = objUsuarios.sNome
    
    'Testa se o usuário selecionado tem alçada, se tiver ele coloca os dados na tela.
    lErro = Traz_Alcada_Tela(objLiberacaoCredito)
    If lErro <> SUCESSO And lErro <> 48109 Then Error 48113
        
    If lErro = 48109 Then
        LimiteOperacao.Text = ""
        LimiteMensal = ""
    End If
    
    Exit Sub
    
Erro_CodUsuario_Click:

    Select Case Err
            
        Case 48111, 48113
        
        Case 48112 'O usuário não está na tabela
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, objUsuarios.sCodUsuario)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142641)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colUsuarios As New Collection

On Error GoTo Erro_Form_Load
            
    Set objEventoUsuario = New AdmEvento
    
    'Carrega a combobox todos os usuários
    lErro = Carrega_Usuarios(colUsuarios)
    If lErro <> SUCESSO Then Error 48099
    
    'Carrega a listbox com usuários que possuem alçada.
    lErro = Carrega_UsuariosComAlcada(colUsuarios)
    If lErro <> SUCESSO Then Error 48088

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 48088, 48099

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142642)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Function Carrega_Usuarios(colUsuarios As Collection) As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos",colUsuarios)
    If lErro <> SUCESSO Then Error 48100

    For Each objUsuarios In colUsuarios
        CodUsuario.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = Err

    Select Case Err

        Case 48100

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142643)

    End Select

    Exit Function

End Function

Private Function Carrega_UsuariosComAlcada(colUsuarios As Collection) As Long
'Carrega a Lista de usuários com habilitação para autorizar crédito

Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim colLiberacaoCredito As New Collection

On Error GoTo Erro_Carrega_UsuariosComAlcada
        
    'Le as Alçadas dos Usuarios passados na Colecao
    lErro = CF("LiberacoesCredito_Filial_Le",colUsuarios, colLiberacaoCredito)
    If lErro <> SUCESSO Then Error 48089
    
    For Each objLiberacaoCredito In colLiberacaoCredito
        UsuariosComAlcada.AddItem objLiberacaoCredito.sCodUsuario
    Next

    Carrega_UsuariosComAlcada = SUCESSO

    Exit Function

Erro_Carrega_UsuariosComAlcada:

    Carrega_UsuariosComAlcada = Err

    Select Case Err

        Case 48089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_VAZIA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142644)

    End Select

    Exit Function

End Function

Private Function Traz_Alcada_Tela(objLiberacaoCredito As ClassLiberacaoCredito) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_Alcada_Tela

    'Verifica se o usuário tem alçada
    lErro = CF("LiberacaoCredito_Le",objLiberacaoCredito)
    If lErro <> SUCESSO And lErro <> 36968 Then Error 48101
            
    If lErro <> SUCESSO Then Error 48109
    
    'Preenche a tela com os dados de objLiberacaoCredito
    For iIndice = 0 To CodUsuario.ListCount - 1
        If objLiberacaoCredito.sCodUsuario = CodUsuario.List(iIndice) Then
            CodUsuario.ListIndex = iIndice
            Exit For
        End If
    Next
    
    LimiteMensal.Text = Format(objLiberacaoCredito.dLimiteMensal, "Fixed")
    LimiteOperacao.Text = Format(objLiberacaoCredito.dLimiteOperacao, "Fixed")
    
    iAlterado = 0
    
    Traz_Alcada_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Alcada_Tela:

    Traz_Alcada_Tela = Err
    
    Select Case Err
            
        Case 48101
            
        Case 48109 'O usuário não tem alçada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142645)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
    
    Set objEventoUsuario = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
            
End Sub

Private Sub LimiteMensal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub LimiteOperacao_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub LimiteOperacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LimiteOperacao_Validate
    
    If Len(LimiteOperacao.Text) = 0 Then Exit Sub
    
    lErro = Valor_Positivo_Critica(LimiteOperacao.Text)
    If lErro <> SUCESSO Then Error 48102
    
    LimiteOperacao.Text = Format(LimiteOperacao.Text, "Fixed")

    Exit Sub
    
Erro_LimiteOperacao_Validate:

    Cancel = True


    Select Case Err
            
        Case 48102
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142646)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub LimiteMensal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LimiteMensal_Validate
    
    If Len(LimiteMensal.Text) = 0 Then Exit Sub
    
    lErro = Valor_Positivo_Critica(LimiteMensal.Text)
    If lErro <> SUCESSO Then Error 48107
    
    LimiteMensal.Text = Format(LimiteMensal.Text, "Fixed")

    Exit Sub
    
Erro_LimiteMensal_Validate:

    Cancel = True


    Select Case Err
            
        Case 48107
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142647)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub ObjEventoUsuario_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim objUsuarios As ClassUsuarios

On Error GoTo Erro_ObjEventoUsuario_evSelecao

    Call Limpa_Tela(Me)
    
    Set objUsuarios = obj1
    
    objLiberacaoCredito.sCodUsuario = objUsuarios.sCodUsuario
    
    'Verifica se o usuário tem alçada
    lErro = CF("LiberacaoCredito_Le",objLiberacaoCredito)
    If lErro <> SUCESSO And lErro <> 36968 Then Error 48096
    
    'Coloca os dados do usuário na tela
    LimiteMensal.Text = objLiberacaoCredito.dLimiteMensal
    LimiteOperacao.Text = objLiberacaoCredito.dLimiteOperacao
    CodUsuario.Text = objLiberacaoCredito.sCodUsuario
    
    'Fecha o comando de setas, se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Me.Show
    
    Exit Sub
    
Erro_ObjEventoUsuario_evSelecao:
    
    Select Case Err
        
        Case 48096 'O usuário não tem alçada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142648)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub UsuariosComAlcada_DblClick()

Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito

On Error GoTo Erro_UsuariosComAlcada
        
    objLiberacaoCredito.sCodUsuario = UsuariosComAlcada.List(UsuariosComAlcada.ListIndex)
    
    lErro = Traz_Alcada_Tela(objLiberacaoCredito)
    If lErro <> SUCESSO Then Error 48110
    
    Exit Sub
    
Erro_UsuariosComAlcada:
    
    Select Case Err
        
        Case 48110
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142649)
    
    End Select
    
    Exit Sub
    
End Sub

'Browse
Private Sub UsuariosLabel_Click()

Dim objUsuarios As New ClassUsuarios
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_UsuariosLabel_Click

    objUsuarios.sCodUsuario = CodUsuario.Text
    
    Call Chama_Tela("UsuarioLista", colSelecao, objUsuarios, objEventoUsuario)

    Exit Sub
    
Erro_UsuariosLabel_Click:
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142650)
    
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros(Optional objLiberacaoCredito As ClassLiberacaoCredito) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objLiberacaoCredito Is Nothing) Then
        
        'traz os dados da alcada para a tela
        lErro = Traz_Alcada_Tela(objLiberacaoCredito)
        If lErro <> SUCESSO And lErro <> 48109 Then Error 48097
        
        'Seleciona o usuario na Combo
        If lErro <> SUCESSO Then
            For iIndice = 0 To CodUsuario.ListCount
                If CodUsuario.List(iIndice) = objLiberacaoCredito.sCodUsuario Then
                    CodUsuario.ListIndex = iIndice
                    Exit For
                End If
            Next
        End If
    
    End If
    
    iAlterado = 0
        
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = Err
    
    Select Case Err
        
        Case 48097
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142651)
        
    End Select
        
    Exit Function
        
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ALCADA_FAT
    Set Form_Load_Ocx = Me
    Caption = "Habilitação de Autorização de Crédito"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AlcadaFat"
    
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

Private Sub Unload(objme As Object)
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodUsuario Then
            Call UsuariosLabel_Click
        
        End If
    End If

End Sub


Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Nome_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nome, Source, X, Y)
End Sub

Private Sub Nome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nome, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub UsuariosLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UsuariosLabel, Source, X, Y)
End Sub

Private Sub UsuariosLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UsuariosLabel, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

