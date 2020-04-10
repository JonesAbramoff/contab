VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form GrupoForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "Grupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame SSFrame1 
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   8
      Top             =   210
      Width           =   5025
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   1590
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Desc 
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   1020
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataValidade 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   1590
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   450
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   570
         TabIndex        =   15
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data Validade:"
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
         Left            =   300
         TabIndex        =   14
         Top             =   1635
         Width           =   1275
      End
      Begin VB.Label Label6 
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
         Left            =   885
         TabIndex        =   13
         Top             =   510
         Width           =   675
      End
   End
   Begin VB.CommandButton Usuarios 
      Caption         =   "Usuários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2190
      Picture         =   "Grupo.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2790
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5550
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Grupo.frx":06F4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Grupo.frx":0872
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Grupo.frx":0DA4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Grupo.frx":0F2E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListGrupos 
      Height          =   2205
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2190
   End
   Begin VB.Label Label1 
      Caption         =   "Grupos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5475
      TabIndex        =   7
      Top             =   930
      Width           =   2055
   End
End
Attribute VB_Name = "GrupoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoExcluir_Click

    If Len(Codigo.Text) = 0 Then Error 8252

    lErro = Grupo_Exclui(Codigo.Text)
    If lErro <> SUCESSO Then Error 8253

    'Exclui grupo da ListBox
    For iIndice = 0 To ListGrupos.ListCount - 1
        If ListGrupos.List(iIndice) = Codigo.Text Then
            ListGrupos.RemoveItem (iIndice)
            Exit For
        End If
    Next
    
    Call Limpa_Tela(GrupoForm)
    
    'Exibe data default
    Call DateParaMasked(DataValidade, DATA_NULA)
        
    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case Err

        Case 8252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODGRUPO_NAO_INFORMADO", Err)

        Case 8253

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161707)

    End Select

End Sub

Private Sub BotaoFechar_Click()

    Unload GrupoForm

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As String
Dim objGrupo As New ClassDicGrupo
Dim iAlteracao As Integer

On Error GoTo Erro_BotaoGravar_Click

    If Len(Codigo.Text) = 0 Then Error 8230

    objGrupo.sCodGrupo = Codigo.Text
    objGrupo.sDescricao = Desc.Text
    If Len(Trim(DataValidade.ClipText)) <> 0 Then
        objGrupo.dtDataValidade = CDate(DataValidade.Text)
    Else
        objGrupo.dtDataValidade = DATA_NULA
    End If

    'If Log.Value = True Then
    '    objGrupo.iLogAtividade = LOG_SIM
    'Else
    '    objGrupo.iLogAtividade = LOG_NAO
    'End If

    lErro = Grupo_Grava(objGrupo, iAlteracao)
    If lErro <> SUCESSO Then Error 8231

    'Se foi inserção, adicionar grupo na ListBox
    If iAlteracao = GRAVACAO Then
        ListGrupos.AddItem (objGrupo.sCodGrupo)
    End If

    Call Limpa_Tela(GrupoForm)

    'Exibe data default
    Call DateParaMasked(DataValidade, DATA_NULA)
        
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 8230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODGRUPO_NAO_INFORMADO", Err)
            Codigo.SetFocus
            
        Case 8231

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161708)

    End Select

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela(GrupoForm)
    
    'Exibe data default
    Call DateParaMasked(DataValidade, DATA_NULA)
        
End Sub

Private Sub DataValidade_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataValidade)

End Sub

Private Sub DataValidade_LostFocus()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_DataValidade_LostFocus

    If DataValidade.ClipText <> "" Then
    
        lErro = Data_Critica(DataValidade.Text)
        If lErro Then Error 8395
        
        dtData = CDate(DataValidade.Text)
        'Compara DataAtual com DataValidade
        If DateDiff("d", Now, dtData) <= 0 Then Error 8396
        
    End If
    
    Exit Sub
    
Erro_DataValidade_LostFocus:

    Select Case Err
            
        Case 8395
            DataValidade.SetFocus
            
        Case 8396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_FUTURA", Err, DataValidade.Text)
            DataValidade.SetFocus
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161709)

    End Select
    
    Exit Sub
   
End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCodGrupo As New Collection
Dim vCodGrupo As Variant
Dim objGrupo As New ClassDicGrupo
Dim iIndex As Integer
'Dim dtDataValidade As Date
Dim sCodGrupo As String
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    Me.HelpContextID = IDH_GRUPO
    
    lErro = Grupos_Le(colCodGrupo)
    If lErro <> SUCESSO Then Error 8219

    'Preenche a ListBox de grupos
    For Each vCodGrupo In colCodGrupo

        ListGrupos.AddItem (vCodGrupo)

    Next

    If Len(gsGrupo) > 0 Then

        sCodGrupo = gsGrupo

        lErro = Grupo_Le(objGrupo)
        If lErro = SUCESSO Then

            Desc.Text = objGrupo.sDescricao
            Call DateParaMasked(DataValidade, objGrupo.dtDataValidade)
        
            'If objGrupo.iLogAtividade = LOG_SIM Then
            '    Log.Value = True
            'Else
            '    Log.Value = False
            'End If

            'Seleciona gsGrupo na ListBox
            For iIndice = 0 To ListGrupos.ListCount - 1

                If ListGrupos.List(iIndice) = gsGrupo Then
                    ListGrupos.Selected(iIndice) = True
                    Exit For
                End If

            Next

         ElseIf lErro = 8224 Then Error 8234

        End If

        Codigo.Text = gsGrupo
        
        gsGrupo = ""

    Else
    
        Call DateParaMasked(DataValidade, DATA_NULA)

    End If

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 8219, 8234

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161710)

    End Select
    
    Exit Sub

End Sub


''Public Function Preenche_Colecao_Grupos(colCodGrupo As Collection) As Long
'''Preenche a coleção colCodGrupo com os códigos de grupos encontrados no BD
'''SUBSTITUÍDA POR Grupos_Le!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''
''Dim lComando As Long
''Dim lErro As Long
''Dim iFim_de_Arquivo As Integer
''Dim sCodGrupo As Integer
''Dim lConexao As Long
''
''On Error GoTo Erro_Preenche_Colecao_Grupos
''
''    sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
''
''    lConexao = GL_lConexaoDic
''
''    lComando = Comando_AbrirExt(lConexao)
''    If lComando = 0 Then Error 8215
''
''    'seleciona no BD todos os códigos de grupos
''    lErro = Comando_Executar(lComando, "SELECT CodGrupo FROM GruposDeUsuarios", sCodGrupo)
''    If lErro <> AD_SQL_SUCESSO Then Error 8216
''
''    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
''
''    If iFim_de_Arquivo = AD_SQL_ERRO Then Error 8217
''
''    'preenche a coleção
''    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
''
''        colCodGrupo.Add (Str(sCodGrupo))
''
''        iFim_de_Arquivo = Comando_BuscarProximo(lComando)
''
''        If iFim_de_Arquivo = AD_SQL_ERRO Then Error 8218
''
''    Loop
''
''    lErro = Comando_Fechar(lComando)
''
''    Preenche_Colecao_Grupos = SUCESSO
''
''    Exit Function
''
''Erro_Preenche_Colecao_Grupos:
''
''    Preenche_Colecao_Grupos = Err
''
''    Select Case Err
''
''        Case 8215
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 8216, 8217, 8218
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO1", Err, sCodGrupo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161711)
''
''    End Select
''
''    Exit Function
''
''End Function

Private Sub ListGrupos_DblClick()

Dim lErro As Long
Dim objGrupo As New ClassDicGrupo
Dim dtDataValidade As Date

On Error GoTo Erro_ListGrupos_DblClick

    objGrupo.sCodGrupo = ListGrupos.List(ListGrupos.ListIndex)

    lErro = Grupo_Le(objGrupo)
    If lErro <> SUCESSO Then
        If lErro = 8222 Or lErro = 8224 Then Error 8220

        'Grupo nao existe no BD: excluir da listbox
        ListGrupos.RemoveItem (ListGrupos.ListIndex)
    End If

    Codigo.Text = objGrupo.sCodGrupo
    Desc.Text = objGrupo.sDescricao
    'dtDataValidade = Format(objGrupo.dtDataValidade, "dd/mm/yy")
    Call DateParaMasked(DataValidade, objGrupo.dtDataValidade)

    'If objGrupo.iLogAtividade = LOG_SIM Then
    '    Log.Value = True
    'Else
    '    Log.Value = False
    'End If

    Exit Sub

Erro_ListGrupos_DblClick:

    Select Case Err

        Case 8220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161712)

    End Select

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_UpDown1_DownClick
    
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub
    
    sData = DataValidade.Text
    
    lErro = Data_Diminui(sData)
    If lErro Then Error 8391
    
    dtData = CDate(sData)
    'Compara DataAtual com DataValidade
    If DateDiff("d", Now, dtData) <= 0 Then Error 8392
    
    DataValidade.Text = sData
    
    Exit Sub
    
Erro_UpDown1_DownClick:

    Select Case Err
            
        Case 8391  'Já foi tratado na rotina chamada
        
        Case 8392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_FUTURA", Err, sData)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161713)

    End Select
    
    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String
Dim dtData As Date

On Error GoTo Erro_UpDown1_UpClick
    
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub
    
    sData = DataValidade.Text
    
    lErro = Data_Aumenta(sData)
    If lErro Then Error 8393
    
    dtData = CDate(sData)
    'Compara DataAtual com DataValidade
    If DateDiff("d", Now, dtData) <= 0 Then Error 8394
    
    DataValidade.Text = sData
    
    Exit Sub
    
Erro_UpDown1_UpClick:

    Select Case Err
            
        Case 8393  'Já foi tratado na rotina chamada
        
        Case 8394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_FUTURA", Err, sData)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161714)

    End Select
    
    Exit Sub

End Sub

Private Sub Usuarios_Click()

Dim sCodGrupo As String
Dim sCodUsuario As String
Dim lErro As String
Dim bAchou As Boolean, iIndice As Integer

On Error GoTo Erro_Usuarios_Click

    sCodGrupo = Codigo.Text
    gsGrupo = ""
    gsUsuario = ""
    
    If Len(sCodGrupo) > 0 Then

        bAchou = False
        For iIndice = 0 To ListGrupos.ListCount - 1
            If ListGrupos.List(iIndice) = sCodGrupo Then
                bAchou = True
                Exit For
            End If
        Next
        
        If Not bAchou Then Error 62421
        
        lErro = Grupo_Obtem_Usuario(sCodGrupo, sCodUsuario)
        If lErro <> SUCESSO And lErro <> 8228 Then Error 8225

        If lErro <> SUCESSO Then
        
            gsUsuario = ""
            gsGrupo = sCodGrupo
        
        Else
        
            gsUsuario = sCodUsuario
            gsGrupo = sCodGrupo
            
        End If

    End If

    UsuarioTela.Show

    Exit Sub

Erro_Usuarios_Click:

    Select Case Err

        Case 8225

        Case 62421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPO_NAO_CADASTRADO", Err, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161715)

    End Select

End Sub

'Public Function Obtem_Atributos_GrupoDeUsu(objGrupo As ClassDicGrupo) As Long
'
'Dim lComando As Long
'Dim lErro As Long
'Dim lConexao As Long
'
'On Error GoTo Erro_Obtem_Atributos_GrupoDeUsu
'
'    objGrupo.sDescricao = String(STRING_GRUPO_DESCRICAO, 0)
'
'    lConexao = GL_lConexaoDic
'
'    lComando = Comando_AbrirExt(lConexao)
'    If lComando = 0 Then Error 8221
'
'    'Seleciona o Grupo desejado no BD
'    lErro = Comando_Executar(lComando, "SELECT Descricao, DataValidade FROM GruposDeUsuarios WHERE CodGrupo=?", objGrupo.sDescricao, objGrupo.dtDataValidade, objGrupo.sCodGrupo)
'    If lErro <> AD_SQL_SUCESSO Then Error 8222
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO Then
'        If lErro = AD_SQL_SEM_DADOS Then
'            'Não existe tal Grupo
'            Error 8223
'        Else
'            Error 8224
'        End If
'    End If
'
'    lErro = Comando_Fechar(lComando)
'
'    Obtem_Atributos_GrupoDeUsu = SUCESSO
'
'Exit Function
'
'Erro_Obtem_Atributos_GrupoDeUsu:
'
'    Obtem_Atributos_GrupoDeUsu = Err
'
'    Select Case Err
'
'        Case 8221
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 8222, 8224
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO1", Err, objGrupo.sCodGrupo)
'            lErro = Comando_Fechar(lComando)
'
'        Case 8223
'            lErro = Comando_Fechar(lComando)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161716)
'
'    End Select
'
'    Exit Function
'
'End Function
