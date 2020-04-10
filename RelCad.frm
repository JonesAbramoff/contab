VERSION 5.00
Begin VB.Form RelCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Relatórios"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "RelCad.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   6210
   Begin VB.CommandButton BotaoEditar 
      Caption         =   "Editar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4545
      TabIndex        =   21
      Top             =   2805
      Width           =   1260
   End
   Begin VB.ComboBox ComboModulo 
      Height          =   315
      ItemData        =   "RelCad.frx":014A
      Left            =   1530
      List            =   "RelCad.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   360
      Width           =   1905
   End
   Begin VB.CheckBox CheckLandscape 
      Caption         =   "Paisagem (Folha Deitada)"
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
      Left            =   1500
      TabIndex        =   15
      Top             =   3960
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3675
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelCad.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelCad.frx":02CC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelCad.frx":07FE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelCad.frx":0988
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox NomeTsk 
      Height          =   330
      Left            =   1530
      MaxLength       =   64
      TabIndex        =   5
      Top             =   2820
      Width           =   2895
   End
   Begin VB.ComboBox CodRelatorio 
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   900
      Width           =   4275
   End
   Begin VB.TextBox Autor 
      Height          =   330
      Left            =   1530
      TabIndex        =   3
      Top             =   2250
      Width           =   4275
   End
   Begin VB.TextBox Tela 
      Height          =   330
      Left            =   1515
      TabIndex        =   4
      Top             =   3375
      Width           =   4275
   End
   Begin VB.TextBox Descricao 
      Height          =   615
      Left            =   1530
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   4275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origem"
      Enabled         =   0   'False
      Height          =   675
      Left            =   1515
      TabIndex        =   18
      Top             =   4380
      Width           =   4305
      Begin VB.OptionButton OptionOrigem 
         Caption         =   "Criado pelo Cliente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2145
         TabIndex        =   20
         Top             =   225
         Width           =   1980
      End
      Begin VB.OptionButton OptionOrigem 
         Caption         =   "Versão Padrão"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   315
         TabIndex        =   19
         Top             =   240
         Width           =   1725
      End
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   750
      TabIndex        =   17
      Top             =   405
      Width           =   690
   End
   Begin VB.Label Label5 
      Caption         =   "Nome do Tsk:"
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
      Left            =   225
      TabIndex        =   9
      Top             =   2865
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Autor:"
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
      Left            =   855
      TabIndex        =   8
      Top             =   2295
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "Tela Auxiliar:"
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
      Height          =   255
      Left            =   285
      TabIndex        =   7
      Top             =   3450
      Width           =   1140
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1410
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Relatório:"
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
      Left            =   570
      TabIndex        =   0
      Top             =   945
      Width           =   825
   End
End
Attribute VB_Name = "RelCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarObjDicModulos As Object

Function Trata_Parametros(objRelatorio As AdmRelatorio) As Long

    Trata_Parametros = SUCESSO

Exit Function

End Function

Private Sub BotaoEditar_Click()
Dim sNomeTsk As String, sNomeTskAux As String
Dim sBuffer As String

    sNomeTsk = Trim(NomeTsk.Text)
    
    If Len(sNomeTsk) = 0 Then Exit Sub
    
    'se o nome do tsk nao contem o path completo
    If InStr(sNomeTsk, "\") = 0 Then
    
        'buscar diretorio configurado
        sBuffer = String(128, 0)
        Call GetPrivateProfileString("Forprint", "DirTsks", "c:\forpw40\", sBuffer, 128, "ADM100.INI")
        
        sBuffer = StringZ(sBuffer)
        If Right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
        sNomeTskAux = sBuffer & sNomeTsk & ".tsk"

    Else
        
        If UCase(Right(sNomeTsk, 4)) <> ".TSK" Then
            
            sNomeTskAux = sNomeTsk & ".tsk"
            
        Else
        
            sNomeTskAux = sNomeTsk
            
        End If
        
    End If
        
    Call Sistema_EditarRel(sNomeTskAux)
    
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim sModulo As String

On Error GoTo Erro_BotaoExcluir_Click

    sModulo = ComboModulo.Text
    
    'Verifica se existe item selecionado na ComboBox
    If CodRelatorio.ListIndex = -1 Then Error 12052

    'pega o relatorio selecionado na ComboBox
    objRelatorio.sCodRel = CodRelatorio.Text

    'deleta o Relatorio no banco de dados
    lErro = CF("Relatorio_Exclui",objRelatorio, sModulo)
    If lErro Then Error 12053

    'Deleta o item selecionado da ComboBox
    CodRelatorio.RemoveItem CodRelatorio.ListIndex

    Call BotaoLimpar_Click

Exit Sub

Erro_BotaoExcluir_Click:

   Select Case Err

        Case 12052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_SELECIONADO", Err)

        Case 12053

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166666)

   End Select

Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload RelCadastro

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim iOperacao As Integer
Dim sGrupo As String
Dim sModulo As String

On Error GoTo Erro_BotaoGravar_Click

    sModulo = mvarObjDicModulos(ComboModulo.Text)
    
    'Verifica se dados do Relatorio foram informados
    If Len(Trim(CodRelatorio.Text)) = 0 Then Error 12022

    'Preenche objeto Relatorio
    objRelatorio.sCodRel = Trim(CodRelatorio.Text)
    objRelatorio.sDescricao = Trim(Descricao.Text)
    objRelatorio.sAutor = Trim(Autor.Text)
    objRelatorio.sTelaAuxiliar = Trim(Tela.Text)
    objRelatorio.sNomeTsk = Trim(NomeTsk.Text)
    objRelatorio.iLandscape = IIf(CheckLandscape.Value = vbChecked, 1, 0)
    
    'grava o Relatorio no banco de dados
    lErro = CF("Relatorio_Grava",objRelatorio, iOperacao, sModulo)

    If lErro Then Error 12023

    If iOperacao = GRAVACAO Then
            'Adiciona o Relatorio a ComboBox
            CodRelatorio.AddItem (objRelatorio.sCodRel)
    End If

    Call BotaoLimpar_Click

Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 12022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_DE_DADOS", Err)

        Case 12023

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166667)

     End Select

     Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    CodRelatorio.Text = ""
    
    Call Limpa_Tela(RelCadastro)

End Sub

Private Sub CodRelatorio_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_CodRelatorio_Change

    If CodRelatorio.Text <> "" Then
    
        objRelatorio.sCodRel = CodRelatorio.Text
    
        lErro = CF("Relatorio_Le",objRelatorio)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12043
    
        If lErro = AD_SQL_SUCESSO Then
            Descricao.Text = objRelatorio.sDescricao
            Autor.Text = objRelatorio.sAutor
            Tela.Text = objRelatorio.sTelaAuxiliar
            NomeTsk.Text = objRelatorio.sNomeTsk
            CheckLandscape.Value = IIf(objRelatorio.iLandscape = 0, vbUnchecked, vbChecked)
            If objRelatorio.iOrigem = REL_ORIGEM_FORPRINT Then
                OptionOrigem(0).Value = True
            Else
                OptionOrigem(1).Value = True
            End If
        End If

    End If
    
    Exit Sub

Erro_CodRelatorio_Change:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 12043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166668)

    End Select

End Sub

Private Sub ComboModulo_Click()

Dim lErro As Long
Dim colRelatorio As New Collection
Dim sCodRel As String
Dim vntCodRel As Variant
Dim sModulo As String

On Error GoTo Erro_ComboModulo_Click

    If ComboModulo.ListIndex <> -1 Then sModulo = ComboModulo.Text
    
    'Preenche a colecao com os nomes dos relatorios existentes no BD
    lErro = Relatorios_Le_NomeModulo(sModulo, colRelatorio)
    If lErro Then Error 12042

    CodRelatorio.Clear
    
    For Each vntCodRel In colRelatorio
        sCodRel = vntCodRel
        CodRelatorio.AddItem (sCodRel)
    Next

    Exit Sub
     
Erro_ComboModulo_Click:

    Select Case Err
          
        Case 12042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err, Error$)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166669)
     
    End Select
     
    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_RelCadastro_Form_Load

    Me.HelpContextID = IDH_RELCADASTRO
    
    Set mvarObjDicModulos = CreateObject("Scripting.Dictionary")
    
    lErro = Carrega_ComboModulo
    If lErro <> SUCESSO Then Error 59296
    
    If ComboModulo.ListCount <> 0 Then ComboModulo.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_RelCadastro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 59296
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166670)

    End Select

    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarObjDicModulos = Nothing
    
End Sub

Private Sub NomeTsk_LostFocus()

Dim lErro As Long

On Error GoTo Erro_NomeTsk_LostFocus

    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeTsk.Text)) > 0 Then

        If Not IniciaLetra(NomeTsk.Text) Then Error 43310

    End If

    Exit Sub

Erro_NomeTsk_LostFocus:

    Select Case Err

        Case 43310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_TSK_NAO_COMECA_LETRA", Err, NomeTsk.Text)
            NomeTsk.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166671)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboModulo() As Long

Dim lErro As Long
Dim colModulo As New Collection
Dim objModulo As AdmModulo

On Error GoTo Erro_Carrega_ComboModulo

    'Lê Módulos existentes no BD
    lErro = Modulos_Le2(colModulo)
    If lErro <> SUCESSO Then Error 59295
    
    'Preenche List da ComboBox Modulo
    For Each objModulo In colModulo
        If objModulo.sSigla <> MODULO_ADM Then
            ComboModulo.AddItem objModulo.sNome
            Call mvarObjDicModulos.Add(objModulo.sNome, objModulo.sSigla)
        End If
    Next
    
    Carrega_ComboModulo = SUCESSO

    Exit Function

Erro_Carrega_ComboModulo:

    Carrega_ComboModulo = Err

    Select Case Err

        Case 59295

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166672)

    End Select

    Exit Function

End Function
