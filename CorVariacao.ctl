VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CorVariacao 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   4665
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   2370
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "CorVariacao.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "CorVariacao.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "CorVariacao.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "CorVariacao.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Cor 
      Height          =   315
      Left            =   1275
      TabIndex        =   6
      Top             =   120
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Variacao 
      Height          =   315
      Left            =   1275
      TabIndex        =   8
      Top             =   570
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   3
      Format          =   "000"
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1275
      TabIndex        =   10
      Top             =   1020
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   25
      PromptChar      =   " "
   End
   Begin VB.Label LabelCor 
      Alignment       =   1  'Right Justify
      Caption         =   "Cor:"
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
      Height          =   315
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label LabelVariacao 
      Alignment       =   1  'Right Justify
      Caption         =   "Variação:"
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
      Height          =   315
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label LabelDescricao 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1050
      Width           =   1125
   End
End
Attribute VB_Name = "CorVariacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCor As AdmEvento
Attribute objEventoCor.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cor/Variação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CorVariacao"

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
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objEventoCor = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187185)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCor = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187186)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCorVariacao As ClassCorVariacao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCorVariacao Is Nothing) Then

        lErro = Traz_CorVariacao_Tela(objCorVariacao)
        If lErro <> SUCESSO Then gError 187187

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 187187

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187188)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCorVariacao As ClassCorVariacao) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objCorVariacao.iCor = StrParaInt(Cor.Text)
    objCorVariacao.iVariacao = StrParaInt(Variacao.Text)
    objCorVariacao.sDescricao = Descricao.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187189)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCorVariacao As New ClassCorVariacao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CorVariacao"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objCorVariacao)
    If lErro <> SUCESSO Then gError 187190

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Cor", objCorVariacao.iCor, 0, "Cor"
    colCampoValor.Add "Variacao", objCorVariacao.iVariacao, 0, "Variacao"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 187190

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187191)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCorVariacao As New ClassCorVariacao

On Error GoTo Erro_Tela_Preenche

    objCorVariacao.iCor = colCampoValor.Item("Cor").vValor
    objCorVariacao.iVariacao = colCampoValor.Item("Variacao").vValor

    If objCorVariacao.iCor <> 0 And objCorVariacao.iVariacao <> 0 Then
        lErro = Traz_CorVariacao_Tela(objCorVariacao)
        If lErro <> SUCESSO Then gError 187192
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 187192

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187193)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCorVariacao As New ClassCorVariacao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Cor.Text)) = 0 Then gError 187194
    If Len(Trim(Variacao.Text)) = 0 Then gError 187195
    If Len(Trim(Descricao.Text)) = 0 Then gError 187352
    '#####################

    'Preenche o objCorVariacao
    lErro = Move_Tela_Memoria(objCorVariacao)
    If lErro <> SUCESSO Then gError 187196

    lErro = Trata_Alteracao(objCorVariacao, objCorVariacao.iCor, objCorVariacao.iVariacao)
    If lErro <> SUCESSO Then gError 187197

    'Grava o/a CorVariacao no Banco de Dados
    lErro = CF("CorVariacao_Grava", objCorVariacao)
    If lErro <> SUCESSO Then gError 187198

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 187194
            Call Rotina_Erro(vbOKOnly, "ERRO_COR_NAO_PREENCHIDA", gErr)
            Cor.SetFocus

        Case 187195
            Call Rotina_Erro(vbOKOnly, "ERRO_VARIACAO_NAO_PREENCHIDA", gErr)
            Variacao.SetFocus

        Case 187196, 187197, 187198

        Case 187352
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            Descricao.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187199)

    End Select

    Exit Function

End Function

Function Limpa_Tela_CorVariacao() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CorVariacao

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_CorVariacao = SUCESSO

    Exit Function

Erro_Limpa_Tela_CorVariacao:

    Limpa_Tela_CorVariacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187200)

    End Select

    Exit Function

End Function

Function Traz_CorVariacao_Tela(objCorVariacao As ClassCorVariacao) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_CorVariacao_Tela

    Call Limpa_Tela_CorVariacao
    
    'Lê o CorVariacao que está sendo Passado
    lErro = CF("CorVariacao_Le", objCorVariacao)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187201

    If lErro = SUCESSO Then
    
        Descricao.Text = objCorVariacao.sDescricao

    End If
    
    If objCorVariacao.iCor <> 0 Then
        Cor.PromptInclude = False
        Cor.Text = CStr(objCorVariacao.iCor)
        Cor.PromptInclude = True
    End If
    
    If objCorVariacao.iVariacao <> 0 Then
        Variacao.PromptInclude = False
        Variacao.Text = CStr(objCorVariacao.iVariacao)
        Variacao.PromptInclude = True
    End If

    iAlterado = 0

    Traz_CorVariacao_Tela = SUCESSO

    Exit Function

Erro_Traz_CorVariacao_Tela:

    Traz_CorVariacao_Tela = gErr

    Select Case gErr

        Case 187201

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187202)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 187203

    'Limpa Tela
    Call Limpa_Tela_CorVariacao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 187203

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187204)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187205)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 187206

    Call Limpa_Tela_CorVariacao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 187206

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187207)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCorVariacao As New ClassCorVariacao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Cor.Text)) = 0 Then gError 187208
    If Len(Trim(Variacao.Text)) = 0 Then gError 187209
    '#####################

    objCorVariacao.iCor = StrParaInt(Cor.Text)
    objCorVariacao.iVariacao = StrParaInt(Variacao.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CORVARIACAO", objCorVariacao.iVariacao)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("CorVariacao_Exclui", objCorVariacao)
        If lErro <> SUCESSO Then gError 187210

        'Limpa Tela
        Call Limpa_Tela_CorVariacao

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 187208
            Call Rotina_Erro(vbOKOnly, "ERRO_COR_NAO_PREENCHIDA", gErr)
            Cor.SetFocus

        Case 187209
            Call Rotina_Erro(vbOKOnly, "ERRO_VARIACAO_NAO_PREENCHIDA", gErr)
            Variacao.SetFocus

        Case 187210

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187211)

    End Select

    Exit Sub

End Sub

Private Sub Cor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cor_Validate

    'Verifica se Cor está preenchida
    If Len(Trim(Cor.Text)) <> 0 Then

       'Critica a Cor
       lErro = Inteiro_Critica(Cor.Text)
       If lErro <> SUCESSO Then gError 187212

    End If

    Exit Sub

Erro_Cor_Validate:

    Cancel = True

    Select Case gErr

        Case 187212

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187359)

    End Select

    Exit Sub

End Sub

Private Sub Cor_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Cor, iAlterado)
    
End Sub

Private Sub Cor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Variacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Variacao_Validate

    'Verifica se Variacao está preenchida
    If Len(Trim(Variacao.Text)) <> 0 Then

       'Critica a Variacao
       lErro = Inteiro_Critica(Variacao.Text)
       If lErro <> SUCESSO Then gError 187213

    End If

    Exit Sub

Erro_Variacao_Validate:

    Cancel = True

    Select Case gErr

        Case 187213

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187358)

    End Select

    Exit Sub

End Sub

Private Sub Variacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Variacao, iAlterado)
    
End Sub

Private Sub Variacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187357)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCorVariacao As ClassCorVariacao

On Error GoTo Erro_objEventoCor_evSelecao

    Set objCorVariacao = obj1

    'Mostra os dados do CorVariacao na tela
    lErro = Traz_CorVariacao_Tela(objCorVariacao)
    If lErro <> SUCESSO Then gError 187214

    Me.Show

    Exit Sub

Erro_objEventoCor_evSelecao:

    Select Case gErr

        Case 187214


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187356)

    End Select

    Exit Sub

End Sub

Private Sub LabelCor_Click()

Dim lErro As Long
Dim objCorVariacao As New ClassCorVariacao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCor_Click

    'Verifica se o Cor foi preenchido
    If Len(Trim(Cor.Text)) <> 0 Then

        objCorVariacao.iCor = Cor.Text

    End If

    Call Chama_Tela("CorVariacaoLista", colSelecao, objCorVariacao, objEventoCor, "", "Cor")

    Exit Sub

Erro_LabelCor_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187355)

    End Select

    Exit Sub

End Sub

Private Sub LabelVariacao_Click()

Dim lErro As Long
Dim objCorVariacao As New ClassCorVariacao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelVariacao_Click

    'Verifica se o Variacao foi preenchido
    If Len(Trim(Variacao.Text)) <> 0 Then

        objCorVariacao.iVariacao = Variacao.Text

    End If

    Call Chama_Tela("CorVariacaoLista", colSelecao, objCorVariacao, objEventoCor, "", "Variação")

    Exit Sub

Erro_LabelVariacao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187354)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cor Then Call LabelCor_Click
        If Me.ActiveControl Is Variacao Then Call LabelVariacao_Click
    
    End If
    
End Sub
