VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RastreamentoLoteLoc 
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   ScaleHeight     =   1170
   ScaleWidth      =   5775
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   3585
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RastreamentoLoteLoc.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "RastreamentoLoteLoc.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "RastreamentoLoteLoc.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "RastreamentoLoteLoc.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Localizacao 
      Height          =   315
      Left            =   1275
      TabIndex        =   6
      Top             =   660
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin VB.Label LabelLocalizacao 
      Alignment       =   1  'Right Justify
      Caption         =   "Localização:"
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
      Left            =   -345
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   690
      Width           =   1500
   End
End
Attribute VB_Name = "RastreamentoLoteLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Localização do Lote"
    
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RastreamentoLoteLoc"

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

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoLocalizacao = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198733)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoLocalizacao = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198734)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objRastreamentoLoteLoc As ClassRastreamentoLoteLoc) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objRastreamentoLoteLoc Is Nothing) Then

        lErro = Traz_RastreamentoLoteLoc_Tela(objRastreamentoLoteLoc)
        If lErro <> SUCESSO Then gError 198735

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 198735

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198736)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objRastreamentoLoteLoc As ClassRastreamentoLoteLoc) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objRastreamentoLoteLoc.sLocalizacao = Localizacao.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198737)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objRastreamentoLoteLoc As New ClassRastreamentoLoteLoc

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RastreamentoLoteLoc"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objRastreamentoLoteLoc)
    If lErro <> SUCESSO Then gError 198738

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Localizacao", objRastreamentoLoteLoc.sLocalizacao, STRING_RASTRO_LOCALIZACAO, "Localizacao"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 198738

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198739)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objRastreamentoLoteLoc As New ClassRastreamentoLoteLoc

On Error GoTo Erro_Tela_Preenche

    objRastreamentoLoteLoc.sLocalizacao = colCampoValor.Item("Localizacao").vValor

    If Len(Trim(objRastreamentoLoteLoc.sLocalizacao)) > 0 Then

        lErro = Traz_RastreamentoLoteLoc_Tela(objRastreamentoLoteLoc)
        If lErro <> SUCESSO Then gError 198740

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 198740

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198741)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRastreamentoLoteLoc As New ClassRastreamentoLoteLoc

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Localizacao.Text)) = 0 Then gError 198742
    '#####################

    'Preenche o objRastreamentoLoteLoc
    lErro = Move_Tela_Memoria(objRastreamentoLoteLoc)
    If lErro <> SUCESSO Then gError 198743

    lErro = Trata_Alteracao(objRastreamentoLoteLoc, objRastreamentoLoteLoc.sLocalizacao)
    If lErro <> SUCESSO Then gError 198744

    'Grava o/a RastreamentoLoteLoc no Banco de Dados
    lErro = CF("RastreamentoLoteLoc_Grava", objRastreamentoLoteLoc)
    If lErro <> SUCESSO Then gError 198745

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198742
            Call Rotina_Erro(vbOKOnly, "ERRO_LOC_RASTROLOTELOC_NAO_PREENCHIDO", gErr)
            Localizacao.SetFocus

        Case 198743, 198744, 198745

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198746)

    End Select

    Exit Function

End Function

Function Limpa_Tela_RastreamentoLoteLoc() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RastreamentoLoteLoc

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_RastreamentoLoteLoc = SUCESSO

    Exit Function

Erro_Limpa_Tela_RastreamentoLoteLoc:

    Limpa_Tela_RastreamentoLoteLoc = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198747)

    End Select

    Exit Function

End Function

Function Traz_RastreamentoLoteLoc_Tela(objRastreamentoLoteLoc As ClassRastreamentoLoteLoc) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_RastreamentoLoteLoc_Tela
    Call Limpa_Tela_RastreamentoLoteLoc
        Localizacao.Text = objRastreamentoLoteLoc.sLocalizacao

    'Lê o RastreamentoLoteLoc que está sendo Passado
    lErro = CF("RastreamentoLoteLoc_Le", objRastreamentoLoteLoc)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198748

    If lErro = SUCESSO Then

        Localizacao.Text = objRastreamentoLoteLoc.sLocalizacao

    End If

    iAlterado = 0

    Traz_RastreamentoLoteLoc_Tela = SUCESSO

    Exit Function

Erro_Traz_RastreamentoLoteLoc_Tela:

    Traz_RastreamentoLoteLoc_Tela = gErr

    Select Case gErr

        Case 198748

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198749)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 198750

    'Limpa Tela
    Call Limpa_Tela_RastreamentoLoteLoc

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 198750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198751)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198752)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 198753

    Call Limpa_Tela_RastreamentoLoteLoc

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 198753

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198754)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRastreamentoLoteLoc As New ClassRastreamentoLoteLoc
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Localizacao.Text)) = 0 Then gError 198755
    '#####################

    objRastreamentoLoteLoc.sLocalizacao = Localizacao.Text

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_RASTREAMENTOLOTELOC", objRastreamentoLoteLoc.sLocalizacao)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("RastreamentoLoteLoc_Exclui", objRastreamentoLoteLoc)
        If lErro <> SUCESSO Then gError 198756

        'Limpa Tela
        Call Limpa_Tela_RastreamentoLoteLoc

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198755
            Call Rotina_Erro(vbOKOnly, "ERRO_LOC_RASTROLOTELOC_NAO_PREENCHIDO", gErr)
            Localizacao.SetFocus

        Case 198756

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198757)

    End Select

    Exit Sub

End Sub

Private Sub Localizacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Localizacao_Validate

    'Verifica se Localizacao está preenchida
    If Len(Trim(Localizacao.Text)) <> 0 Then

       '#######################################
       'CRITICA Localizacao
       '#######################################

    End If

    Exit Sub

Erro_Localizacao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198758)

    End Select

    Exit Sub

End Sub

Private Sub Localizacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoLocalizacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastreamentoLoteLoc As ClassRastreamentoLoteLoc

On Error GoTo Erro_objEventoLocalizacao_evSelecao

    Set objRastreamentoLoteLoc = obj1

    'Mostra os dados do RastreamentoLoteLoc na tela
    lErro = Traz_RastreamentoLoteLoc_Tela(objRastreamentoLoteLoc)
    If lErro <> SUCESSO Then gError 198759

    Me.Show

    Exit Sub

Erro_objEventoLocalizacao_evSelecao:

    Select Case gErr

        Case 198759


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198760)

    End Select

    Exit Sub

End Sub

Private Sub LabelLocalizacao_Click()

Dim lErro As Long
Dim objRastreamentoLoteLoc As New ClassRastreamentoLoteLoc
Dim colSelecao As New Collection

On Error GoTo Erro_LabelLocalizacao_Click

    'Verifica se o Localizacao foi preenchido
    If Len(Trim(Localizacao.Text)) <> 0 Then

        objRastreamentoLoteLoc.sLocalizacao = Localizacao.Text

    End If

    Call Chama_Tela("RastreamentoLoteLocLista", colSelecao, objRastreamentoLoteLoc, objEventoLocalizacao)

    Exit Sub

Erro_LabelLocalizacao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198761)

    End Select

    Exit Sub

End Sub
