VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PVAndamentoOcx 
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   5145
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2115
      Picture         =   "PVAndamento.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Numeração Automática"
      Top             =   900
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      DrawStyle       =   1  'Dash
      Height          =   555
      Left            =   2925
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "PVAndamento.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "PVAndamento.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "PVAndamento.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1560
         Picture         =   "PVAndamento.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameAuto 
      Caption         =   "Automático"
      Enabled         =   0   'False
      Height          =   855
      Left            =   30
      TabIndex        =   6
      Top             =   2265
      Width           =   5055
      Begin VB.ComboBox FatorAuto 
         Height          =   315
         ItemData        =   "PVAndamento.ctx":0A7E
         Left            =   1440
         List            =   "PVAndamento.ctx":0A8E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   3540
      End
      Begin VB.Label Label1 
         Caption         =   "Alterar Quando:"
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
         Left            =   60
         TabIndex        =   8
         Top             =   405
         Width           =   1470
      End
   End
   Begin VB.OptionButton OptAuto 
      Caption         =   "Automático"
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
      Left            =   2490
      TabIndex        =   5
      Top             =   1905
      Width           =   1725
   End
   Begin VB.OptionButton OptManual 
      Caption         =   "Manual"
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
      Left            =   1035
      TabIndex        =   4
      Top             =   1905
      Value           =   -1  'True
      Width           =   1185
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1470
      TabIndex        =   0
      Top             =   900
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      Top             =   1350
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin VB.Label LabelDescricao 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1395
      Width           =   930
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   765
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   930
      Width           =   645
   End
End
Attribute VB_Name = "PVAndamentoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Andamento do PV"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PVAndamento"

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

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205648)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205649)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objPVAndamento Is Nothing) Then

        lErro = Traz_PVAndamento_Tela(objPVAndamento)
        If lErro <> SUCESSO Then gError 205650

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 205650

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205651)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objPVAndamento.iCodigo = StrParaInt(Codigo.Text)
    objPVAndamento.sDescricao = Descricao.Text
    If OptAuto.Value Then
        objPVAndamento.iAuto = MARCADO
    Else
        objPVAndamento.iAuto = DESMARCADO
    End If
    objPVAndamento.iFatorAuto = Codigo_Extrai(FatorAuto.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205652)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objPVAndamento As New ClassPVAndamento

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PVAndamento"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objPVAndamento)
    If lErro <> SUCESSO Then gError 205653

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objPVAndamento.iCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 205653

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205654)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objPVAndamento As New ClassPVAndamento

On Error GoTo Erro_Tela_Preenche

    objPVAndamento.iCodigo = colCampoValor.Item("Codigo").vValor

    If objPVAndamento.iCodigo <> 0 Then

        lErro = Traz_PVAndamento_Tela(objPVAndamento)
        If lErro <> SUCESSO Then gError 205655

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 205655

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205656)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPVAndamento As New ClassPVAndamento

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 205657
    '#####################

    'Preenche o objPVAndamento
    lErro = Move_Tela_Memoria(objPVAndamento)
    If lErro <> SUCESSO Then gError 205658

    lErro = Trata_Alteracao(objPVAndamento, objPVAndamento.iCodigo)
    If lErro <> SUCESSO Then gError 205659

    'Grava o/a PVAndamento no Banco de Dados
    lErro = CF("PVAndamento_Grava", objPVAndamento)
    If lErro <> SUCESSO Then gError 205660

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205657
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205658 To 205660

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205661)

    End Select

    Exit Function

End Function

Function Limpa_Tela_PVAndamento() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PVAndamento

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    OptManual.Value = True
    
    Call Trata_Auto

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_PVAndamento = SUCESSO

    Exit Function

Erro_Limpa_Tela_PVAndamento:

    Limpa_Tela_PVAndamento = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205662)

    End Select

    Exit Function

End Function

Function Traz_PVAndamento_Tela(objPVAndamento As ClassPVAndamento) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_PVAndamento_Tela

    Call Limpa_Tela_PVAndamento

    If objPVAndamento.iCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objPVAndamento.iCodigo)
        Codigo.PromptInclude = True
    End If

    'Lê o PVAndamento que está sendo Passado
    lErro = CF("PVAndamento_Le", objPVAndamento)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205663

    If lErro = SUCESSO Then


        If objPVAndamento.iCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objPVAndamento.iCodigo)
            Codigo.PromptInclude = True
        End If

        Descricao.Text = objPVAndamento.sDescricao

        If objPVAndamento.iAuto = MARCADO Then
            OptAuto.Value = True
        Else
            OptManual.Value = True
        End If

        Call Combo_Seleciona_ItemData(FatorAuto, objPVAndamento.iFatorAuto)

    End If

    iAlterado = 0

    Traz_PVAndamento_Tela = SUCESSO

    Exit Function

Erro_Traz_PVAndamento_Tela:

    Traz_PVAndamento_Tela = gErr

    Select Case gErr

        Case 205663

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205664)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 205665

    'Limpa Tela
    Call Limpa_Tela_PVAndamento

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205665

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205666)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205667)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 205668

    Call Limpa_Tela_PVAndamento

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 205668

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205669)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPVAndamento As New ClassPVAndamento
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 205670
    '#####################

    objPVAndamento.iCodigo = StrParaInt(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PVANDAMENTO", objPVAndamento.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("PVAndamento_Exclui", objPVAndamento)
        If lErro <> SUCESSO Then gError 205671

        'Limpa Tela
        Call Limpa_Tela_PVAndamento

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205670
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205671

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205672)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Inteiro_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 205673

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 205673

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205674)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205675)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FatorAuto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPVAndamento As ClassPVAndamento

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPVAndamento = obj1

    'Mostra os dados do PVAndamento na tela
    lErro = Traz_PVAndamento_Tela(objPVAndamento)
    If lErro <> SUCESSO Then gError 205676

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 205676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205677)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objPVAndamento As New ClassPVAndamento
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objPVAndamento.iCodigo = Codigo.Text

    End If

    Call Chama_Tela("PVAndamentoLista", colSelecao, objPVAndamento, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205678)

    End Select

    Exit Sub

End Sub

Private Sub Trata_Auto()

Dim lErro As Long

On Error GoTo Erro_Trata_Auto

    If OptAuto.Value Then
        FrameAuto.Enabled = True
        FatorAuto.ListIndex = 0
    Else
        FrameAuto.Enabled = False
        FatorAuto.ListIndex = -1
    End If

    Exit Sub

Erro_Trata_Auto:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205678)

    End Select

    Exit Sub

End Sub

Private Sub OptAuto_Click()
    Call Trata_Auto
End Sub

Private Sub OptManual_Click()
    Call Trata_Auto
End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_PVANDAMENTO", "PVAndamento", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 205679
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 205679

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205680)
    
    End Select

    Exit Sub
    
End Sub
