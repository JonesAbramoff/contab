Version 5.0
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FamiliasTipoInfo
   ClientHeight = 6000
   ClientLeft = 0
   ClientTop = 0
   ClientWidth = 9510
   KeyPreview = -1         'True
   ScaleHeight = 5745
   ScaleWidth = 8145
   Begin VB.PictureBox Picture1
      Height = 510
      Left = 7320
      ScaleHeight = 450
      ScaleWidth = 2025
      TabIndex = 0
      TabStop = 0             'False
      Top = 30
      Width = 2085
      Begin VB.CommandButton BotaoGravar
         Height = 360
         Left = 60
         Picture         =   "FamiliasTipoInfo.ctx":0000
         Style = 1              'Graphical
         TabIndex = 1
         ToolTipText = "Gravar"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoExcluir
         Height = 360
         Left = 570
         Picture         =   "FamiliasTipoInfo.ctx":015A
         Style = 1              'Graphical
         TabIndex = 2
         ToolTipText = "Excluir"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoLimpar
         Height = 360
         Left = 1065
         Picture         =   "FamiliasTipoInfo.ctx":02E4
         Style = 1              'Graphical
         TabIndex = 3
         ToolTipText = "Limpar"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoFechar
         Height = 360
         Left = 1545
         Picture         =   "FamiliasTipoInfo.ctx":0816
         Style = 1              'Graphical
         TabIndex = 4
         ToolTipText = "Fechar"
         Top = 45
         Width = 420
      End
   End
   Begin MSMask.MaskEdBox CodInfo
      Height          =   315
      Left            =   2000
      TabIndex        =   6
      Top             =   300
      Width           =   550
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodInfo
      Alignment       =   1  'Right Justify
      Caption         =   "CodInfo:"
      BeginProperty Font
         Name            = "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0              'False
         Italic          =   0              'False
         Strikethrough   =   0              'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   375
      MousePointer    = 14       'Arrow and Question
      TabIndex        = 7
      Top             = 325
      Width           = 1500
   End
   Begin MSMask.MaskEdBox Descricao
      Height          =   315
      Left            =   2000
      TabIndex        =   8
      Top             =   750
      Width           =   5500
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin VB.Label LabelDescricao
      Alignment       =   1  'Right Justify
      Caption         =   "Descricao:"
      BeginProperty Font
         Name            = "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0              'False
         Italic          =   0              'False
         Strikethrough   =   0              'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        = 9
      Top             = 775
      Width           = 1500
   End
   Begin MSMask.MaskEdBox Sigla
      Height          =   315
      Left            =   2000
      TabIndex        =   10
      Top             =   1200
      Width           =   1100
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   " "
   End
   Begin VB.Label LabelSigla
      Alignment       =   1  'Right Justify
      Caption         =   "Sigla:"
      BeginProperty Font
         Name            = "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0              'False
         Italic          =   0              'False
         Strikethrough   =   0              'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        = 11
      Top             = 1225
      Width           = 1500
   End
   Begin MSMask.MaskEdBox ValidoPara
      Height          =   315
      Left            =   2000
      TabIndex        =   12
      Top             =   1650
      Width           =   550
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin VB.Label LabelValidoPara
      Alignment       =   1  'Right Justify
      Caption         =   "ValidoPara:"
      BeginProperty Font
         Name            = "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0              'False
         Italic          =   0              'False
         Strikethrough   =   0              'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        = 13
      Top             = 1675
      Width           = 1500
   End
   Begin MSMask.MaskEdBox Posicao
      Height          =   315
      Left            =   2000
      TabIndex        =   14
      Top             =   2100
      Width           =   550
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin VB.Label LabelPosicao
      Alignment       =   1  'Right Justify
      Caption         =   "Posicao:"
      BeginProperty Font
         Name            = "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0              'False
         Italic          =   0              'False
         Strikethrough   =   0              'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        = 15
      Top             = 2125
      Width           = 1500
   End
End
Attribute VB_Name = "FamiliasTipoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodInfo As AdmEvento

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Familias Tipo Info"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "FamiliasTipoInfo"

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

Sub Form_UnLoad(Cancel as Integer)

Dim lErro As Long

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodInfo = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160102)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodInfo = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160103)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objFamiliasTipoInfo AS ClassFamiliasTipoInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objFamiliasTipoInfo Is Nothing) Then

        lErro = Traz_FamiliasTipoInfo_Tela(objFamiliasTipoInfo)
        If lErro <> SUCESSO Then gError 130571

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 130571

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160104)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objFamiliasTipoInfo AS ClassFamiliasTipoInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objFamiliasTipoInfo.iCodInfo = StrParaInt(CodInfo.text)
    objFamiliasTipoInfo.sDescricao = Descricao.text
    objFamiliasTipoInfo.sSigla = Sigla.text
    objFamiliasTipoInfo.iValidoPara = StrParaInt(ValidoPara.text)
    objFamiliasTipoInfo.iPosicao = StrParaInt(Posicao.text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160105)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objFamiliasTipoInfo As New ClassFamiliasTipoInfo

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FamiliasTipoInfo"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objFamiliasTipoInfo)
    If lErro <> SUCESSO Then gError 130572

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodInfo", objFamiliasTipoInfo.iCodInfo, 0, "CodInfo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 130572

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160106)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objFamiliasTipoInfo As New ClassFamiliasTipoInfo

On Error GoTo Erro_Tela_Preenche

    objFamiliasTipoInfo.iCodInfo = colCampoValor.Item("CodInfo").vValor

    If objFamiliasTipoInfo.iCodInfo<> 0Then
        lErro = Traz_FamiliasTipoInfo_Tela(objFamiliasTipoInfo)
        If lErro <> SUCESSO Then gError 130573
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 130573

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160107)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFamiliasTipoInfo As New ClassFamiliasTipoInfo

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodInfo.Text)) =0 then gError 130574
    '#####################

    'Preenche o objFamiliasTipoInfo
    lErro = Move_Tela_Memoria(objFamiliasTipoInfo)
    If lErro <> SUCESSO Then gError 130575

    lErro = Trata_Alteracao(objFamiliasTipoInfo, objFamiliasTipoInfo.iCodInfo)
    If lErro <> SUCESSO Then gError 130576

    'Grava o/a FamiliasTipoInfo no Banco de Dados
    lErro = CF("FamiliasTipoInfo_Grava", objFamiliasTipoInfo)
    If lErro <> SUCESSO Then gError 130577

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130574
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODINFO_FAMILIASTIPOINFO_NAO_PREENCHIDO">, gErr)
            CodInfo.SetFocus

        Case 130575, 130576, 130577

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160108)

    End Select

    Exit Function

End Function

Function Limpa_Tela_FamiliasTipoInfo() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_FamiliasTipoInfo

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_FamiliasTipoInfo = SUCESSO

    Exit Function

Erro_Limpa_Tela_FamiliasTipoInfo:

    Limpa_Tela_FamiliasTipoInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160109)

    End Select

    Exit Function

End Function

Function Traz_FamiliasTipoInfo_Tela(objFamiliasTipoInfo AS ClassFamiliasTipoInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_FamiliasTipoInfo_Tela

    'Lê o FamiliasTipoInfo que está sendo Passado
    lErro = CF("FamiliasTipoInfo_Le", objFamiliasTipoInfo)
    If lErro <> SUCESSO AND lErro <> 130552 Then gError 130578

    If lErro = SUCESSO Then 

        If objFamiliasTipoInfo.iCodInfo <> 0 Then CodInfo.text = Cstr(objFamiliasTipoInfo.iCodInfo)
        Descricao.text = objFamiliasTipoInfo.sDescricao
        Sigla.text = objFamiliasTipoInfo.sSigla
        If objFamiliasTipoInfo.iValidoPara <> 0 Then ValidoPara.text = Cstr(objFamiliasTipoInfo.iValidoPara)
        If objFamiliasTipoInfo.iPosicao <> 0 Then Posicao.text = Cstr(objFamiliasTipoInfo.iPosicao)

    End If 

    Traz_FamiliasTipoInfo_Tela = SUCESSO

    Exit Function

Erro_Traz_FamiliasTipoInfo_Tela:

    Traz_FamiliasTipoInfo_Tela = gErr

    Select Case gErr

        Case 130578

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160110)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 130579

    'Limpa Tela
    Call Limpa_Tela_FamiliasTipoInfo

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 130579

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160111)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160112)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 130580

    Call Limpa_Tela_FamiliasTipoInfo

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 130580

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160113)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFamiliasTipoInfo As New ClassFamiliasTipoInfo
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodInfo.Text)) =0 then gError 130581
    '#####################

    objFamiliasTipoInfo.iCodInfo = StrParaInt(CodInfo.text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_FAMILIASTIPOINFO", objFamiliasTipoInfo.iCodInfo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("FamiliasTipoInfo_Exclui", objFamiliasTipoInfo)
        If lErro <> SUCESSO Then gError 130582

        'Limpa Tela
        Call Limpa_Tela_FamiliasTipoInfo

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130581
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODINFO_FAMILIASTIPOINFO_NAO_PREENCHIDO">, gErr)
            CodInfo.SetFocus

        Case 130582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160114)

    End Select

    Exit Sub

End Sub

Private Sub CodInfo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodInfo_Validate

    'Verifica se CodInfo está preenchida
    If Len(Trim(CodInfo.Text)) <> 0 Then 

       'Critica a CodInfo
       lErro = Inteiro_Critica(CodInfo.Text)
       If lErro <> SUCESSO Then gError 130583

    End If

    Exit Sub

Erro_CodInfo_Validate:

    Cancel = True

    Select Case gErr

        Case 130583

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160115)

    End Select

    Exit Sub

End Sub

Private Sub CodInfo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodInfo, iAlterado)
    
End Sub

Private Sub CodInfo_Change()

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160116)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sigla_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Sigla_Validate

    'Verifica se Sigla está preenchida
    If Len(Trim(Sigla.Text)) <> 0 Then 

       '#######################################
       'CRITICA Sigla
       '#######################################

    End If

    Exit Sub

Erro_Sigla_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160117)

    End Select

    Exit Sub

End Sub

Private Sub Sigla_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValidoPara_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValidoPara_Validate

    'Verifica se ValidoPara está preenchida
    If Len(Trim(ValidoPara.Text)) <> 0 Then 

       'Critica a ValidoPara
       lErro = Inteiro_Critica(ValidoPara.Text)
       If lErro <> SUCESSO Then gError 130584

    End If

    Exit Sub

Erro_ValidoPara_Validate:

    Cancel = True

    Select Case gErr

        Case 130584

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160118)

    End Select

    Exit Sub

End Sub

Private Sub ValidoPara_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ValidoPara, iAlterado)
    
End Sub

Private Sub ValidoPara_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Posicao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Posicao_Validate

    'Verifica se Posicao está preenchida
    If Len(Trim(Posicao.Text)) <> 0 Then 

       'Critica a Posicao
       lErro = Inteiro_Critica(Posicao.Text)
       If lErro <> SUCESSO Then gError 130585

    End If

    Exit Sub

Erro_Posicao_Validate:

    Cancel = True

    Select Case gErr

        Case 130585

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160119)

    End Select

    Exit Sub

End Sub

Private Sub Posicao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Posicao, iAlterado)
    
End Sub

Private Sub Posicao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodInfo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFamiliasTipoInfo As ClassFamiliasTipoInfo

On Error GoTo Erro_objEventoCodInfo_evSelecao

    Set objFamiliasTipoInfo = obj1

    'Mostra os dados do FamiliasTipoInfo na tela
    lErro = Traz_FamiliasTipoInfo_Tela(objFamiliasTipoInfo)
    If lErro <> SUCESSO Then gError 130586

    Me.Show

    Exit Sub

Erro_objEventoCodInfo_evSelecao:

    Select Case gErr

        Case 130586


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160120)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodInfo_Click()

Dim lErro As Long
Dim objFamiliasTipoInfo As New ClassFamiliasTipoInfo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodInfo_Click

    'Verifica se o CodInfo foi preenchido
    If Len(Trim(CodInfo.Text)) <> 0 Then

        objFamiliasTipoInfo.iCodInfo= CodInfo.Text

    End If

    Call Chama_Tela("FamiliasTipoInfoLista", colSelecao, objFamiliasTipoInfo, objEventoCodInfo)

    Exit Sub

Erro_LabelCodInfo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160121)

    End Select

    Exit Sub

End Sub
