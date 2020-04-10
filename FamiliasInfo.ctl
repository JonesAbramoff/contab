Version 5.0
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FamiliasInfo
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
         Picture         =   "FamiliasInfo.ctx":0000
         Style = 1              'Graphical
         TabIndex = 1
         ToolTipText = "Gravar"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoExcluir
         Height = 360
         Left = 570
         Picture         =   "FamiliasInfo.ctx":015A
         Style = 1              'Graphical
         TabIndex = 2
         ToolTipText = "Excluir"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoLimpar
         Height = 360
         Left = 1065
         Picture         =   "FamiliasInfo.ctx":02E4
         Style = 1              'Graphical
         TabIndex = 3
         ToolTipText = "Limpar"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoFechar
         Height = 360
         Left = 1545
         Picture         =   "FamiliasInfo.ctx":0816
         Style = 1              'Graphical
         TabIndex = 4
         ToolTipText = "Fechar"
         Top = 45
         Width = 420
      End
   End
   Begin MSMask.MaskEdBox CodFamilia
      Height          =   315
      Left            =   2000
      TabIndex        =   6
      Top             =   300
      Width           =   880
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodFamilia
      Alignment       =   1  'Right Justify
      Caption         =   "CodFamilia:"
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
   Begin MSMask.MaskEdBox Seq
      Height          =   315
      Left            =   2000
      TabIndex        =   8
      Top             =   750
      Width           =   550
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin VB.Label LabelSeq
      Alignment       =   1  'Right Justify
      Caption         =   "Seq:"
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
      TabIndex        = 9
      Top             = 775
      Width           = 1500
   End
   Begin MSMask.MaskEdBox CodInfo
      Height          =   315
      Left            =   2000
      TabIndex        =   10
      Top             =   1200
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
      TabIndex        = 11
      Top             = 1225
      Width           = 1500
   End
   Begin MSMask.MaskEdBox Valor
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
   Begin VB.Label LabelValor
      Alignment       =   1  'Right Justify
      Caption         =   "Valor:"
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
End
Attribute VB_Name = "FamiliasInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodFamilia As AdmEvento

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Familias Info"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "FamiliasInfo"

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

    Set objEventoCodFamilia = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159985)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodFamilia = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159986)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objFamiliasInfo AS ClassFamiliasInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objFamiliasInfo Is Nothing) Then

        lErro = Traz_FamiliasInfo_Tela(objFamiliasInfo)
        If lErro <> SUCESSO Then gError 130527

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 130527

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159987)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objFamiliasInfo AS ClassFamiliasInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objFamiliasInfo.lCodFamilia = StrParaLong(CodFamilia.text)
    objFamiliasInfo.iSeq = StrParaInt(Seq.text)
    objFamiliasInfo.iCodInfo = StrParaInt(CodInfo.text)
    objFamiliasInfo.iValor = StrParaInt(Valor.text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159988)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objFamiliasInfo As New ClassFamiliasInfo

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FamiliasInfo"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objFamiliasInfo)
    If lErro <> SUCESSO Then gError 130528

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodFamilia", objFamiliasInfo.lCodFamilia, 0, "CodFamilia"
    colCampoValor.Add "Seq", objFamiliasInfo.iSeq, 0, "Seq"
    colCampoValor.Add "CodInfo", objFamiliasInfo.iCodInfo, 0, "CodInfo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 130528

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159989)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objFamiliasInfo As New ClassFamiliasInfo

On Error GoTo Erro_Tela_Preenche

    objFamiliasInfo.lCodFamilia = colCampoValor.Item("CodFamilia").vValor
    objFamiliasInfo.iSeq = colCampoValor.Item("Seq").vValor
    objFamiliasInfo.iCodInfo = colCampoValor.Item("CodInfo").vValor

    If objFamiliasInfo.lCodFamilia<> 0 AND objFamiliasInfo.iSeq<> 0 AND objFamiliasInfo.iCodInfo<> 0Then
        lErro = Traz_FamiliasInfo_Tela(objFamiliasInfo)
        If lErro <> SUCESSO Then gError 130529
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 130529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159990)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFamiliasInfo As New ClassFamiliasInfo

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) =0 then gError 130530
    If Len(Trim(Seq.Text)) =0 then gError 130531
    If Len(Trim(CodInfo.Text)) =0 then gError 130532
    '#####################

    'Preenche o objFamiliasInfo
    lErro = Move_Tela_Memoria(objFamiliasInfo)
    If lErro <> SUCESSO Then gError 130533

    lErro = Trata_Alteracao(objFamiliasInfo, objFamiliasInfo.lCodFamilia, objFamiliasInfo.iSeq, objFamiliasInfo.iCodInfo)
    If lErro <> SUCESSO Then gError 130534

    'Grava o/a FamiliasInfo no Banco de Dados
    lErro = CF("FamiliasInfo_Grava", objFamiliasInfo)
    If lErro <> SUCESSO Then gError 130535

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130530
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODFAMILIA_FAMILIASINFO_NAO_PREENCHIDO">, gErr)
            CodFamilia.SetFocus

        Case 130531
            Call Rotina_Erro(vbOKOnly, <"ERRO_SEQ_FAMILIASINFO_NAO_PREENCHIDO">, gErr)
            Seq.SetFocus

        Case 130532
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODINFO_FAMILIASINFO_NAO_PREENCHIDO">, gErr)
            CodInfo.SetFocus

        Case 130533, 130534, 130535

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159991)

    End Select

    Exit Function

End Function

Function Limpa_Tela_FamiliasInfo() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_FamiliasInfo

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_FamiliasInfo = SUCESSO

    Exit Function

Erro_Limpa_Tela_FamiliasInfo:

    Limpa_Tela_FamiliasInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159992)

    End Select

    Exit Function

End Function

Function Traz_FamiliasInfo_Tela(objFamiliasInfo AS ClassFamiliasInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_FamiliasInfo_Tela

    'Lê o FamiliasInfo que está sendo Passado
    lErro = CF("FamiliasInfo_Le", objFamiliasInfo)
    If lErro <> SUCESSO AND lErro <> 130508 Then gError 130536

    If lErro = SUCESSO Then 

        If objFamiliasInfo.lCodFamilia <> 0 Then CodFamilia.text = Cstr(objFamiliasInfo.lCodFamilia)
        If objFamiliasInfo.iSeq <> 0 Then Seq.text = Cstr(objFamiliasInfo.iSeq)
        If objFamiliasInfo.iCodInfo <> 0 Then CodInfo.text = Cstr(objFamiliasInfo.iCodInfo)
        If objFamiliasInfo.iValor <> 0 Then Valor.text = Cstr(objFamiliasInfo.iValor)

    End If 

    Traz_FamiliasInfo_Tela = SUCESSO

    Exit Function

Erro_Traz_FamiliasInfo_Tela:

    Traz_FamiliasInfo_Tela = gErr

    Select Case gErr

        Case 130536

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159993)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 130537

    'Limpa Tela
    Call Limpa_Tela_FamiliasInfo

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 130537

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159994)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159995)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 130538

    Call Limpa_Tela_FamiliasInfo

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 130538

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159996)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFamiliasInfo As New ClassFamiliasInfo
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) =0 then gError 130539
    If Len(Trim(Seq.Text)) =0 then gError 130540
    If Len(Trim(CodInfo.Text)) =0 then gError 130541
    '#####################

    objFamiliasInfo.lCodFamilia = StrParaLong(CodFamilia.text)
    objFamiliasInfo.iSeq = StrParaInt(Seq.text)
    objFamiliasInfo.iCodInfo = StrParaInt(CodInfo.text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_FAMILIASINFO", objFamiliasInfo.iCodInfo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("FamiliasInfo_Exclui", objFamiliasInfo)
        If lErro <> SUCESSO Then gError 130542

        'Limpa Tela
        Call Limpa_Tela_FamiliasInfo

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130539
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODFAMILIA_FAMILIASINFO_NAO_PREENCHIDO">, gErr)
            CodFamilia.SetFocus

        Case 130540
            Call Rotina_Erro(vbOKOnly, <"ERRO_SEQ_FAMILIASINFO_NAO_PREENCHIDO">, gErr)
            Seq.SetFocus

        Case 130541
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODINFO_FAMILIASINFO_NAO_PREENCHIDO">, gErr)
            CodInfo.SetFocus

        Case 130542

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159997)

    End Select

    Exit Sub

End Sub

Private Sub CodFamilia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodFamilia_Validate

    'Verifica se CodFamilia está preenchida
    If Len(Trim(CodFamilia.Text)) <> 0 Then 

       'Critica a CodFamilia
       lErro = Long_Critica(CodFamilia.Text)
       If lErro <> SUCESSO Then gError 130543

    End If

    Exit Sub

Erro_CodFamilia_Validate:

    Cancel = True

    Select Case gErr

        Case 130543

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159998)

    End Select

    Exit Sub

End Sub

Private Sub CodFamilia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodFamilia, iAlterado)
    
End Sub

Private Sub CodFamilia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Seq_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Seq_Validate

    'Verifica se Seq está preenchida
    If Len(Trim(Seq.Text)) <> 0 Then 

       'Critica a Seq
       lErro = Inteiro_Critica(Seq.Text)
       If lErro <> SUCESSO Then gError 130544

    End If

    Exit Sub

Erro_Seq_Validate:

    Cancel = True

    Select Case gErr

        Case 130544

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159999)

    End Select

    Exit Sub

End Sub

Private Sub Seq_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Seq, iAlterado)
    
End Sub

Private Sub Seq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodInfo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodInfo_Validate

    'Verifica se CodInfo está preenchida
    If Len(Trim(CodInfo.Text)) <> 0 Then 

       'Critica a CodInfo
       lErro = Inteiro_Critica(CodInfo.Text)
       If lErro <> SUCESSO Then gError 130545

    End If

    Exit Sub

Erro_CodInfo_Validate:

    Cancel = True

    Select Case gErr

        Case 130545

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160000)

    End Select

    Exit Sub

End Sub

Private Sub CodInfo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodInfo, iAlterado)
    
End Sub

Private Sub CodInfo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se Valor está preenchida
    If Len(Trim(Valor.Text)) <> 0 Then 

       'Critica a Valor
       lErro = Inteiro_Critica(Valor.Text)
       If lErro <> SUCESSO Then gError 130546

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 130546

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160001)

    End Select

    Exit Sub

End Sub

Private Sub Valor_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Valor, iAlterado)
    
End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodFamilia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFamiliasInfo As ClassFamiliasInfo

On Error GoTo Erro_objEventoCodFamilia_evSelecao

    Set objFamiliasInfo = obj1

    'Mostra os dados do FamiliasInfo na tela
    lErro = Traz_FamiliasInfo_Tela(objFamiliasInfo)
    If lErro <> SUCESSO Then gError 130547

    Me.Show

    Exit Sub

Erro_objEventoCodFamilia_evSelecao:

    Select Case gErr

        Case 130547


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160002)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodFamilia_Click()

Dim lErro As Long
Dim objFamiliasInfo As New ClassFamiliasInfo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodFamilia_Click

    'Verifica se o CodFamilia foi preenchido
    If Len(Trim(CodFamilia.Text)) <> 0 Then

        objFamiliasInfo.lCodFamilia= CodFamilia.Text

    End If

    Call Chama_Tela("FamiliasInfoLista", colSelecao, objFamiliasInfo, objEventoCodFamilia)

    Exit Sub

Erro_LabelCodFamilia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160003)

    End Select

    Exit Sub

End Sub

Private Sub LabelSeq_Click()

Dim lErro As Long
Dim objFamiliasInfo As New ClassFamiliasInfo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSeq_Click

    'Verifica se o Seq foi preenchido
    If Len(Trim(Seq.Text)) <> 0 Then

        objFamiliasInfo.iSeq= Seq.Text

    End If

    Call Chama_Tela("FamiliasInfoLista", colSelecao, objFamiliasInfo, objEventoCodFamilia)

    Exit Sub

Erro_LabelSeq_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160004)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodInfo_Click()

Dim lErro As Long
Dim objFamiliasInfo As New ClassFamiliasInfo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodInfo_Click

    'Verifica se o CodInfo foi preenchido
    If Len(Trim(CodInfo.Text)) <> 0 Then

        objFamiliasInfo.iCodInfo= CodInfo.Text

    End If

    Call Chama_Tela("FamiliasInfoLista", colSelecao, objFamiliasInfo, objEventoCodFamilia)

    Exit Sub

Erro_LabelCodInfo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160005)

    End Select

    Exit Sub

End Sub
