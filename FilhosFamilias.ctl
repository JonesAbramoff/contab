Version 5.0
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FilhosFamilias
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
         Picture         =   "FilhosFamilias.ctx":0000
         Style = 1              'Graphical
         TabIndex = 1
         ToolTipText = "Gravar"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoExcluir
         Height = 360
         Left = 570
         Picture         =   "FilhosFamilias.ctx":015A
         Style = 1              'Graphical
         TabIndex = 2
         ToolTipText = "Excluir"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoLimpar
         Height = 360
         Left = 1065
         Picture         =   "FilhosFamilias.ctx":02E4
         Style = 1              'Graphical
         TabIndex = 3
         ToolTipText = "Limpar"
         Top = 45
         Width = 420
      End
      Begin VB.CommandButton BotaoFechar
         Height = 360
         Left = 1545
         Picture         =   "FilhosFamilias.ctx":0816
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
   Begin MSMask.MaskEdBox SeqFilho
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
   Begin VB.Label LabelSeqFilho
      Alignment       =   1  'Right Justify
      Caption         =   "SeqFilho:"
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
   Begin MSMask.MaskEdBox Nome
      Height          =   315
      Left            =   2000
      TabIndex        =   10
      Top             =   1200
      Width           =   4400
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin VB.Label LabelNome
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
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
   Begin MSMask.MaskEdBox NomeHebr
      Height          =   315
      Left            =   2000
      TabIndex        =   12
      Top             =   1650
      Width           =   4400
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin VB.Label LabelNomeHebr
      Alignment       =   1  'Right Justify
      Caption         =   "NomeHebr:"
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
   Begin MSMask.MaskEdBox DtNasc
      Height          =   315
      Left            =   2000
      TabIndex        =   14
      Top             =   2100
      Width           =   1300
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      MaxLength      =   8
      Format         =   "dd/mm/yy"
      Mask           =   "##/##/##"
      PromptChar     =   "_"
   End
   Begin MSComCtl2.UpDown UpDownDtNasc
      Height          =   300
      Left            =   3310
      TabIndex        =   15
      TabStop         =   0             'False
      Top             =   2100
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1            'True
   End
   Begin VB.Label LabelDtNasc
      Alignment       =   1  'Right Justify
      Caption         =   "DtNasc:"
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
      TabIndex        = 16
      Top             = 2125
      Width           = 1500
   End
   Begin MSMask.MaskEdBox DtNascNoite
      Height          =   315
      Left            =   2000
      TabIndex        =   17
      Top             =   2550
      Width           =   550
      _ExtentX        =   2699
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin VB.Label LabelDtNascNoite
      Alignment       =   1  'Right Justify
      Caption         =   "DtNascNoite:"
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
      TabIndex        = 18
      Top             = 2575
      Width           = 1500
   End
End
Attribute VB_Name = "FilhosFamilias"
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
    Caption = "Filhos das Familias"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "FilhosFamilias"

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160228)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160229)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objFilhosFamilias AS ClassFilhosFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objFilhosFamilias Is Nothing) Then

        lErro = Traz_FilhosFamilias_Tela(objFilhosFamilias)
        If lErro <> SUCESSO Then gError 130610

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 130610

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160230)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objFilhosFamilias AS ClassFilhosFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objFilhosFamilias.lCodFamilia = StrParaLong(CodFamilia.text)
    objFilhosFamilias.iSeqFilho = StrParaInt(SeqFilho.text)
    objFilhosFamilias.sNome = Nome.text
    objFilhosFamilias.sNomeHebr = NomeHebr.text
    if len(trim(DtNasc.ClipText))<>0 then objFilhosFamilias.dtDtNasc = Format(DtNasc.text, DtNasc.Format)
    objFilhosFamilias.iDtNascNoite = StrParaInt(DtNascNoite.text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160231)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objFilhosFamilias As New ClassFilhosFamilias

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FilhosFamilias"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objFilhosFamilias)
    If lErro <> SUCESSO Then gError 130611

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodFamilia", objFilhosFamilias.lCodFamilia, 0, "CodFamilia"
    colCampoValor.Add "SeqFilho", objFilhosFamilias.iSeqFilho, 0, "SeqFilho"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 130611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160232)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objFilhosFamilias As New ClassFilhosFamilias

On Error GoTo Erro_Tela_Preenche

    objFilhosFamilias.lCodFamilia = colCampoValor.Item("CodFamilia").vValor
    objFilhosFamilias.iSeqFilho = colCampoValor.Item("SeqFilho").vValor

    If objFilhosFamilias.lCodFamilia<> 0 AND objFilhosFamilias.iSeqFilho<> 0Then
        lErro = Traz_FilhosFamilias_Tela(objFilhosFamilias)
        If lErro <> SUCESSO Then gError 130612
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 130612

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160233)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFilhosFamilias As New ClassFilhosFamilias

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) =0 then gError 130613
    If Len(Trim(SeqFilho.Text)) =0 then gError 130614
    '#####################

    'Preenche o objFilhosFamilias
    lErro = Move_Tela_Memoria(objFilhosFamilias)
    If lErro <> SUCESSO Then gError 130615

    lErro = Trata_Alteracao(objFilhosFamilias, objFilhosFamilias.lCodFamilia, objFilhosFamilias.iSeqFilho)
    If lErro <> SUCESSO Then gError 130616

    'Grava o/a FilhosFamilias no Banco de Dados
    lErro = CF("FilhosFamilias_Grava", objFilhosFamilias)
    If lErro <> SUCESSO Then gError 130617

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130613
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODFAMILIA_FILHOSFAMILIAS_NAO_PREENCHIDO">, gErr)
            CodFamilia.SetFocus

        Case 130614
            Call Rotina_Erro(vbOKOnly, <"ERRO_SEQFILHO_FILHOSFAMILIAS_NAO_PREENCHIDO">, gErr)
            SeqFilho.SetFocus

        Case 130615, 130616, 130617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160234)

    End Select

    Exit Function

End Function

Function Limpa_Tela_FilhosFamilias() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_FilhosFamilias

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_FilhosFamilias = SUCESSO

    Exit Function

Erro_Limpa_Tela_FilhosFamilias:

    Limpa_Tela_FilhosFamilias = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160235)

    End Select

    Exit Function

End Function

Function Traz_FilhosFamilias_Tela(objFilhosFamilias AS ClassFilhosFamilias) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_FilhosFamilias_Tela

    'Lê o FilhosFamilias que está sendo Passado
    lErro = CF("FilhosFamilias_Le", objFilhosFamilias)
    If lErro <> SUCESSO AND lErro <> 130591 Then gError 130618

    If lErro = SUCESSO Then 

        If objFilhosFamilias.lCodFamilia <> 0 Then CodFamilia.text = Cstr(objFilhosFamilias.lCodFamilia)
        If objFilhosFamilias.iSeqFilho <> 0 Then SeqFilho.text = Cstr(objFilhosFamilias.iSeqFilho)
        Nome.text = objFilhosFamilias.sNome
        NomeHebr.text = objFilhosFamilias.sNomeHebr

        If objFilhosFamilias.dtDtNasc <> 0 Then 
            DtNasc.PromptInclude = False 
            DtNasc.text = Format(objFilhosFamilias.dtDtNasc,"dd/mm/yy")
            DtNasc.PromptInclude = True 
        End If

        If objFilhosFamilias.iDtNascNoite <> 0 Then DtNascNoite.text = Cstr(objFilhosFamilias.iDtNascNoite)

    End If 

    Traz_FilhosFamilias_Tela = SUCESSO

    Exit Function

Erro_Traz_FilhosFamilias_Tela:

    Traz_FilhosFamilias_Tela = gErr

    Select Case gErr

        Case 130618

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160236)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 130619

    'Limpa Tela
    Call Limpa_Tela_FilhosFamilias

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 130619

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160237)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160238)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 130620

    Call Limpa_Tela_FilhosFamilias

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 130620

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160239)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFilhosFamilias As New ClassFilhosFamilias
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(CodFamilia.Text)) =0 then gError 130621
    If Len(Trim(SeqFilho.Text)) =0 then gError 130622
    '#####################

    objFilhosFamilias.lCodFamilia = StrParaLong(CodFamilia.text)
    objFilhosFamilias.iSeqFilho = StrParaInt(SeqFilho.text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_FILHOSFAMILIAS", objFilhosFamilias.iSeqFilho)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("FilhosFamilias_Exclui", objFilhosFamilias)
        If lErro <> SUCESSO Then gError 130623

        'Limpa Tela
        Call Limpa_Tela_FilhosFamilias

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130621
            Call Rotina_Erro(vbOKOnly, <"ERRO_CODFAMILIA_FILHOSFAMILIAS_NAO_PREENCHIDO">, gErr)
            CodFamilia.SetFocus

        Case 130622
            Call Rotina_Erro(vbOKOnly, <"ERRO_SEQFILHO_FILHOSFAMILIAS_NAO_PREENCHIDO">, gErr)
            SeqFilho.SetFocus

        Case 130623

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160240)

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
       If lErro <> SUCESSO Then gError 130624

    End If

    Exit Sub

Erro_CodFamilia_Validate:

    Cancel = True

    Select Case gErr

        Case 130624

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160241)

    End Select

    Exit Sub

End Sub

Private Sub CodFamilia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodFamilia, iAlterado)
    
End Sub

Private Sub CodFamilia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SeqFilho_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SeqFilho_Validate

    'Verifica se SeqFilho está preenchida
    If Len(Trim(SeqFilho.Text)) <> 0 Then 

       'Critica a SeqFilho
       lErro = Inteiro_Critica(SeqFilho.Text)
       If lErro <> SUCESSO Then gError 130625

    End If

    Exit Sub

Erro_SeqFilho_Validate:

    Cancel = True

    Select Case gErr

        Case 130625

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160242)

    End Select

    Exit Sub

End Sub

Private Sub SeqFilho_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(SeqFilho, iAlterado)
    
End Sub

Private Sub SeqFilho_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Nome_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Nome_Validate

    'Verifica se Nome está preenchida
    If Len(Trim(Nome.Text)) <> 0 Then 

       '#######################################
       'CRITICA Nome
       '#######################################

    End If

    Exit Sub

Erro_Nome_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160243)

    End Select

    Exit Sub

End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeHebr_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeHebr_Validate

    'Verifica se NomeHebr está preenchida
    If Len(Trim(NomeHebr.Text)) <> 0 Then 

       '#######################################
       'CRITICA NomeHebr
       '#######################################

    End If

    Exit Sub

Erro_NomeHebr_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160244)

    End Select

    Exit Sub

End Sub

Private Sub NomeHebr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDtNasc_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDtNasc_DownClick

    DtNasc.SetFocus

    If Len(DtNasc.ClipText) > 0 Then

        sData = DtNasc.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 130626

        DtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownDtNasc_DownClick:

    Select Case gErr

        Case 130626

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160245)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDtNasc_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDtNasc_UpClick

    DtNasc.SetFocus

    If Len(Trim(DtNasc.ClipText)) > 0 Then

        sData = DtNasc.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 130627

        DtNasc.Text = sData

    End If

    Exit Sub

Erro_UpDownDtNasc_UpClick:

    Select Case gErr

        Case 130627

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160246)

    End Select

    Exit Sub

End Sub

Private Sub DtNasc_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DtNasc, iAlterado)
    
End Sub

Private Sub DtNasc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DtNasc_Validate

    If Len(Trim(DtNasc.ClipText)) <> 0 Then 

        lErro = Data_Critica(DtNasc.Text)
        If lErro <> SUCESSO Then gError 130628

    End If

    Exit Sub

Erro_DtNasc_Validate:

    Cancel = True

    Select Case gErr

        Case 130628

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160247)

    End Select

    Exit Sub

End Sub

Private Sub DtNasc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DtNascNoite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DtNascNoite_Validate

    'Verifica se DtNascNoite está preenchida
    If Len(Trim(DtNascNoite.Text)) <> 0 Then 

       'Critica a DtNascNoite
       lErro = Inteiro_Critica(DtNascNoite.Text)
       If lErro <> SUCESSO Then gError 130629

    End If

    Exit Sub

Erro_DtNascNoite_Validate:

    Cancel = True

    Select Case gErr

        Case 130629

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160248)

    End Select

    Exit Sub

End Sub

Private Sub DtNascNoite_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DtNascNoite, iAlterado)
    
End Sub

Private Sub DtNascNoite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodFamilia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFilhosFamilias As ClassFilhosFamilias

On Error GoTo Erro_objEventoCodFamilia_evSelecao

    Set objFilhosFamilias = obj1

    'Mostra os dados do FilhosFamilias na tela
    lErro = Traz_FilhosFamilias_Tela(objFilhosFamilias)
    If lErro <> SUCESSO Then gError 130630

    Me.Show

    Exit Sub

Erro_objEventoCodFamilia_evSelecao:

    Select Case gErr

        Case 130630


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160249)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodFamilia_Click()

Dim lErro As Long
Dim objFilhosFamilias As New ClassFilhosFamilias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodFamilia_Click

    'Verifica se o CodFamilia foi preenchido
    If Len(Trim(CodFamilia.Text)) <> 0 Then

        objFilhosFamilias.lCodFamilia= CodFamilia.Text

    End If

    Call Chama_Tela("FilhosFamiliasLista", colSelecao, objFilhosFamilias, objEventoCodFamilia)

    Exit Sub

Erro_LabelCodFamilia_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160250)

    End Select

    Exit Sub

End Sub

Private Sub LabelSeqFilho_Click()

Dim lErro As Long
Dim objFilhosFamilias As New ClassFilhosFamilias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSeqFilho_Click

    'Verifica se o SeqFilho foi preenchido
    If Len(Trim(SeqFilho.Text)) <> 0 Then

        objFilhosFamilias.iSeqFilho= SeqFilho.Text

    End If

    Call Chama_Tela("FilhosFamiliasLista", colSelecao, objFilhosFamilias, objEventoCodFamilia)

    Exit Sub

Erro_LabelSeqFilho_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160251)

    End Select

    Exit Sub

End Sub
