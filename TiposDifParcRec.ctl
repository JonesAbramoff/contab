VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TiposDifParcRecOcx 
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   KeyPreview      =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   7635
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2715
      Picture         =   "TiposDifParcRec.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   315
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5415
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TiposDifParcRec.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "TiposDifParcRec.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "TiposDifParcRec.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "TiposDifParcRec.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1995
      TabIndex        =   0
      Top             =   300
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      PromptChar      =   " "
   End
   Begin VB.TextBox Descricao 
      Height          =   315
      Left            =   2000
      MaxLength       =   150
      TabIndex        =   2
      Top             =   750
      Width           =   5500
   End
   Begin MSMask.MaskEdBox ContaContabilCR 
      Height          =   315
      Left            =   2595
      TabIndex        =   3
      Top             =   1230
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ContaContabilRecDesp 
      Height          =   315
      Left            =   2580
      TabIndex        =   4
      Top             =   1665
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   360
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   330
      Width           =   1500
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
      Height          =   315
      Left            =   375
      TabIndex        =   11
      Top             =   765
      Width           =   1500
   End
   Begin VB.Label LabelContaContabilCR 
      Alignment       =   1  'Right Justify
      Caption         =   "Conta de Contas a Receber:"
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
      Left            =   60
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   1275
      Width           =   2445
   End
   Begin VB.Label LabelContaContabilRecDesp 
      Alignment       =   1  'Right Justify
      Caption         =   "Conta de Receita/Despesa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   13
      Top             =   1695
      Width           =   2460
   End
End
Attribute VB_Name = "TiposDifParcRecOcx"
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
Private WithEvents objEventoContaContabilRecDesp As AdmEvento
Attribute objEventoContaContabilRecDesp.VB_VarHelpID = -1
Private WithEvents objEventoContaContabilCR As AdmEvento
Attribute objEventoContaContabilCR.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tipos de Diferenças nas parcelas a Receber"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TiposDifParcRec"

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
    Set objEventoContaContabilCR = Nothing
    Set objEventoContaContabilRecDesp = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177697)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoContaContabilCR = New AdmEvento
    Set objEventoContaContabilRecDesp = New AdmEvento

    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 177696
    
    ContaContabilCR.Mask = sMascaraConta
    ContaContabilRecDesp.Mask = sMascaraConta

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 177676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177698)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTiposDifParcRec As ClassTiposDifParcRec) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTiposDifParcRec Is Nothing) Then

        lErro = Traz_TiposDifParcRec_Tela(objTiposDifParcRec)
        If lErro <> SUCESSO Then gError 177677

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 177677

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177711)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objTiposDifParcRec As ClassTiposDifParcRec) As Long

Dim lErro As Long
Dim sContaCR As String
Dim iContaPreenchidaCR As Integer
Dim sConta As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objTiposDifParcRec.iCodigo = StrParaInt(Codigo.Text)
    objTiposDifParcRec.sDescricao = Descricao.Text
    
    lErro = CF("Conta_Formata", ContaContabilCR.Text, sContaCR, iContaPreenchidaCR)
    If lErro <> SUCESSO Then gError 177725
    
    If iContaPreenchidaCR = CONTA_VAZIA Then
        objTiposDifParcRec.sContaContabilCR = ""
    Else
        objTiposDifParcRec.sContaContabilCR = sContaCR
    End If
    
    lErro = CF("Conta_Formata", ContaContabilRecDesp.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 177726
    
    If iContaPreenchida = CONTA_VAZIA Then
        objTiposDifParcRec.sContaContabilRecDesp = ""
    Else
        objTiposDifParcRec.sContaContabilRecDesp = sConta
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 177725, 177726

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177712)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTiposDifParcRec As New ClassTiposDifParcRec

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposDifParcRec"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objTiposDifParcRec)
    If lErro <> SUCESSO Then gError 177678

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTiposDifParcRec.iCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 177678

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177713)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTiposDifParcRec As New ClassTiposDifParcRec

On Error GoTo Erro_Tela_Preenche

    objTiposDifParcRec.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTiposDifParcRec.iCodigo <> 0 Then
        lErro = Traz_TiposDifParcRec_Tela(objTiposDifParcRec)
        If lErro <> SUCESSO Then gError 177679
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 177679

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177714)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTiposDifParcRec As New ClassTiposDifParcRec

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 177680
    
    'Preenche o objTiposDifParcRec
    lErro = Move_Tela_Memoria(objTiposDifParcRec)
    If lErro <> SUCESSO Then gError 177681

    lErro = Trata_Alteracao(objTiposDifParcRec, objTiposDifParcRec.iCodigo)
    If lErro <> SUCESSO Then gError 177682

    'Grava o/a TiposDifParcRec no Banco de Dados
    lErro = CF("TiposDifParcRec_Grava", objTiposDifParcRec)
    If lErro <> SUCESSO Then gError 177683

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 177680
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDIFPARCREC_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 177681, 177682, 177683

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177715)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TiposDifParcRec() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TiposDifParcRec

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_TiposDifParcRec = SUCESSO

    Exit Function

Erro_Limpa_Tela_TiposDifParcRec:

    Limpa_Tela_TiposDifParcRec = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177716)

    End Select

    Exit Function

End Function

Function Traz_TiposDifParcRec_Tela(objTiposDifParcRec As ClassTiposDifParcRec) As Long

Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaEnxutaCR As String

On Error GoTo Erro_Traz_TiposDifParcRec_Tela

    'Lê o TiposDifParcRec que está sendo Passado
    lErro = CF("TiposDifParcRec_Le", objTiposDifParcRec)
    If lErro <> SUCESSO And lErro <> 177657 Then gError 177684

    If lErro = SUCESSO Then

        Codigo.Text = CStr(objTiposDifParcRec.iCodigo)
        
        Descricao.Text = objTiposDifParcRec.sDescricao
        
        If Len(Trim(objTiposDifParcRec.sContaContabilCR)) <> 0 Then
            lErro = Mascara_RetornaContaEnxuta(objTiposDifParcRec.sContaContabilCR, sContaEnxutaCR)
            If lErro <> SUCESSO Then gError 177728
    
            ContaContabilCR.PromptInclude = False
            ContaContabilCR.Text = sContaEnxutaCR
            ContaContabilCR.PromptInclude = True
        End If

        If Len(Trim(objTiposDifParcRec.sContaContabilRecDesp)) <> 0 Then
            lErro = Mascara_RetornaContaEnxuta(objTiposDifParcRec.sContaContabilRecDesp, sContaEnxuta)
            If lErro <> SUCESSO Then gError 177729
    
            ContaContabilRecDesp.PromptInclude = False
            ContaContabilRecDesp.Text = sContaEnxuta
            ContaContabilRecDesp.PromptInclude = True
        End If

    End If

    iAlterado = 0

    Traz_TiposDifParcRec_Tela = SUCESSO

    Exit Function

Erro_Traz_TiposDifParcRec_Tela:

    Traz_TiposDifParcRec_Tela = gErr

    Select Case gErr

        Case 177684, 177728, 177729

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177717)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 177685

    'Limpa Tela
    Call Limpa_Tela_TiposDifParcRec

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 177685

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177699)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177700)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 177686

    Call Limpa_Tela_TiposDifParcRec

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 177686

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177701)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTiposDifParcRec As New ClassTiposDifParcRec
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 177687

    objTiposDifParcRec.iCodigo = StrParaInt(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPOSDIFPARCREC", objTiposDifParcRec.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("TiposDifParcRec_Exclui", objTiposDifParcRec)
        If lErro <> SUCESSO Then gError 177688

        'Limpa Tela
        Call Limpa_Tela_TiposDifParcRec

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 177687
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDIFPARCREC_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 177688

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177702)

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
       If lErro <> SUCESSO Then gError 177689

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 177689

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177704)

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


    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177705)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaContabilCR_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabilCR_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabilCR.Text, ContaContabilCR.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 177730
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 177731
        
        ContaContabilCR.PromptInclude = False
        ContaContabilCR.Text = sContaMascarada
        ContaContabilCR.PromptInclude = True
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        lErro = CF("Conta_Critica", ContaContabilCR.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 177732
        
        If lErro = 5700 Then gError 177733

    End If
    
    Exit Sub

Erro_ContaContabilCR_Validate:

    Cancel = True

    Select Case gErr

        Case 177730, 177732
            
        Case 177731
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 177733
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabilCR.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177734)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabilCR_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaContabilRecDesp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabilRecDesp_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabilRecDesp.Text, ContaContabilRecDesp.ClipText, objPlanoConta, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 177735
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 177736
        
        ContaContabilRecDesp.PromptInclude = False
        ContaContabilRecDesp.Text = sContaMascarada
        ContaContabilRecDesp.PromptInclude = True
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then
    
        lErro = CF("Conta_Critica", ContaContabilRecDesp.Text, sContaFormatada, objPlanoConta, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 177737
        
        If lErro = 5700 Then gError 177738

    End If
    
    Exit Sub

Erro_ContaContabilRecDesp_Validate:

    Cancel = True

    Select Case gErr

        Case 177735, 177737
            
        Case 177736
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 177738
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabilRecDesp.Text)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177739)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabilRecDesp_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDifParcRec As ClassTiposDifParcRec

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTiposDifParcRec = obj1

    'Mostra os dados do TiposDifParcRec na tela
    lErro = Traz_TiposDifParcRec_Tela(objTiposDifParcRec)
    If lErro <> SUCESSO Then gError 177690

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 177690


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177708)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTiposDifParcRec As New ClassTiposDifParcRec
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTiposDifParcRec.iCodigo = StrParaInt(Codigo.Text)

    End If

    Call Chama_Tela("TiposDifParcRecLista", colSelecao, objTiposDifParcRec, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177709)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático para o código de Banco
    lErro = CF("TiposDifParcRec_Automatico", iCodigo)
    If lErro <> SUCESSO Then gError 177694

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 177694
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 177695)
    
    End Select

    Exit Sub
    
End Sub

Private Sub LabelContaContabilCR_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_LabelContaContabilCR_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabilCR.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 177718

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoContaContabilCR)

    Exit Sub

Erro_LabelContaContabilCR_Click:

    Select Case gErr

    Case 177718

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177710)

    End Select

End Sub

Private Sub LabelContaContabilRecDesp_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_LabelContaContabilRecDesp_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabilCR.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 177719

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaCRLista", colSelecao, objPlanoConta, objEventoContaContabilRecDesp)

    Exit Sub

Erro_LabelContaContabilRecDesp_Click:

    Select Case gErr

    Case 177719

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177720)

    End Select

End Sub

Private Sub objEventoContaContabilCR_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabilCR_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaContabilCR.Text = ""

    Else

        ContaContabilCR.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 177721

        ContaContabilCR.Text = sContaEnxuta

        ContaContabilCR.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabilCR_evSelecao:

    Select Case gErr

        Case 177721
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177722)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabilRecDesp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabilRecDesp_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaContabilRecDesp.Text = ""

    Else

        ContaContabilRecDesp.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 177723

        ContaContabilRecDesp.Text = sContaEnxuta

        ContaContabilRecDesp.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabilRecDesp_evSelecao:

    Select Case gErr

        Case 177723
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177724)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is ContaContabilCR Then
            Call LabelContaContabilCR_Click
        ElseIf Me.ActiveControl Is ContaContabilRecDesp Then
            Call LabelContaContabilRecDesp_Click
        End If
    End If
End Sub

