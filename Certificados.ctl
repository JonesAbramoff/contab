VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Certificados 
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   6690
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1980
      Picture         =   "Certificados.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Numeração Automática"
      Top             =   195
      Width           =   300
   End
   Begin VB.ListBox Certificados 
      Height          =   1815
      ItemData        =   "Certificados.ctx":00EA
      Left            =   4455
      List            =   "Certificados.ctx":00EC
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1125
      Width           =   2085
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   4500
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "Certificados.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "Certificados.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Certificados.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "Certificados.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.TextBox Descricao 
      Height          =   1350
      Left            =   1065
      MaxLength       =   250
      TabIndex        =   2
      Top             =   1125
      Width           =   3330
   End
   Begin MSMask.MaskEdBox Sigla 
      Height          =   315
      Left            =   1065
      TabIndex        =   1
      Top             =   660
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Validade 
      Height          =   315
      Left            =   1065
      TabIndex        =   3
      Top             =   2610
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1095
      TabIndex        =   0
      Top             =   195
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
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
      Left            =   -555
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   225
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "meses:"
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
      Left            =   1365
      TabIndex        =   14
      Top             =   2655
      Width           =   945
   End
   Begin VB.Label Label13 
      Caption         =   "Certificados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4440
      TabIndex        =   13
      Top             =   855
      Width           =   2055
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
      Left            =   -555
      TabIndex        =   11
      Top             =   1170
      Width           =   1500
   End
   Begin VB.Label LabelSigla 
      Alignment       =   1  'Right Justify
      Caption         =   "Sigla:"
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
      Left            =   -555
      TabIndex        =   12
      Top             =   690
      Width           =   1500
   End
   Begin VB.Label LabelValidade 
      Alignment       =   1  'Right Justify
      Caption         =   "Validade:"
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
      Left            =   -555
      TabIndex        =   10
      Top             =   2640
      Width           =   1500
   End
End
Attribute VB_Name = "Certificados"
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
    Caption = "Certificados"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Certificados"

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213231)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As AdmColCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    
    Set colCodigoDescricao = New AdmColCodigoNome
    
    'Lê o Código e a Descrição de cada Tipo de Mão-de-Obra
    lErro = CF("Cod_Nomes_Le", "Certificados", "Codigo", "Sigla", STRING_MAXIMO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'preenche a ListBox certificados com os objetos da colecao
    For Each objCodigoDescricao In colCodigoDescricao
        Certificados.AddItem objCodigoDescricao.sNome
        Certificados.ItemData(Certificados.NewIndex) = objCodigoDescricao.iCodigo
    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213232)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal objCertificados As ClassCertificados) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCertificados Is Nothing) Then

        lErro = Traz_Certificados_Tela(objCertificados)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213233)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCertificados As ClassCertificados) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objCertificados.lCodigo = StrParaLong(Codigo.Text)
    objCertificados.sDescricao = Descricao.Text
    objCertificados.sSigla = Sigla.Text
    objCertificados.lValidade = StrParaLong(Validade.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213234)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCertificados As New ClassCertificados

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Certificados"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objCertificados)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCertificados.lCodigo, 0, "Codigo"
    colCampoValor.Add "Sigla", objCertificados.sSigla, STRING_MAXIMO, "Sigla"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213235)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCertificados As New ClassCertificados

On Error GoTo Erro_Tela_Preenche

    objCertificados.lCodigo = colCampoValor.Item("Codigo").vValor

    If objCertificados.lCodigo <> 0 Then

        lErro = Traz_Certificados_Tela(objCertificados)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213236)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCertificados As New ClassCertificados

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 213237
    If Len(Trim(Sigla.Text)) = 0 Then gError 213238
    '#####################

    'Preenche o objCertificados
    lErro = Move_Tela_Memoria(objCertificados)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Trata_Alteracao(objCertificados, objCertificados.lCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Grava o/a Certificados no Banco de Dados
    lErro = CF("Certificados_Grava", objCertificados)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Remove o item da lista de Certificados
    Call Certificados_Exclui(objCertificados.lCodigo)

    'Insere o item na lista de Certificados
    Call Certificados_Adiciona(objCertificados)

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 213237
            Call Rotina_Erro(vbOKOnly, "ERRO_CERTIFICADO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 213238
            Call Rotina_Erro(vbOKOnly, "ERRO_CERTIFICADO_SIGLA_NAO_PREENCHIDA", gErr)
            Sigla.SetFocus
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213239)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Certificados() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Certificados

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_Certificados = SUCESSO

    Exit Function

Erro_Limpa_Tela_Certificados:

    Limpa_Tela_Certificados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213240)

    End Select

    Exit Function

End Function

Function Traz_Certificados_Tela(objCertificados As ClassCertificados) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Certificados_Tela

    Call Limpa_Tela_Certificados

    Sigla.Text = objCertificados.sSigla

    'Lê o Certificados que está sendo Passado
    lErro = CF("Certificados_Le", objCertificados)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then

        If objCertificados.lCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objCertificados.lCodigo)
            Codigo.PromptInclude = True
        End If

        Descricao.Text = objCertificados.sDescricao
        Sigla.Text = objCertificados.sSigla

        If objCertificados.lValidade <> 0 Then
            Validade.PromptInclude = False
            Validade.Text = CStr(objCertificados.lValidade)
            Validade.PromptInclude = True
        End If

    End If

    iAlterado = 0

    Traz_Certificados_Tela = SUCESSO

    Exit Function

Erro_Traz_Certificados_Tela:

    Traz_Certificados_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213241)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_Certificados

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213242)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213243)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_Certificados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213244)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCertificados As New ClassCertificados
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 213245

    objCertificados.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CERTIFICADOS", objCertificados.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Certificados_Exclui", objCertificados)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Remove o item da lista de Certificados
        Call Certificados_Exclui(objCertificados.lCodigo)

        'Limpa Tela
        Call Limpa_Tela_Certificados

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 213245
            Call Rotina_Erro(vbOKOnly, "ERRO_CERTIFICADO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213246)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213247)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213248)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213249)

    End Select

    Exit Sub

End Sub

Private Sub Sigla_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Validade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Validade_Validate

    'Verifica se Validade está preenchida
    If Len(Trim(Validade.Text)) <> 0 Then

       'Critica a Validade
       lErro = Long_Critica(Validade.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Validade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213250)

    End Select

    Exit Sub

End Sub

Private Sub Validade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Validade, iAlterado)
    
End Sub

Private Sub Validade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCertificados As ClassCertificados

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCertificados = obj1

    'Mostra os dados do Certificados na tela
    lErro = Traz_Certificados_Tela(objCertificados)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213251)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objCertificados As New ClassCertificados
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objCertificados.lCodigo = Codigo.Text

    End If
    objCertificados.sSigla = Trim(Sigla.Text)

    Call Chama_Tela("CertificadosLista", colSelecao, objCertificados, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213252)

    End Select

    Exit Sub

End Sub

Private Sub Certificados_DblClick()

Dim lErro As Long
Dim objCertificados As New ClassCertificados

On Error GoTo Erro_Certificados_DblClick

    'Guarda o valor do codigo do Tipo da Mão-de-Obra selecionado na ListBox Tipos
    objCertificados.lCodigo = Certificados.ItemData(Certificados.ListIndex)

    'Mostra os dados do TiposDeMaodeObra na tela
    lErro = Traz_Certificados_Tela(objCertificados)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Certificados_DblClick:

    Certificados.SetFocus

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213253)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "ESTConfig", "NUM_PROX_CERTIFICADOS", "Certificados", "Codigo", lCodigo)
    If lErro <> SUCESSO And lErro <> 25191 Then gError ERRO_SEM_MENSAGEM
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213254)
    
    End Select

    Exit Sub

End Sub

Private Sub Certificados_Adiciona(objCertificados As ClassCertificados)

    Certificados.AddItem objCertificados.sSigla
    Certificados.ItemData(Certificados.NewIndex) = objCertificados.lCodigo

End Sub

Private Sub Certificados_Exclui(lCodigo As Long)

Dim iIndice As Integer

    For iIndice = 0 To Certificados.ListCount - 1

        If Certificados.ItemData(iIndice) = lCodigo Then

            Certificados.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
    
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
        
    End If
    
End Sub
