VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TiposRegApuracaoIPIOcx 
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   ScaleHeight     =   2025
   ScaleWidth      =   6660
   Begin VB.TextBox Descricao 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   900
      Width           =   5340
   End
   Begin VB.ComboBox Secao 
      Height          =   315
      ItemData        =   "TiposRegApuracaoIPI.ctx":0000
      Left            =   1200
      List            =   "TiposRegApuracaoIPI.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1410
      Width           =   3165
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4395
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   180
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TiposRegApuracaoIPI.ctx":0035
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TiposRegApuracaoIPI.ctx":01B3
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "TiposRegApuracaoIPI.ctx":06E5
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TiposRegApuracaoIPI.ctx":086F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   315
      Left            =   1770
      Picture         =   "TiposRegApuracaoIPI.ctx":09C9
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   405
      Width           =   300
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   405
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
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
      Left            =   450
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   480
      Width           =   660
   End
   Begin VB.Label Label2 
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
      Left            =   180
      TabIndex        =   10
      Top             =   975
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Seção:"
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
      Left            =   495
      TabIndex        =   9
      Top             =   1470
      Width           =   615
   End
End
Attribute VB_Name = "TiposRegApuracaoIPIOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

'Eventos dos Browses
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Function Trata_Parametros(Optional objTipoApuracaoReg As ClassTiposRegApuracao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se foi passado um Tipo de Registro de apuração IPI
    If Not objTipoApuracaoReg Is Nothing Then

        'Se o código foi passado
        If objTipoApuracaoReg.iCodigo > 0 Then

            'Lê o Tipo de Apuração de Registro IPI com o código passado
            lErro = CF("TipoRegApuracaoIPI_Le",objTipoApuracaoReg)
            If lErro <> SUCESSO And lErro <> 79024 Then gError 79047

            'Se o Tipo de Apuração está cadastrado, exibe seus dados
            If lErro = SUCESSO Then

                'Traz os dados do Tipo de apuração IPI para a tela
                Call Traz_RegApuracaoIPI_Tela(objTipoApuracaoReg)

            'Se não
            Else

                'Coloca o código passado na tela
                Codigo.PromptInclude = False
                Codigo.Text = objTipoApuracaoReg.iCodigo
                Codigo.PromptInclude = True

            End If

        End If

    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 79047

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175029)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175030)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCodigo = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTiposReg As New ClassTiposRegApuracao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposRegApuracaoIPI"

    'Le os dados da tela
    Call Move_Tela_Memoria(objTiposReg)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTiposReg.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objTiposReg.sDescricao, STRING_DESCRICAO_CAMPO, "Descricao"
    colCampoValor.Add "Secao", objTiposReg.iSecao, 0, "Secao"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175031)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objTiposReg As New ClassTiposRegApuracao
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objTiposReg com os dados passados em colCampoValor
    objTiposReg.iCodigo = colCampoValor.Item("Codigo").vValor
    objTiposReg.sDescricao = colCampoValor.Item("Descricao").vValor
    objTiposReg.iSecao = colCampoValor.Item("Secao").vValor

    'Verifica se o Código está preenchido
    If objTiposReg.iCodigo <> 0 Then

        'Traz os dados do tipo de apuranção para a tela tela
        Call Traz_RegApuracaoIPI_Tela(objTiposReg)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175032)

    End Select

    Exit Sub

End Sub

Sub Traz_RegApuracaoIPI_Tela(objTipoReg As ClassTiposRegApuracao)
'Traz os dados da apuração para a tela

Dim iIndice As Integer

    Codigo.PromptInclude = False
    Codigo.Text = objTipoReg.iCodigo
    Codigo.PromptInclude = True
    Descricao.Text = objTipoReg.sDescricao

    'Seleciona a seção
    For iIndice = 0 To Secao.ListCount - 1
        If Secao.ItemData(iIndice) = objTipoReg.iSecao Then
            Secao.ListIndex = iIndice
            Exit For
        End If
    Next

    iAlterado = 0

End Sub

Sub Move_Tela_Memoria(objTipoReg As ClassTiposRegApuracao)
'Move dados da tela para a memória

    objTipoReg.iCodigo = StrParaInt(Codigo.Text)
    objTipoReg.sDescricao = Descricao.Text
    If Secao.ListIndex <> -1 Then objTipoReg.iSecao = Secao.ItemData(Secao.ListIndex)

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático para próximo Tipo de Registro de apuração IPI
    lErro = CF("TipoRegApuracaoIPI_Codigo_Automatico",lCodigo)
    If lErro <> SUCESSO Then gError 79048
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 79048

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175033)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        'Critica o Código
        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 79049

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 79049

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175034)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim colSelecao As New Collection
Dim objTipoReg As ClassTiposRegApuracao

    Call Chama_Tela("TiposRegApuracaoIPILista", colSelecao, objTipoReg, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoReg As ClassTiposRegApuracao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTipoReg = obj1

    Call Traz_RegApuracaoIPI_Tela(objTipoReg)

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175035)

    End Select

    Exit Sub

End Sub

Private Sub Secao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava um Tipo de apuração
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 79050

    'Limpa a tela
    Call Limpa_Tela_TipoRegApuracao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 79050

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175036)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoReg As New ClassTiposRegApuracao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Código esta preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 79051

    'Verifica se a descrição está preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 79052

    'Verifica se a Seção foi preenchida
    If Len(Trim(Secao.Text)) = 0 Then gError 79053

    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objTipoReg)

    'Necessário para o funcionamento da função Trata_Alteracao neste caso
    objTipoReg.sNomeTabela = "IPI"

    lErro = Trata_Alteracao(objTipoReg, objTipoReg.iCodigo)
    If lErro <> SUCESSO Then gError 32335

    'Grava um tipo de apuração
    lErro = CF("TipoRegApuracaoIPI_Grava",objTipoReg)
    If lErro <> SUCESSO Then gError 79054

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 32335

        Case 79051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 79052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 79053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SECAO_NAO_PREENCHIDA", gErr)

        Case 79054

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175037)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Exclui Tipo de apuraçao IPI

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoReg As New ClassTiposRegApuracao

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Se o código não foi preenchido, erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 79055

    'Guarda o código
    objTipoReg.iCodigo = CInt(Codigo.Text)

    'Lê o tipo de registro para apuração IPI
    lErro = CF("TipoRegApuracaoIPI_Le",objTipoReg)
    If lErro <> SUCESSO And lErro <> 79024 Then gError 79056

    'Se não encontrou, erro
    If lErro = 79024 Then gError 79057

    'Pede a confirmação da exclusão do tipo de registro de apuração IPI
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TIPOREGAPURACAOIPI", objTipoReg.iCodigo)
    If vbMsgRes = vbNo Then gError 79058

    'Exclui o tipo de apuração
    lErro = CF("TipoRegApuracaoIPI_Exclui",objTipoReg)
    If lErro <> SUCESSO Then gError 79059

    'Limpa a tela
    Call Limpa_Tela_TipoRegApuracao

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 79055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 79056, 79058, 79059

        Case 79057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOIPI_NAO_CADASTRADA", gErr, objTipoReg.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175038)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 79057

    'Limpa a tela
    Call Limpa_Tela_TipoRegApuracao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 79057

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175039)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_TipoRegApuracao()

Dim lErro As Long

    'Limpa o restante da tela
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
    Descricao.Text = ""
    Secao.ListIndex = -1

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
    
        Case KEYCODE_BROWSER
            If Me.ActiveControl Is Codigo Then
                Call LabelCodigo_Click
            End If
            
    End Select
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tipos de Registro para apuração de IPI"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TiposRegApuracaoIPI"

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

'**** fim do trecho a ser copiado *****

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

