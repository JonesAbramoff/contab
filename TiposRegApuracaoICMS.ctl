VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TiposRegApuracaoICMSOcx 
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2010
   ScaleWidth      =   6690
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1245
      TabIndex        =   0
      Top             =   390
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   315
      Left            =   1800
      Picture         =   "TiposRegApuracaoICMS.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numera��o Autom�tica"
      Top             =   390
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4455
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TiposRegApuracaoICMS.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "TiposRegApuracaoICMS.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TiposRegApuracaoICMS.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TiposRegApuracaoICMS.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Secao 
      Height          =   315
      ItemData        =   "TiposRegApuracaoICMS.ctx":0A7E
      Left            =   1245
      List            =   "TiposRegApuracaoICMS.ctx":0A8B
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1395
      Width           =   3165
   End
   Begin VB.TextBox Descricao 
      Height          =   315
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   2
      Top             =   892
      Width           =   5370
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Se��o:"
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
      Left            =   525
      TabIndex        =   10
      Top             =   1455
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descri��o:"
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
      Left            =   210
      TabIndex        =   9
      Top             =   960
      Width           =   930
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "C�digo:"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   8
      Top             =   465
      Width           =   660
   End
End
Attribute VB_Name = "TiposRegApuracaoICMSOcx"
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
        
    'Se foi passado um Tipo de Registro de apura��o ICMS
    If Not objTipoApuracaoReg Is Nothing Then
    
        'Se o c�digo foi passado
        If objTipoApuracaoReg.iCodigo > 0 Then
        
            'L� o Tipo de Apura��o de Registro ICMS com o c�digo passado
            lErro = CF("TipoRegApuracaoICMS_Le",objTipoApuracaoReg)
            If lErro <> SUCESSO And lErro <> 67893 Then gError 67954
            
            'Se o Tipo de Apura��o est� cadastrado, exibe seus dados
            If lErro = SUCESSO Then
            
                'Traz os dados do Tipo de apura��o ICMS para a tela
                Call Traz_RegApuracaoICMS_Tela(objTipoApuracaoReg)
                        
            'Se n�o
            Else
                
                'Coloca o c�digo passado na tela
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
    
        Case 67954
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175009)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175010)

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

    'Informa tabela associada � Tela
    sTabela = "TiposRegApuracaoICMS"

    'Le os dados da tela
    Call Move_Tela_Memoria(objTiposReg)

    'Preenche a cole��o colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTiposReg.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objTiposReg.sDescricao, STRING_DESCRICAO_CAMPO, "Descricao"
    colCampoValor.Add "Secao", objTiposReg.iSecao, 0, "Secao"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175011)

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
    
    'Verifica se o C�digo est� preenchido
    If objTiposReg.iCodigo <> 0 Then

        'Traz os dados do tipo de apuran��o para a tela tela
        Call Traz_RegApuracaoICMS_Tela(objTiposReg)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175012)

    End Select

    Exit Sub

End Sub

Sub Traz_RegApuracaoICMS_Tela(objTipoReg As ClassTiposRegApuracao)
'Traz os dados da apura��o para a tela

Dim iIndice As Integer

    Codigo.PromptInclude = False
    Codigo.Text = objTipoReg.iCodigo
    Codigo.PromptInclude = True
    
    Descricao.Text = objTipoReg.sDescricao
    
    'Seleciona a se��o
    For iIndice = 0 To Secao.ListCount - 1
        If Secao.ItemData(iIndice) = objTipoReg.iSecao Then
            Secao.ListIndex = iIndice
            Exit For
        End If
    Next
                
    iAlterado = 0
    
End Sub

Sub Move_Tela_Memoria(objTipoReg As ClassTiposRegApuracao)
'Move dados da tela para a mem�ria

    objTipoReg.iCodigo = StrParaInt(Codigo.Text)
    objTipoReg.sDescricao = Descricao.Text
    If Secao.ListIndex <> -1 Then objTipoReg.iSecao = Secao.ItemData(Secao.ListIndex)
    
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera c�digo autom�tico para pr�ximo Tipo de Registro de apura��o ICMS
    lErro = CF("TipoRegApuracaoICMS_Codigo_Automatico",lCodigo)
    If lErro <> SUCESSO Then gError 67913
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 67913
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175013)
    
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
    
    'Se o C�digo foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
               
        'Critica o C�digo
        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 70092
    
    End If
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 70092
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175014)
        
    End Select
    
    Exit Sub
        
End Sub

Private Sub LabelCodigo_Click()

Dim colSelecao As New Collection
Dim objTipoReg As New ClassTiposRegApuracao

    Call Chama_Tela("TiposRegApuracaoICMSLista", colSelecao, objTipoReg, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoReg As ClassTiposRegApuracao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTipoReg = obj1

    Call Traz_RegApuracaoICMS_Tela(objTipoReg)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175015)

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

    'Grava um Tipo de apura��o
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 67883

    'Limpa a tela
    Call Limpa_Tela_TipoRegApuracao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 67883

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175016)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoReg As New ClassTiposRegApuracao

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o C�digo esta preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 67884

    'Verifica se a descri��o est� preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 67885

    'Verifica se a Se��o foi preenchida
    If Len(Trim(Secao.Text)) = 0 Then gError 67886
    
    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objTipoReg)

    'Necess�rio para o funcionamento da fun��o Trata_Alteracao neste caso
    objTipoReg.sNomeTabela = "ICMS"

    lErro = Trata_Alteracao(objTipoReg, objTipoReg.iCodigo)
    If lErro <> SUCESSO Then gError 32334

    'Grava um tipo de apura��o
    lErro = CF("TipoRegApuracaoICMS_Grava",objTipoReg)
    If lErro <> SUCESSO Then gError 67887
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 32334

        Case 67884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 67885
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 67886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SECAO_NAO_PREENCHIDA", gErr)

        Case 67887

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175017)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Exclui Tipo de apura�ao ICMS

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoReg As New ClassTiposRegApuracao

On Error GoTo Erro_BotaoExcluir_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o c�digo n�o foi preenchido, erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 67888
    
    'Guarda o c�digo
    objTipoReg.iCodigo = CInt(Codigo.Text)
    
    'L� o tipo de registro para apura��o ICMS
    lErro = CF("TipoRegApuracaoICMS_Le",objTipoReg)
    If lErro <> SUCESSO And lErro <> 67893 Then gError 67894
    
    'Se n�o encontrou, erro
    If lErro = 67893 Then gError 67895
    
    'Pede a confirma��o da exclus�o do tipo de registro de apura��o ICMS
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TIPOREGAPURACAOICMS", objTipoReg.iCodigo)
    If vbMsgRes = vbNo Then gError 69191
    
    'Exclui o tipo de apura��o
    lErro = CF("TipoRegApuracaoICMS_Exclui",objTipoReg)
    If lErro <> SUCESSO Then gError 67889

    'Limpa a tela
    Call Limpa_Tela_TipoRegApuracao

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 67888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 67889, 67894, 69191
        
        Case 67895
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOICMS_NAO_CADASTRADA", gErr, objTipoReg.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175018)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se h� altera��es e quer salv�-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 67895

    'Limpa a tela
    Call Limpa_Tela_TipoRegApuracao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 67895

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175019)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If
    
    End If
    
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

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tipos de Registro para apura��o de ICMS"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TiposRegApuracaoICMS"

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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

