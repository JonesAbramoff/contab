VERSION 5.00
Begin VB.UserControl TranspConsultaPhar 
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ScaleHeight     =   2040
   ScaleWidth      =   5550
   Begin VB.PictureBox Picture9 
      Height          =   540
      Left            =   4245
      ScaleHeight     =   480
      ScaleWidth      =   1095
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   1155
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   585
         Picture         =   "TranspConsultaPhar.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   75
         Picture         =   "TranspConsultaPhar.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoConsulta 
      Height          =   675
      Left            =   2295
      Picture         =   "TranspConsultaPhar.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1185
      Width           =   1275
   End
   Begin VB.ComboBox Transportadora 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   660
      Width           =   3750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transportadora:"
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
      Left            =   285
      TabIndex        =   1
      Top             =   720
      Width           =   1365
   End
End
Attribute VB_Name = "TranspConsultaPhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0

    'Carrega a listbox de atendentes
    lErro = Carrega_Transportadoras()
    If lErro <> SUCESSO Then gError 102718

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 102718

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143145)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub BotaoConsulta_Click()

Dim colSelecao As New Collection

    If Len(Trim(Transportadora.Text)) > 0 Then
    
        colSelecao.Add LCodigo_Extrai(Transportadora.Text)

        Call Chama_Tela("NFsDirectLista", colSelecao, Nothing, Nothing, "CodTransportadora = ?")
        
    End If

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    Call Limpa_Tela(Me)

    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Transportadora.ListIndex = -1

    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160759)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long

Dim lErro As Long
Dim objUsuario As ClassUsuarios

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143146)

    End Select

    iAlterado = 0

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Call ComandoSeta_Liberar(Me.Name)
End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'**** TRATAMENTO DO SISTEMA DE SETAS - INÍCIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143155)

    End Select

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objAtendente As New ClassAtendentes
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143156)

    End Select

End Sub

Public Sub Form_Activate()
    'Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    'gi_ST_SetaIgnoraClick = 1
End Sub
'**** TRATAMENTO DO SISTEMA DE SETAS - FIM ****



'***************************************************
'Início do trecho de codigo comum as telas
'***************************************************
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Consulta Transportadoras"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "Consulta Transportadoras"
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
'***************************************************
'Fim Trecho de codigo comum as telas
'***************************************************

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - INÍCIO ***
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***

Private Function Carrega_Transportadoras() As Long
'Carrega as Combos Transportadora e TransportadoraFinal

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadoras

    'Lê Códigos e NomesReduzidos da tabela Transportadora e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 37471
    
    'preenche as combos iniciais e finais
    For Each objCodigoNome In colCodigoDescricao
        
        If objCodigoNome.iCodigo <> 0 Then
            Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
        End If
    
    Next

    Carrega_Transportadoras = SUCESSO

    Exit Function

Erro_Carrega_Transportadoras:

    Carrega_Transportadoras = gErr

    Select Case gErr

        'Erro já tratado
        Case 37471

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173576)

    End Select

    Exit Function

End Function
