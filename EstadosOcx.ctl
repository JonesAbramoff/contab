VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl EstadosOcx 
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LockControls    =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   4935
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1620
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "EstadosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "EstadosOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "EstadosOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alíquotas para ICMS"
      Height          =   825
      Left            =   120
      TabIndex        =   7
      Top             =   1515
      Width           =   4620
      Begin MSMask.MaskEdBox Interna 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   345
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "##0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Importacao 
         Height          =   285
         Left            =   3345
         TabIndex        =   2
         Top             =   360
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "##0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Interna:"
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
         Left            =   105
         TabIndex        =   9
         Top             =   390
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "Importação:"
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
         Left            =   2235
         TabIndex        =   8
         Top             =   390
         Width           =   1080
      End
   End
   Begin VB.ComboBox Sigla 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   300
      Width           =   888
   End
   Begin VB.Label Label1 
      Caption         =   "Nome:"
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
      Left            =   150
      TabIndex        =   12
      Top             =   990
      Width           =   600
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   225
      TabIndex        =   11
      Top             =   330
      Width           =   540
   End
   Begin VB.Label Nome 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   810
      TabIndex        =   10
      Top             =   930
      Width           =   3930
   End
End
Attribute VB_Name = "EstadosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iSiglaAlterada As Integer

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then Error 41524

    Call Limpa_Tela_Estados
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
          
        Case 41524
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159482)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se algum campo foi alterado
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 28493

    'Limpa a tela Estados
    Call Limpa_Tela_Estados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 28493

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159483)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_Estados() As Long
'Limpa todos os campos de input da tela Estados

Dim lErro As Long

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Sigla.ListIndex = -1
    
    Sigla.Text = ""
    
    Nome.Caption = ""
    
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colEstado As New Collection
Dim objEstado As ClassEstado

On Error GoTo Erro_Form_Load

    'Preenche a ComboBox com as siglas dos Estados existentes no BD
    lErro = CF("Estados_Le_Todos",colEstado)
    If lErro <> SUCESSO Then Error 28253

    For Each objEstado In colEstado

       'Insere na ComboBox a sigla do Estado
       Sigla.AddItem objEstado.sSigla

    Next

    'Preenche a ComboBox Sigla
    Sigla.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 28253

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159484)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Importacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Importacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentual As Double

On Error GoTo Erro_Importacao_Validate

    'Verifica se ICMSAliquotaImportacao está preenchida
    If Len(Trim(Importacao.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Importacao.Text)
        If lErro <> SUCESSO Then Error 28496
            
        lErro = Porcentagem_Critica(Importacao.Text)
        If lErro <> SUCESSO Then Error 40772
    
    End If

    Exit Sub

Erro_Importacao_Validate:

    Cancel = True


    Select Case Err

        Case 28496, 40772

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159485)

    End Select

    Exit Sub

End Sub
Private Sub Interna_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Interna_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentual As Double

On Error GoTo Erro_Interna_Validate

    'Verifica se ICMSAliquotaInterna está preenchida
    If Len(Trim(Interna.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Interna.Text)
        If lErro <> SUCESSO Then Error 28498
        
        lErro = Porcentagem_Critica(Interna.Text)
        If lErro <> SUCESSO Then Error 40773

    End If

    Exit Sub

Erro_Interna_Validate:

    Cancel = True


    Select Case Err

        Case 28498, 40773

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159486)

    End Select

    Exit Sub

End Sub

Private Sub Sigla_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sigla_Click()

Dim lErro As Long
Dim objEstado As New ClassEstado

On Error GoTo Erro_Sigla_Click

    If Sigla.ListIndex = -1 Then Exit Sub
    
    objEstado.sSigla = Sigla.Text

    'Lê o Estado
    lErro = CF("Estado_Le",objEstado)
    If lErro <> SUCESSO And lErro <> 28485 Then Error 28494

    If lErro = 28485 Then Error 28495

    'Traz os dados do Estado para tela
    lErro = Traz_Estado_Tela(objEstado)
    If lErro <> SUCESSO Then Error 19474
    
   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub
    
Erro_Sigla_Click:

    Select Case Err

        Case 28494, 19474

        Case 28495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ESTADO_NAO_CADASTRADA", Err, Sigla.Text)
            Sigla.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159487)

    End Select

    Exit Sub
    
End Sub

Private Sub Sigla_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Error_Sigla_Validate

    'Verifica se foi preenchida a ComboBox Sigla
    If Len(Trim(Sigla.Text)) = 0 Then Exit Sub

    If Sigla.ListIndex <> -1 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Item_Igual(Sigla)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 28490

    'Não existe o item na ComboBox Sigla
    If lErro = 12253 Then Error 28491

    Exit Sub

Error_Sigla_Validate:

    Cancel = True


    Select Case Err

        Case 28490

        Case 28491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ESTADO_NAO_CADASTRADA", Err, Sigla.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159488)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim sSigla As String
Dim objEstado As New ClassEstado

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se dados Sigla de Estado foi preenchida
    If Len(Trim(Sigla.Text)) = 0 Then Error 28500

    'Preenche objEstado
    objEstado.sSigla = Sigla.Text
    If Len(Trim(Interna.Text)) <> 0 Then objEstado.dICMSAliquotaInterna = CDbl(Interna.Text) / 100
    objEstado.dICMSAliquotaExportacao = 0 'Pode ser usado futuramente
    If Len(Trim(Importacao.Text)) <> 0 Then objEstado.dICMSAliquotaImportacao = CDbl(Importacao.Text) / 100

    lErro = Trata_Alteracao(objEstado, objEstado.sSigla)
    If lErro <> SUCESSO Then Error 32316

    'Grava os campos de Estado no Banco de Dados
    lErro = CF("Estado_Grava",objEstado)
    If lErro <> SUCESSO Then Error 28501

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 32316

        Case 28500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ESTADO_NAO_PREENCHIDA", Err)
            Sigla.SetFocus

        Case 28501

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159489)

     End Select

     Exit Function

End Function

Function Trata_Parametros(Optional objEstado As ClassEstado) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Estado selecionado, exibir seus dados
    If Not (objEstado Is Nothing) Then

        Sigla.Text = objEstado.sSigla

        lErro = Combo_Item_Igual(Sigla)
        'Se não encontrou o Estado em questão
        If lErro <> SUCESSO And lErro <> 12253 Then Error 28511

    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 28510

        Case 28511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, objEstado.sSigla)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159490)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim objEstado As New ClassEstado

    'Informa tabela associada à Tela
    sTabela = "Estados"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Sigla", Sigla.Text, STRING_ESTADO_SIGLA, "Sigla"
    colCampoValor.Add "Nome", Nome.Caption, STRING_ESTADO_NOME, "Nome"
    colCampoValor.Add "ICMSAliquotaInterna", objEstado.dICMSAliquotaInterna, 0, "ICMSAliquotaInterna"
'    colCampoValor.Add "ICMSAliquotaExportacao", objEstado.dICMSAliquotaExportacao, 0, "ICMSAliquotaExportacao"
    colCampoValor.Add "ICMSAliquotaImportacao", objEstado.dICMSAliquotaImportacao, 0, "ICMSAliquotaImportacao"

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objEstado As New ClassEstado

On Error GoTo Erro_Tela_Preenche

    objEstado.sSigla = colCampoValor.Item("Sigla").vValor

    If objEstado.sSigla <> "" Then

        'Carrega objEstado com os dados passados em colCampoValor
        objEstado.sNome = colCampoValor.Item("Nome").vValor
        objEstado.dICMSAliquotaInterna = colCampoValor.Item("ICMSAliquotaInterna").vValor
'        objEstado.dICMSAliquotaExportacao = colCampoValor.Item("ICMSAliquotaExportacao").vValor
        objEstado.dICMSAliquotaImportacao = colCampoValor.Item("ICMSAliquotaImportacao").vValor

        lErro = Traz_Estado_Tela(objEstado)
        If lErro <> SUCESSO Then Error 19475
        
    End If
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case Err
    
        Case 19475

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159491)

    End Select

    Exit Sub

End Sub

Private Function Traz_Estado_Tela(objEstado As ClassEstado) As Long
'Traz para a tela os dados do Estado passado
    
    'Coloca na tela os dados do Estado
    
    Sigla.Text = objEstado.sSigla
    Nome.Caption = objEstado.sNome
    If objEstado.dICMSAliquotaInterna > 0 Then
        Interna.Text = objEstado.dICMSAliquotaInterna * 100
    Else
        Interna.Text = ""
    End If
    
'    If objEstado.dICMSAliquotaExportacao > 0 Then
'        Exportacao.Text = objEstado.dICMSAliquotaExportacao
'    Else
'        Exportacao.Text = ""
'    End If
'
    If objEstado.dICMSAliquotaImportacao > 0 Then
        Importacao.Text = objEstado.dICMSAliquotaImportacao * 100
    Else
        Importacao.Text = ""
    End If
        
    iAlterado = 0
    
    Traz_Estado_Tela = SUCESSO
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ESTADOS
    Set Form_Load_Ocx = Me
    Caption = "Estados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Estados"
    
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
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Nome_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nome, Source, X, Y)
End Sub

Private Sub Nome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nome, Button, Shift, X, Y)
End Sub

