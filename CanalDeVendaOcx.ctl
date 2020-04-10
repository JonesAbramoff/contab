VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CanalDeVendaOcx 
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   5685
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2190
      Picture         =   "CanalDeVendaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   345
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3405
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CanalDeVendaOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CanalDeVendaOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CanalDeVendaOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CanalDeVendaOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox CanaisList 
      Height          =   1425
      ItemData        =   "CanalDeVendaOcx.ctx":0A7E
      Left            =   150
      List            =   "CanalDeVendaOcx.ctx":0A80
      TabIndex        =   4
      Top             =   2145
      Width           =   5400
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1665
      TabIndex        =   0
      Top             =   330
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   315
      Left            =   1665
      TabIndex        =   2
      Top             =   885
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1665
      TabIndex        =   3
      Top             =   1440
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      Caption         =   "Canais de Venda"
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
      Left            =   165
      TabIndex        =   13
      Top             =   1935
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Nome Reduzido:"
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
      Left            =   150
      TabIndex        =   12
      Top             =   1500
      Width           =   1425
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1020
      TabIndex        =   11
      Top             =   930
      Width           =   570
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   990
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "CanalDeVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("CanalVenda_Automatico",iCodigo)
    If lErro <> SUCESSO Then Error 57528

    Codigo.PromptInclude = False
    Codigo.Text = CStr(iCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57528
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144140)
    
    End Select

    Exit Sub

End Sub

Sub Traz_Canal_Tela(objCanal As ClassCanalVenda)

    'mostra dados do Canal de Venda na tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(objCanal.iCodigo)
    Codigo.PromptInclude = True
    NomeReduzido.Text = objCanal.sNomeReduzido
    Nome.Text = objCanal.sNome
    
    iAlterado = 0

End Sub

Private Sub CanaisList_DblClick()

Dim lErro As Long
Dim objCanal As New ClassCanalVenda

On Error GoTo Erro_CanaisList_DblClick
    
    objCanal.iCodigo = CanaisList.ItemData(CanaisList.ListIndex)
    
    'Le o Canal de Venda selecionado
    lErro = CF("CanalVenda_Le",objCanal)
    If lErro <> SUCESSO Then Error 23593
    
    'Traz os dados para a Tela
    Call Traz_Canal_Tela(objCanal)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_CanaisList_DblClick:

    Select Case Err

        Case 23593

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144141)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim objCanal As New ClassCanalVenda

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 23598

    objCanal.iCodigo = CInt(Codigo.Text)
    
    'Verifica se está no BD
    lErro = CF("CanalVenda_Le",objCanal)
    If lErro <> SUCESSO And lErro <> 23597 Then Error 23599
    
    If lErro = 23597 Then Error 49755
            
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CANAL", objCanal.iCodigo)

    If vbMsgRes = vbYes Then

        'exclui canal
        lErro = CF("CanalVenda_Exclui",objCanal)
        If lErro <> SUCESSO Then Error 23600
        
        'Exclui da ListBox
        Call CanaisList_Remove(objCanal)
       
        'Limpa tela e gera automaticamente novo Código para o Canal
        lErro = Limpa_Tela_Canal
        If lErro <> SUCESSO Then Error 23621

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 23599, 23600, 23609, 23621

        Case 49755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_CADASTRADO", Err, objCanal.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144142)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se codigo é numérico
        If Not IsNumeric(Codigo.Text) Then Error 23641

        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 23642

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err
        
        Case 23641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, Codigo.Text)

        Case 23642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", Err, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144143)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub


Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 23622

    lErro = Limpa_Tela_Canal
    If lErro <> SUCESSO Then Error 23623

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23622, 23623

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144144)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 23623

    lErro = Limpa_Tela_Canal
    If lErro <> SUCESSO Then Error 23624
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 23623, 23624

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144145)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim colCodigoDescricao As New AdmColCodigoNome
Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Form_Load

    'Preenche a listbox Canais de venda
    'Le cada codigo e Nome Reduzido da tabela CanalVenda
    lErro = CF("Cod_Nomes_Le","CanalVenda", "Codigo", "NomeReduzido", STRING_CANAL_VENDA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 23592

    'preenche a listbox Canais com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        CanaisList.AddItem objCodigoDescricao.iCodigo & SEPARADOR & objCodigoDescricao.sNome
        CanaisList.ItemData(CanaisList.NewIndex) = objCodigoDescricao.iCodigo

    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144146)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objCanal As ClassCanalVenda) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de CanalVenda

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objCanal Is Nothing) Then
        
        'Le o canal de venda passado
        lErro = CF("CanalVenda_Le",objCanal)
        If lErro <> SUCESSO And lErro <> 23597 Then Error 23643

        If lErro = SUCESSO Then
            
            'Preenche a Tela com o canal de venda
            Call Traz_Canal_Tela(objCanal)
        
        Else
            'Senão preenche só o código
            Codigo.PromptInclude = False
            Codigo.Text = objCanal.iCodigo
            Codigo.PromptInclude = True
                        
        End If
        
    Else

        'Limpa a Tela e gera proximo Codigo para canal
        lErro = Limpa_Tela_Canal
        If lErro <> SUCESSO Then Error 23644

    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 23643, 23644

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144147)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCanal As New ClassCanalVenda
On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 23625

    'verifica preenchimento do nome
    If Len(Trim(Nome.Text)) = 0 Then Error 23626

    'verifica preenchimento do nome reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 23627

    'preenche objCanal
    objCanal.iCodigo = CInt(Codigo.Text)
    objCanal.sNome = Nome.Text
    objCanal.sNomeReduzido = NomeReduzido.Text
    
    lErro = Trata_Alteracao(objCanal, objCanal.iCodigo)
    If lErro <> SUCESSO Then Error 32332
    
    'Grava o canal de venda
    lErro = CF("CanalVenda_Grava",objCanal)
    If lErro <> SUCESSO Then Error 23628

    'Atualiza ListBox de Canais
    Call CanaisList_Remove(objCanal)
    Call CanaisList_Adiciona(objCanal)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 23626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", Err)

        Case 23627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", Err)

        Case 23628, 32332

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144148)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub


Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 57823

    End If
        
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    
    Select Case Err
    
        Case 57823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144149)
    
    End Select
    
    Exit Sub
    
End Sub

Function Limpa_Tela_Canal() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Canal

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)

    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True

    'Desselecionar Lisbox
    CanaisList.ListIndex = -1

    'Zerar iAlterado
    iAlterado = 0

    Limpa_Tela_Canal = SUCESSO

    Exit Function

Erro_Limpa_Tela_Canal:

    Limpa_Tela_Canal = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144150)

    End Select

    Exit Function

End Function

Private Sub CanaisList_Adiciona(objCanal As ClassCanalVenda)
'Inclui Canal na List

Dim iIndice As Integer

    For iIndice = 0 To CanaisList.ListCount - 1

        If CanaisList.ItemData(iIndice) > objCanal.iCodigo Then Exit For
        
    Next

    CanaisList.AddItem objCanal.iCodigo & SEPARADOR & objCanal.sNomeReduzido, iIndice
    CanaisList.ItemData(iIndice) = objCanal.iCodigo

End Sub

Private Sub CanaisList_Remove(objCanal As ClassCanalVenda)
'Percorre a ListBox Canaislist para remover ocanal caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To CanaisList.ListCount - 1
    
        If CanaisList.ItemData(iIndice) = objCanal.iCodigo Then
    
            CanaisList.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim iIndice As Integer

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    Codigo.PromptInclude = False
    Codigo.Text = CStr(colCampoValor.Item("Codigo").vValor)
    Codigo.PromptInclude = True
    Nome.Text = colCampoValor.Item("Nome").vValor
    NomeReduzido.Text = colCampoValor.Item("NomeReduzido").vValor
    
    'Seleciona Nome Reduzido na ListBox
    For iIndice = 0 To CanaisList.ListCount - 1

        If CanaisList.ItemData(iIndice) = CInt(Codigo.Text) Then
            CanaisList.ListIndex = iIndice
            Exit For
        End If

    Next
    
    iAlterado = 0

End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

'Dim Geral
Dim objCampoValor As AdmCampoValor
'Dim específicos
Dim iCodigo As Integer

    'Informa tabela associada à Tela
    sTabela = "CanalVenda"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(Codigo.Text)) <> 0 Then iCodigo = CInt(Codigo.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", Nome.Text, STRING_CANAL_VENDA_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", NomeReduzido.Text, STRING_CANAL_VENDA_NOME_REDUZIDO, "NomeReduzido"
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CANAIS_VENDA
    Set Form_Load_Ocx = Me
    Caption = "Canais de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CanalDeVenda"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub


Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

