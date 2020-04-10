VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CidadeCadastroOcx 
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5730
   ForeColor       =   &H00000000&
   ScaleHeight     =   6270
   ScaleWidth      =   5730
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3420
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CidadeCadastroOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CidadeCadastroOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CidadeCadastroOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CidadeCadastroOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListCodigo 
      Height          =   3960
      ItemData        =   "CidadeCadastroOcx.ctx":0994
      Left            =   165
      List            =   "CidadeCadastroOcx.ctx":0996
      TabIndex        =   1
      Top             =   2190
      Width           =   5340
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1785
      Picture         =   "CidadeCadastroOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Numeração Automática"
      Top             =   435
      Width           =   300
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1185
      TabIndex        =   2
      Top             =   420
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   990
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CodIBGE 
      Height          =   315
      Left            =   1185
      TabIndex        =   4
      Top             =   1515
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "#######"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   1965
      Width           =   690
   End
   Begin VB.Label DescIBGE 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1995
      TabIndex        =   13
      Top             =   1515
      Width           =   3555
   End
   Begin VB.Label LabelCodIBGE 
      AutoSize        =   -1  'True
      Caption         =   "Código IBGE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   15
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   1575
      Width           =   1140
   End
   Begin VB.Label LabelDescricao 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   195
      TabIndex        =   6
      Top             =   1035
      Width           =   915
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
      Left            =   480
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   465
      Width           =   660
   End
End
Attribute VB_Name = "CidadeCadastroOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Eventos do browse
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoMunicIBGE As AdmEvento
Attribute objEventoMunicIBGE.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

'*** FUNÇÕES DE APOIO A TELA - INÍCIO ***
Private Sub ListaCodigo_Inclui(objCidades As ClassCidades)
'Adiciona na ListBox

Dim iIndice As Integer

    For iIndice = 0 To ListCodigo.ListCount - 1
        If ListCodigo.ItemData(iIndice) > objCidades.iCodigo Then Exit For
    Next

    'Adiciona na lista de código e descrição
    ListCodigo.AddItem objCidades.iCodigo & SEPARADOR & objCidades.sDescricao, iIndice
    ListCodigo.ItemData(ListCodigo.NewIndex) = objCidades.iCodigo
    
    Exit Sub

End Sub

Private Sub ListaCodigo_Exclui(objCidades As ClassCidades)
'Percorre a ListBox para remover o tipo caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To ListCodigo.ListCount - 1

        If ListCodigo.ItemData(iIndice) = objCidades.iCodigo Then
            ListCodigo.RemoveItem (iIndice)
            Exit For
        End If
    Next
    
    Exit Sub

End Sub

Function Move_Tela_Memoria(objCidades As ClassCidades) As Long

    'Move os dados da tela para memória
    objCidades.iCodigo = StrParaInt(Codigo.Text)
    objCidades.sDescricao = Descricao.Text
    objCidades.sCodIBGE = Trim(CodIBGE.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

End Function

Function Traz_CidadeCadastro_Tela(objCidades As ClassCidades) As Long
'Preenche a tela com as informações do banco

On Error GoTo Erro_Traz_CidadeCadastro_Tela

    'Limpa a tela
    Call Limpa_Tela_CidadeCadastro

    'Mostra os dados na tela
    Codigo.Text = CStr(objCidades.iCodigo)
    Descricao.Text = objCidades.sDescricao
    
    CodIBGE.PromptInclude = False
    CodIBGE.Text = objCidades.sCodIBGE
    CodIBGE.PromptInclude = True
    Call CodIBGE_Validate(bSGECancelDummy)

    iAlterado = 0

    Exit Function

Erro_Traz_CidadeCadastro_Tela:

    Traz_CidadeCadastro_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144597)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Verifica se dados do CidadeCadastro necessários foram preenchidos
'Grava CidadeCadastro no BD
'Atualiza List

Dim lErro As Long
Dim objCidades As New ClassCidades

On Error GoTo Erro_Gravar_Registro

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 125018

    'Verifica se a Descricao foi preenchida
    If Len(Trim(Descricao.Text)) = 0 Then gError 125019

    lErro = Move_Tela_Memoria(objCidades)
    If lErro <> SUCESSO Then gError 125020
    
    lErro = CF("Cidade_Grava", objCidades)
    If lErro <> SUCESSO Then gError 125021

    'Exclui da ListBox
    Call ListaCodigo_Exclui(objCidades)

    'Inclui na ListBox
    Call ListaCodigo_Inclui(objCidades)

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
        
        Case 125018
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 125019
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 125020, 125021
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144598)

    End Select
    
    Exit Function

End Function

Function Limpa_Tela_CidadeCadastro()

    DescIBGE.Caption = ""

    'Limpa a tela
    Call Limpa_Tela(Me)

End Function
'*** FUNÇÕES DE APOIO A TELA - INÍCIO ***

'*** FUNÇÕES DE INICIALIZAÇÃO DA TELA - INÍCIO ***
Function Trata_Parametros(Optional objCidades As ClassCidades) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de CidadeCadastro

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objCidades Is Nothing) Then

        lErro = CF("Cidade_Le", objCidades)
        If lErro <> SUCESSO And lErro <> 125041 Then gError 125046

        If lErro = SUCESSO Then

            Call Traz_CidadeCadastro_Tela(objCidades)

        Else
        
            If objCidades.iCodigo <> 0 Then
                Codigo.Text = CStr(objCidades.iCodigo)
            Else
                Codigo.Text = ""
            End If
            
            Descricao.Text = objCidades.sDescricao
        End If

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 125046
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADES_NAO_CADASTRADO", gErr, objCidades.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144599)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objCidades As New ClassCidades
Dim colCidades As New Collection

On Error GoTo Erro_Form_Load
    
    'Inicializa o Browse
    Set objEventoCodigo = New AdmEvento
    Set objEventoMunicIBGE = New AdmEvento

    'Preenche a listbox ListCadastro
    lErro = CF("Cidade_LeTodos", colCidades)
    If lErro <> SUCESSO Then gError 125047

    'preenche a listCadastro com os objCidades
    For Each objCidades In colCidades

        ListCodigo.AddItem objCidades.iCodigo & SEPARADOR & objCidades.sDescricao
        ListCodigo.ItemData(ListCodigo.NewIndex) = objCidades.iCodigo

    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 125047

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144600)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub
'*** FUNÇÕES DE INICIALIZAÇÃO DA TELA - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCidades As New ClassCidades

On Error GoTo Erro_BotaoExcluir_Click:

    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 125048

    objCidades.iCodigo = CInt(Codigo.Text)

    'Envia aviso perguntando se realmente deseja excluir CidadeCadastro
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CIDADES", objCidades.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui o CidadesCadastro
        lErro = CF("Cidade_Exclui", objCidades)
        If lErro <> SUCESSO And lErro = 125034 Then gError 125049

        If lErro = 125034 Then gError 125050

        'Exclui da List
        Call ListaCodigo_Exclui(objCidades)

        'Limpa a tela
        Call Limpa_Tela_CidadeCadastro

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

    End If

    iAlterado = 0

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125048
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 125049

        Case 125050
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADES_NAO_CADASTRADO", gErr, objCidades.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144601)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 125051

    'Limpa a tela
    Call Limpa_Tela_CidadeCadastro

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144602)

    End Select
    
    iAlterado = 0

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 125052

    'Limpa a tela
    Call Limpa_Tela_CidadeCadastro

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 125052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144603)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código disponível para CidadeCadastro
    lErro = CF("Config_Obter_Inteiro_Automatico", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 125053
    
    'Coloca o Código obtido na tela
    Codigo.Text = iCodigo
        
    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 125053
            'Erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144604)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub ListCodigo_DblClick()

Dim lErro As Long
Dim objCidades As New ClassCidades

On Error GoTo Erro_ListCodigo_DblClick

    objCidades.iCodigo = ListCodigo.ItemData(ListCodigo.ListIndex)

    'Lê o Cadastro
    lErro = CF("Cidade_Le", objCidades)
    If lErro <> SUCESSO And lErro <> 125041 Then gError 125054

    'Verifica se o codigo não está cadastrado
    If lErro = 125041 Then gError 125055
    
    lErro = Traz_CidadeCadastro_Tela(objCidades)
    If lErro <> SUCESSO Then gError 125056

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_ListCodigo_DblClick:

    Select Case gErr

        Case 125054, 125056

        Case 125055
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADES_NAO_CADASTRADO", gErr, objCidades.iCodigo)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144605)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objCidades As New ClassCidades
Dim colSelecao As New Collection
    
    'Preenche na memória o Código passado
    If Len(Trim(Codigo.ClipText)) > 0 Then objCidades.iCodigo = Codigo.Text

    Call Chama_Tela("CidadeLista", colSelecao, objCidades, objEventoCodigo)

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO GOTFOCUS DOS CONTROLES - INÍCIO ***
Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES - FIM ***

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera as variáveis globais
    Set objEventoCodigo = Nothing
    Set objEventoMunicIBGE = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'*** FUNÇÕES DO SISTEMA DE SETA - INÍCIO ***
Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCidades As New ClassCidades

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objCidades.iCodigo = CStr(colCampoValor.Item("Codigo").vValor)
    objCidades.sDescricao = colCampoValor.Item("Descricao").vValor

    lErro = CF("Cidade_Le", objCidades)
    If lErro <> SUCESSO And lErro <> 125041 Then gError 125056

    lErro = Traz_CidadeCadastro_Tela(objCidades)
    If lErro <> SUCESSO Then gError 125056

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 125056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144606)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim lErro As Long
Dim objCidades As New ClassCidades

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada a tela
    sTabela = "Cidades"

    lErro = Move_Tela_Memoria(objCidades)
    If lErro <> SUCESSO Then gError 125057

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCidades.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objCidades.sDescricao, STRING_CIDADE, "Descricao"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 125057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144607)

    End Select

    Exit Sub

End Sub
'*** FUNÇÕES DO SISTEMA DE SETA - FIM ***

'*** FUNÇÕES DO BROWSE - INÍCIO

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidades As New ClassCidades
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoCodigo_evSelecao
    
    Set objCidades = obj1

    lErro = Traz_CidadeCadastro_Tela(objCidades)
    If lErro <> SUCESSO And lErro <> 125041 Then gError 125058
    
    If lErro = 125041 Then gError 125059

    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 125058
        
        Case 125059
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CIDADES_NAO_CADASTRADO", gErr, objCidades.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144608)

    End Select
    
    iAlterado = 0

    Exit Sub

End Sub
'*** FUNÇÕES DO BROWSE - FIM ***

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cidades"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CidadeCadastro"
    
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

Private Sub CodIBGE_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodIBGE_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodIBGE, iAlterado)
End Sub

Private Sub CodIBGE_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDescIBGE As String

On Error GoTo Erro_CodIBGE_Validate

    If Len(Trim(CodIBGE.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(CodIBGE.Text)
        If lErro <> SUCESSO Then gError 57975
        
        lErro = CF("Le_Campo_Tabela", "IBGEMunicipios", "DescIBGE", TIPO_STR, "CodIBGE", CodIBGE.Text, sDescIBGE, DescIBGE)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200981
        
        If lErro <> SUCESSO Then gError 200982
    Else
    
        DescIBGE.Caption = ""
        
    End If
    
    Exit Sub
    
Erro_CodIBGE_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 57975, 200981 'Erro tratado na rotina chamada
        
        Case 200982
            Call Rotina_Erro(vbOKOnly, "ERRO_MUNIC_CODIBGE_NAO_CADASTRADO", gErr, CodIBGE.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164299)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelCodIBGE_Click()

Dim colSelecao As New Collection
Dim objIBGEMunicipios As New ClassIBGEMunicipios

On Error GoTo Erro_LabelCodIBGE_Click

    objIBGEMunicipios.sMunic = Descricao.Text

    'Chama Tela TituloReceberLista
    Call Chama_Tela("IBGEMunicipiosLista", colSelecao, objIBGEMunicipios, objEventoMunicIBGE)

    Exit Sub

Erro_LabelCodIBGE_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200973)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMunicIBGE_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objIBGEMunicipios As ClassIBGEMunicipios

On Error GoTo Erro_objEventoMunicIBGE_evSelecao

    Set objIBGEMunicipios = obj1
    
    CodIBGE.PromptInclude = False
    CodIBGE.Text = objIBGEMunicipios.sCodUF & objIBGEMunicipios.sCodMunic
    CodIBGE.PromptInclude = True
    Call CodIBGE_Validate(bSGECancelDummy)
    
    Descricao.Text = left(objIBGEMunicipios.sMunic, STRING_CIDADE)
    
    Exit Sub

Erro_objEventoMunicIBGE_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200974)

    End Select

    Exit Sub

End Sub
