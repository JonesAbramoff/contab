VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl ItemServico 
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   KeyPreview      =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   6195
   Begin VB.Frame Frame2 
      Caption         =   "Data"
      Height          =   615
      Left            =   105
      TabIndex        =   12
      Top             =   1305
      Width           =   6000
      Begin VB.OptionButton Inicio 
         Caption         =   "Início"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   14
         Top             =   217
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton Fim 
         Caption         =   "Fim"
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
         Left            =   3885
         TabIndex        =   13
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   1740
      Picture         =   "ItemServicoGR.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   285
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ItemServicoGR.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ItemServicoGR.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ItemServicoGR.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ItemServicoGR.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox TextDescricao 
      Height          =   345
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   3
      Top             =   810
      Width           =   4905
   End
   Begin VB.ListBox ListItensServico 
      Height          =   1620
      ItemData        =   "ItemServicoGR.ctx":0A7E
      Left            =   105
      List            =   "ItemServicoGR.ctx":0A80
      TabIndex        =   0
      Top             =   2265
      Width           =   6000
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   270
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
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
      Left            =   405
      TabIndex        =   11
      Top             =   300
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
      Left            =   165
      TabIndex        =   10
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Itens de Serviço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2055
      Width           =   1410
   End
End
Attribute VB_Name = "ItemServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Modulo Item Servico

Option Explicit

'constantes
Const DATA_INICIO = 0
Const DATA_FIM = 1

'Property Variables:
Dim m_Caption As String
Event Unload()

'variaveis globais
Dim iAlterado As Integer

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 97019
    
    Unload Me
    
    Exit Sub

Erro_BotaoFechar_Click:
    
    Select Case gErr

        Case 97019

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a funcao que ira efetuar a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 97016

    'limpa a tela apos a gravacao
    Call Limpa_Tela(Me)
    
    Inicio.Value = True
    
    'fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    'faz com que o registro seja marcado como nao alterado apos a gravacao
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 97016

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objItemServico As ClassItemServico) As Long
'Move os dados da tela para a memoria...

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
    
    'copiando os dados o codigo da tela para a memoria..
    objItemServico.iCodigo = StrParaInt(MaskCodigo.ClipText)
    objItemServico.sDescricao = TextDescricao.Text
    
    'cyntia
    'Se a opção inicio está marcado
    If Inicio.Value = True Then
        'Tipo recebe o código da opção INICIO
        objItemServico.iTipoData = DATA_INICIO
    Else
        'Tipo recebe o código da opção FIM
        objItemServico.iTipoData = DATA_FIM
    End If
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemServico As New ClassItemServico

On Error GoTo Erro_Gravar_Registro

     'Coloca o MouseIcon de Ampulheta durante a gravação
     GL_objMDIForm.MousePointer = vbHourglass

     If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 97025

     If Len(Trim(TextDescricao.Text)) = 0 Then gError 97026
     
     lErro = Move_Tela_Memoria(objItemServico)
     If lErro <> SUCESSO Then gError 97036

     lErro = Trata_Alteracao(objItemServico, objItemServico.iCodigo)
     If lErro <> SUCESSO Then gError 97027

     lErro = CF("ItemServico_Grava", objItemServico)
     If lErro <> SUCESSO Then gError 97028

     'fechando comando de setas
     Call ComandoSeta_Fechar(Me.Name)

     'atualizando a list
     Call ListItensServico_Exclui(objItemServico.iCodigo)
     Call ListItensServico_Adicionar(objItemServico)

     'Coloca o MouseIcon de setinha
     GL_objMDIForm.MousePointer = vbDefault

     Gravar_Registro = SUCESSO
     
     Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 97025
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 97026
             Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 97027, 97028, 97036

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Coloca o MouseIcon de setinha
    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 97017

    'limpando a tela
    Call Limpa_Tela(Me)
    
    Inicio.Value = True
        
    'fechando comando de setas
    Call ComandoSeta_Fechar(Me.Name)
        
    iAlterado = 0
       
    Exit Sub
    
Erro_Botaolimpar_Click:

    Select Case gErr
    
        Case 97017
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCod As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = ItemServico_Codigo_Automatico(iCod)
    If lErro <> SUCESSO Then Error 97009

    MaskCodigo.PromptInclude = False
    MaskCodigo.Text = iCod
    MaskCodigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 97009

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function ItemServico_Codigo_Automatico(iCod As Integer) As Long
'funcao que gera o codigo automatico

Dim lErro As Long

On Error GoTo Erro_ItemServico_Codigo_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_Obter_Inteiro_Automatico", "FatConfig", "NUM_PROX_ITEMSERVICO", "ItemServico", "Codigo", iCod)
    If lErro <> SUCESSO Then gError 97030

    ItemServico_Codigo_Automatico = SUCESSO

    Exit Function

Erro_ItemServico_Codigo_Automatico:

    ItemServico_Codigo_Automatico = gErr

    Select Case gErr

        Case 97030
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objItemServico As New ClassItemServico
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    objItemServico.iCodigo = colCampoValor.Item("Codigo").vValor

    If objItemServico.iCodigo <> 0 Then

        'Carrega o obj com dados de colcampovalor
        objItemServico.iCodigo = colCampoValor.Item("Codigo").vValor
        objItemServico.sDescricao = colCampoValor.Item("Descricao").vValor

        'Traz para tela os dados carregados...
        lErro = Traz_ItemServico_Tela(objItemServico)
        If lErro <> SUCESSO And lErro <> 97011 Then gError 97008
        
        iAlterado = 0

    End If
    
    Exit Sub
    
Erro_Tela_Preenche:

    Select Case gErr
    
        Case 97008
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objItemServico As New ClassItemServico

    'Informa tabela associada à Tela
     sTabela = "ItemServico"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(MaskCodigo.ClipText)) > 0 Then
        objItemServico.iCodigo = StrParaInt(MaskCodigo.ClipText)
    Else
        objItemServico.iCodigo = 0
    End If

    objItemServico.sDescricao = TextDescricao.Text
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objItemServico.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objItemServico.sDescricao, STRING_ITEMSERVICO_DESCRICAO, "Descricao"
    
End Sub

Private Sub ListItensServico_Adicionar(objItemServico As ClassItemServico)

Dim iInd As Integer
    
    For iInd = 0 To ListItensServico.ListCount - 1
        If ListItensServico.ItemData(iInd) > objItemServico.iCodigo Then
            Exit For
        End If
    Next
    
   'Concatena o código com a descrição e adiciona na list...
    ListItensServico.AddItem objItemServico.iCodigo & SEPARADOR & objItemServico.sDescricao, iInd
    ListItensServico.ItemData(ListItensServico.NewIndex) = objItemServico.iCodigo

End Sub

Private Sub ListItensServico_Exclui(iCod As Integer)

Dim iIndice As Integer

    For iIndice = 0 To ListItensServico.ListCount - 1

        If ListItensServico.ItemData(iIndice) = iCod Then

            ListItensServico.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objItemServico As New ClassItemServico
Dim vbMsgRet As VbMsgBoxResult
Dim iCod As Integer

On Error GoTo Erro_BotaoExcluir_Click

    'coloca o cursor no formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Codigo do Item de Servico foi informado
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 97020
    
    objItemServico.iCodigo = StrParaInt(MaskCodigo.ClipText)

    'Verifica se o Item com o codigo em questao existe
    lErro = CF("ItemServico_Le", objItemServico)
    If lErro <> SUCESSO And lErro <> 97035 Then gError 97021

    'Item não está cadastrado
    If lErro = 97035 Then gError 97022

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "EXCLUSAO_ITEMSERVICO", objItemServico.iCodigo)

    If vbMsgRet = vbYes Then

        'exclui o item de servico
        lErro = CF("ItemServico_Exclui", objItemServico)
        If lErro <> SUCESSO Then gError 97023

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        'Exclui o Item de Servico da ListBox
        Call ListItensServico_Exclui(objItemServico.iCodigo)

        'Limpa a Tela
        Call Limpa_Tela(Me)
        
        Inicio.Value = True

        iAlterado = 0

    End If

    'coloca o cursor com formato de setinha
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'coloca o cursor com formato de setinha
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 97020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 97021, 97023

        Case 97022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_NAO_CADASTRADO", gErr, objItemServico.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colItensServico As New Collection
Dim objItemServico As ClassItemServico

On Error GoTo Erro_ItemServico_Form_Load

    'Preenche a ListBox com os Itens de Servico existentes no BD
    lErro = CF("ItemServico_Le_Todos", colItensServico)
    If lErro <> SUCESSO And lErro <> 97007 Then Error 97000

    For Each objItemServico In colItensServico

        ListItensServico.AddItem objItemServico.iCodigo & SEPARADOR & objItemServico.sDescricao
        ListItensServico.ItemData(ListItensServico.NewIndex) = objItemServico.iCodigo

    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_ItemServico_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 97000

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objItemServico As ClassItemServico) As Long
'Inicio do Desenvolvimento 03/05/2001 Tulio

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Item de Servico selecionado, exibir seus dados
    If Not (objItemServico Is Nothing) Then

        'Verifica se Item de Servico existe
        lErro = Traz_ItemServico_Tela(objItemServico)
        If lErro <> 97011 And lErro <> SUCESSO Then gError 97001

        If lErro <> SUCESSO Then
            'Item de Servico não está cadastrado
            
            'Limpa a Tela
            Call Limpa_Tela(Me)
            
            Inicio.Value = True

            MaskCodigo.PromptInclude = False
            MaskCodigo.Text = objItemServico.iCodigo
            MaskCodigo.PromptInclude = True
                   
       End If
          
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 97001

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Function Traz_ItemServico_Tela(objItemServico As ClassItemServico) As Long
    
Dim lErro As Long

On Error GoTo Erro_Traz_ItemServico_Tela

    lErro = CF("ItemServico_Le", objItemServico)
    If lErro <> SUCESSO And lErro <> 97035 Then gError 97010
    
    'se nao achou
    If lErro = 97035 Then gError 97011
    
    'Preenchendo o codigo
    MaskCodigo.PromptInclude = False
    MaskCodigo.Text = objItemServico.iCodigo
    MaskCodigo.PromptInclude = True
    
    'Preenchendo a descricao
    TextDescricao.Text = objItemServico.sDescricao
       
    'Cyntia
    'Se o tipo de data for inicio
    If objItemServico.iTipoData = DATA_INICIO Then
        'marca a opção Data inicio
        Inicio.Value = True
    Else
        'marca a opção data fim
        Fim.Value = True
    End If
    
    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    Traz_ItemServico_Tela = SUCESSO
            
    Exit Function
    
Erro_Traz_ItemServico_Tela:

    Traz_ItemServico_Tela = gErr
    
    Select Case gErr
    
        Case 97010, 97011
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_HISTORICO_PADRAO
    Set Form_Load_Ocx = Me
    Caption = "Itens de Serviço"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ItemServico"

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

Private Sub ListItensServico_DblClick()

Dim lErro As Long
Dim objItemServico As New ClassItemServico

On Error GoTo Erro_ListItensServico_DblClick
    
    'Pega no ItemData o código da Item de Servico Selecionado
    objItemServico.iCodigo = ListItensServico.ItemData(ListItensServico.ListIndex)

    'Traz p/ a tela os dados do item
    lErro = Traz_ItemServico_Tela(objItemServico)
    If lErro <> SUCESSO And lErro <> 97011 Then gError 97014

    'nao encontrou o item...
    'significa q ele esta na lista, mas nao esta no bd.. logo, deve ser
    'removido da list para manter a consistencia list x bd
    If lErro = 97011 Then
        Call ListItensServico_Exclui(objItemServico.iCodigo)
    End If
    
    iAlterado = 0
    
    Exit Sub

Erro_ListItensServico_DblClick:

    Select Case gErr

        Case 97014
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ListItensServico_KeyPress(KeyAscii As Integer)

    If ListItensServico.ListIndex <> -1 Then
        If KeyAscii = ENTER_KEY Then
            Call ListItensServico_DblClick
        End If
    End If

End Sub

Private Sub MaskCodigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
'cyntia
Private Sub Fim_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Inicio_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validade

    If Len(Trim(MaskCodigo.ClipText)) <> 0 Then
        lErro = Inteiro_Critica(MaskCodigo.ClipText)
        If lErro <> SUCESSO Then gError 97012
        
    End If
    
    Cancel = False
    
    Exit Sub
    
Erro_MaskCodigo_Validade:

    Cancel = True
    
    Select Case gErr
    
        Case 97012
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
  
    End Select
    
    Exit Sub

End Sub

Private Sub TextDescricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TextFormulario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

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
