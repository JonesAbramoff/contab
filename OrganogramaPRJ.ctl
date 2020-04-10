VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OrganogramaPRJ 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      Caption         =   "Etapas"
      Height          =   5010
      Left            =   195
      TabIndex        =   18
      Top             =   915
      Width           =   9225
      Begin VB.CommandButton BotaoInsereFilho 
         Caption         =   "Insere Filho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3585
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton BotaoRemove 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4890
         TabIndex        =   5
         Top             =   4545
         Width           =   1215
      End
      Begin VB.CommandButton BotaoInsereIrmao 
         Caption         =   "Insere"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton BotaoFundo 
         Height          =   315
         Left            =   8775
         Picture         =   "OrganogramaPRJ.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3015
         Width           =   330
      End
      Begin VB.CommandButton BotaoSobe 
         Height          =   315
         Left            =   8775
         Picture         =   "OrganogramaPRJ.ctx":0312
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2085
         Width           =   330
      End
      Begin VB.CommandButton BotaoDesce 
         Height          =   315
         Left            =   8775
         Picture         =   "OrganogramaPRJ.ctx":04D4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2550
         Width           =   330
      End
      Begin VB.CommandButton BotaoTopo 
         Height          =   315
         Left            =   8775
         Picture         =   "OrganogramaPRJ.ctx":0696
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1605
         Width           =   330
      End
      Begin MSComctlLib.TreeView TvwOrg 
         Height          =   4125
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   7276
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   453
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7695
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   13
      Top             =   75
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "OrganogramaPRJ.ctx":09A8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "OrganogramaPRJ.ctx":0B26
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OrganogramaPRJ.ctx":1058
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Projeto 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   75
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzidoPRJ 
      Height          =   315
      Left            =   4725
      TabIndex        =   1
      Top             =   75
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Top             =   510
      Width           =   5685
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   225
      TabIndex        =   16
      Top             =   555
      Width           =   930
   End
   Begin VB.Label LabelNomeRedPRJ 
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
      Height          =   315
      Left            =   3255
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   135
      Width           =   1410
   End
   Begin VB.Label LabelProjeto 
      AutoSize        =   -1  'True
      Caption         =   "Projeto:"
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
      TabIndex        =   14
      Top             =   165
      Width           =   675
   End
End
Attribute VB_Name = "OrganogramaPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim sProjetoAnt As String
Dim sNomeProjetoAnt As String

Dim gcolEtapas As Collection
Dim gcolEtapasOriginal As Collection
Dim gobjEtapaAtual As ClassPRJEtapas

Private WithEvents objEventoProjeto As AdmEvento
Attribute objEventoProjeto.VB_VarHelpID = -1

Dim iAlterado As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Organograma"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OrganogramaPRJ"

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

On Error GoTo Erro_Form_UnLoad

    Set objEventoProjeto = Nothing
    
    Set gcolEtapas = Nothing
    Set gcolEtapasOriginal = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182937)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProjeto = New AdmEvento

    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError 189063

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 189063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182938)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objProjeto As ClassProjetos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objProjeto Is Nothing) Then

        If Len(Trim(objProjeto.sCodigo)) > 0 Then
            
            lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
            If lErro <> SUCESSO Then gError 189112
            
            Call Projeto_Validate(bSGECancelDummy)
        Else
            NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
            Call NomeReduzidoPrj_Validate(bSGECancelDummy)
        End If
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 189112

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182939)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoInsereFilho_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoInsereFilho_Click

    lErro = InsereFilho()
    If lErro <> SUCESSO Then gError 182966

    Exit Sub

Erro_BotaoInsereFilho_Click:

    Select Case gErr

        Case 182966

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182940)

    End Select

    Exit Sub

End Sub

Private Sub BotaoInsereIrmao_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoInsereIrmao_Click

    lErro = InsereIrmao()
    If lErro <> SUCESSO Then gError 182967

    Exit Sub

Erro_BotaoInsereIrmao_Click:

    Select Case gErr

        Case 182967

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182941)

    End Select

    Exit Sub

End Sub

Private Sub Calcula_Proxima_Chave(iProxChave As Integer)

Dim sChave As String
Dim objNode1 As Node
Dim iAtual As Integer
Dim lErro As Long

On Error GoTo Erro_Calcula_Proxima_Chave

    iProxChave = 0

    For Each objNode1 In TvwOrg.Nodes

        iAtual = StrParaInt(right(objNode1.Key, Len(objNode1.Key) - 1))

        If iAtual > iProxChave Then iProxChave = iAtual

    Next

     iProxChave = iProxChave + 1

     Exit Sub

Erro_Calcula_Proxima_Chave:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182942)

    End Select

    Exit Sub

End Sub

Private Function InsereFilho() As Long

Dim lErro As Long
Dim objNode As Node, objNodePai As Node
Dim iProxChave As Integer, sChaveTvw As String
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_InsereFilho

    If Not (TvwOrg.SelectedItem Is Nothing) Then

        Set objNodePai = TvwOrg.SelectedItem

        Call Calcula_Proxima_Chave(iProxChave)

        sChaveTvw = "X" & CStr(iProxChave)

        Set objNode = TvwOrg.Nodes.Add(objNodePai.Index, tvwChild, sChaveTvw, SEM_TITULO)
        
    Else

        gError 182968 'Erro . Tem que ter um elemento selecionado

    End If

    objEtapa.iIndiceTvw = objNode.Index
    objEtapa.sChaveTvw = sChaveTvw

    objNode.Selected = True
    
    gcolEtapas.Add objEtapa, sChaveTvw

    InsereFilho = SUCESSO

    Exit Function

Erro_InsereFilho:

    InsereFilho = gErr

    Select Case gErr
        
        Case 182968
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_INSERCAO_FILHO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182943)

    End Select

    Exit Function

End Function

Private Function InsereIrmao() As Long

Dim lErro As Long
Dim objNode As Node, objNodeIrmao As Node
Dim sChaveTvw As String
Dim iProxChave As Integer
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_InsereIrmao

    Call Calcula_Proxima_Chave(iProxChave)

    sChaveTvw = "X" & CStr(iProxChave)

    If Not (TvwOrg.SelectedItem Is Nothing) Then

        Set objNodeIrmao = TvwOrg.SelectedItem

        Set objNode = TvwOrg.Nodes.Add(objNodeIrmao.Index, tvwNext, sChaveTvw, SEM_TITULO)

    Else

        Set objNode = TvwOrg.Nodes.Add(, tvwLast, sChaveTvw, SEM_TITULO)

    End If

    objEtapa.iIndiceTvw = objNode.Index
    objEtapa.sChaveTvw = sChaveTvw

    objNode.Selected = True
    
    gcolEtapas.Add objEtapa, sChaveTvw

    InsereIrmao = SUCESSO

    Exit Function

Erro_InsereIrmao:

    InsereIrmao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182944)

    End Select

    Exit Function

End Function

Private Sub BotaoRemove_Click()
'Remove o nó selecionado da árvore

Dim lErro As Long
Dim iIndice As Integer
Dim objNode As Node
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoRemove_Click

    Set objNode = TvwOrg.SelectedItem

    If objNode Is Nothing Then gError 182969
    
    'Testa se o nó tem filhos
    If objNode.Children > 0 Then

        'Envia aviso perguntando se realmente deseja excluir
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ELEMENTO_TEM_FILHOS")

        If vbMsgRes = vbNo Then gError 182970

    End If

    TvwOrg.Nodes.Remove (objNode.Key)
    
    lErro = Remove_Item_Colecoes
    If lErro <> SUCESSO Then gError 182971

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoRemove_Click:

    Select Case gErr

        Case 182969
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_REMOVER", gErr)
            
        Case 182970, 182971

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182945)

    End Select

    Exit Sub

End Sub

Sub Projeto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Projeto_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Then

        If Len(Trim(Projeto.ClipText)) > 0 Then
            
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 189074
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 182972
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 182973
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
            NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
            Descricao.Caption = objProjeto.sDescricao
        
        Else
        
            Descricao.Caption = ""
            
        End If
        
        sProjetoAnt = Projeto.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 182974
        
    End If
   
    Exit Sub

Erro_Projeto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 182972, 182974, 189074
        
        Case 182973
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 182946)

    End Select

    Exit Sub

End Sub

Sub NomeReduzidoPrj_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim lNumIntDocPRJ As Long

On Error GoTo Erro_NomeReduzidoPrj_Validate

    'Se alterou o projeto
    If sNomeProjetoAnt <> NomeReduzidoPRJ.Text Then

        If Len(Trim(NomeReduzidoPRJ.Text)) > 0 Then
            
            objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le_NomeReduzido", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 182975
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 182976
            
            lNumIntDocPRJ = objProjeto.lNumIntDoc
            
            lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
            If lErro <> SUCESSO Then gError 189113
            
            Descricao.Caption = objProjeto.sDescricao
            
        Else
        
            Descricao.Caption = ""
            
        End If
        
        sNomeProjetoAnt = NomeReduzidoPRJ.Text
        
        lErro = Trata_Projeto(lNumIntDocPRJ)
        If lErro <> SUCESSO Then gError 182977
        
    End If
    
    Exit Sub

Erro_NomeReduzidoPrj_Validate:

    Cancel = True

    Select Case gErr
    
        Case 182975, 182977, 189113
        
        Case 182976
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO3", gErr, objProjeto.sNomeReduzido, objProjeto.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 182947)

    End Select

    Exit Sub

End Sub

Sub LabelProjeto_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_LabelProjeto_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Projeto.ClipText)) <> 0 Then

        lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
        If lErro <> SUCESSO Then gError 189075

        objProjeto.sCodigo = sProjeto

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoProjeto, , "Código")

    Exit Sub

Erro_LabelProjeto_Click:

    Select Case gErr
    
        Case 189075

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182948)

    End Select

    Exit Sub
    
End Sub

Sub LabelNomeRedPRJ_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeRedPRJ_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(NomeReduzidoPRJ.Text)) <> 0 Then

        objProjeto.sNomeReduzido = NomeReduzidoPRJ.Text

    End If

    Call Chama_Tela("ProjetosLista", colSelecao, objProjeto, objEventoProjeto, , "Nome Reduzido")

    Exit Sub

Erro_LabelNomeRedPRJ_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182949)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoProjeto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As ClassProjetos

On Error GoTo Erro_objEventoProjeto_evSelecao

    Set objProjeto = obj1

    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError 189114
    
    NomeReduzidoPRJ.Text = objProjeto.sNomeReduzido
    
    Call Projeto_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoProjeto_evSelecao:

    Select Case gErr
    
        Case 189114

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182950)

    End Select

    Exit Sub

End Sub

Private Function Remove_Item_Colecoes() As Long
'Remove o item, relativo ao nó removido, das coleções

Dim lErro As Long
Dim objNode As Node
Dim colEtapasNovo As New Collection

On Error GoTo Erro_Remove_Item_Colecoes

    'pesquisa na arvore todos os elementos que sobraram e cria um novo colEtapas com estes elementos
    For Each objNode In TvwOrg.Nodes
        colEtapasNovo.Add gcolEtapas.Item(objNode.Key), objNode.Key
    Next

    Set gcolEtapas = colEtapasNovo
    
    Remove_Item_Colecoes = SUCESSO

    Exit Function

Erro_Remove_Item_Colecoes:

    Remove_Item_Colecoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182951)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colEtapasExc As New Collection
Dim objEtapa1 As ClassPRJEtapas
Dim objEtapa2 As ClassPRJEtapas
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim colSaida As New Collection
Dim colCampos As New Collection
Dim colEtapas As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Projeto.ClipText)) = 0 Then gError 182975
    If Len(Trim(NomeReduzidoPRJ.Text)) = 0 Then gError 182976

    For Each objEtapa1 In gcolEtapas
        Set objEtapa2 = New ClassPRJEtapas
        Call objEtapa1.Cria_Copia(objEtapa2)
        colEtapas.Add objEtapa2, objEtapa2.sChaveTvw
    Next

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(colEtapasExc)
    If lErro <> SUCESSO Then gError 182977
    
    If gcolEtapas.Count = 0 Then gError 182926
    
    iIndice1 = 0
    For Each objEtapa1 In gcolEtapas
'        iIndice2 = 0
'        iIndice1 = iIndice1 + 1
'        For Each objEtapa2 In gcolEtapas
'            iIndice2 = iIndice2 + 1
'            If iIndice1 < iIndice2 Then
'                If objEtapa1.sNomeReduzido = objEtapa2.sNomeReduzido Then gError 182935
'            End If
'        Next
        If Len(objEtapa1.sNomeReduzido) > STRING_ETAPAPRJ_NOMEREDUZIDO Then gError 209071
    Next
    
    lErro = Move_Posicoes_Arvore
    If lErro <> SUCESSO Then gError 185824
   
    colCampos.Add "iNovo"
    colCampos.Add "iPosicao"
    
    Call Ordena_Colecao(gcolEtapas, colSaida, colCampos)

    Set gcolEtapas = colSaida
       
    'Grava o/a Projetos no Banco de Dados
    lErro = CF("OrganogramaPRJ_Grava", colSaida, colEtapasExc)
    If lErro <> SUCESSO Then
        'Desfaz o que a gravação fez
        Set gcolEtapas = colEtapas
        gError 182978
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 182926
            Call Rotina_Erro(vbOKOnly, "ERRO_ORGANOGRAMA_SEM_ETAPAS", gErr)
            
        Case 182935
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPAS_NOME_REPETIDO", gErr)
        
        Case 182975
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus

        Case 182976
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_PRJ_NAO_PREENCHIDO", gErr)
            NomeReduzidoPRJ.SetFocus
            
        Case 182977, 182978, 185824
        
        Case 209071
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_ETAPAPRJ_GRANDE", gErr, STRING_ETAPAPRJ_NOMEREDUZIDO, objEtapa1.sNomeReduzido, Len(objEtapa1.sNomeReduzido))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182952)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colEtapasExc As Collection) As Long
'Remove o item, relativo ao nó removido, das coleções

Dim lErro As Long
Dim objNode As Node
Dim objNodeFilho As Node
Dim objEtapa As ClassPRJEtapas
Dim objEtapaAux As ClassPRJEtapas
Dim objEtapaFilha As ClassPRJEtapas
Dim objProjeto As New ClassProjetos
Dim iSeq As Integer
Dim bAchou As Boolean
Dim iIndice As Integer
Dim objNodeNext As Node
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189077

    objProjeto.sCodigo = sProjeto
    objProjeto.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjeto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 182979

    'Monta a coleção de etapas excluídas
    For Each objEtapa In gcolEtapasOriginal
        bAchou = False
        For Each objEtapaAux In gcolEtapas
            If objEtapa.sChaveTvw = objEtapaAux.sChaveTvw Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then colEtapasExc.Add objEtapa
    Next

    iSeq = 0
    For iIndice = 1 To TvwOrg.Nodes.Count
        If iIndice = 1 Then
            Set objNodeNext = TvwOrg.Nodes.Item(1).FirstSibling
        Else
            Set objNodeNext = objNodeNext.Next
        End If
        
        If objNodeNext Is Nothing Then Exit For
        
        If objNodeNext.Parent Is Nothing Then
            iSeq = iSeq + 1
            objNodeNext.Tag = CStr(iSeq)
        End If
    Next
    
    For Each objNode In TvwOrg.Nodes
        Set objEtapa = gcolEtapas.Item(objNode.Key)
        iSeq = 0
        For Each objNodeFilho In TvwOrg.Nodes
            If Not (objNodeFilho.Parent Is Nothing) Then
                If objNodeFilho.Parent.Key = objNode.Key Then
                    iSeq = iSeq + 1
                    objNodeFilho.Tag = objNode.Tag & "." & CStr(iSeq)
                    Set objEtapaFilha = gcolEtapas.Item(objNodeFilho.Key)
                    objEtapaFilha.sCodigoPaiOrg = objNode.Tag
                End If
            End If
        Next
        
        objEtapa.sReferencia = "" 'objNode.Tag
        objEtapa.sCodigo = objNode.Tag
        objEtapa.sNomeReduzido = objNode.Text
        objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
        objEtapa.dtDataFim = DATA_NULA
        objEtapa.dtDataFimReal = DATA_NULA
        objEtapa.dtDataInicio = DATA_NULA
        objEtapa.dtDataInicioReal = DATA_NULA
        objEtapa.iNivel = Obtem_Nivel(objEtapa.sCodigo)
        objEtapa.iSeq = Obtem_Seq(objEtapa.sCodigo)
        
        'Se não tinha código anterior marca o registro como novo
        'é utilizado para manter uma ordem correta de atualização.
        If Len(Trim(objEtapa.sCodigoAnt)) = 0 Then
            objEtapa.iNovo = MARCADO
        Else
            objEtapa.iNovo = DESMARCADO
        End If
        
    Next
    
    For Each objNode In TvwOrg.Nodes
        objNode.Tag = ""
    Next
       
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 182979, 189077

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182953)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 182980

    'Limpa Tela
    Call Limpa_Tela_OrganogramaPRJ

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 182980

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182954)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182955)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 182981

    Call Limpa_Tela_OrganogramaPRJ

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 182981

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182956)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_OrganogramaPRJ() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_OrganogramaPRJ

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    TvwOrg.Nodes.Clear
    Set gcolEtapas = New Collection
    
    Descricao.Caption = ""
    
    sProjetoAnt = ""
    sNomeProjetoAnt = ""

    iAlterado = 0

    Limpa_Tela_OrganogramaPRJ = SUCESSO

    Exit Function

Erro_Limpa_Tela_OrganogramaPRJ:

    Limpa_Tela_OrganogramaPRJ = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182957)

    End Select

    Exit Function

End Function

Function Trata_Projeto(ByVal lNumIntDocPRJ As Long) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Trata_Projeto
    vbResult = vbNo
    If Len(Trim(sNomeProjetoAnt)) > 0 Then vbResult = Rotina_Aviso(vbYesNo, "AVISO_ORGANOGRAMA_TROCA_PRJ")
    If vbResult = vbNo Then
        TvwOrg.Nodes.Clear
        Set gcolEtapas = New Collection
        Set gcolEtapasOriginal = New Collection
        
        If lNumIntDocPRJ <> 0 Then
    
            objProjeto.lNumIntDoc = lNumIntDocPRJ
        
            lErro = CF("PRJEtapas_Le_Projeto", objProjeto)
            If lErro <> SUCESSO Then gError 182982
    
            lErro = Carrega_Arvore(objProjeto)
            If lErro <> SUCESSO Then gError 182983
            
        End If
    End If
    sProjetoAnt = Projeto.Text
    sNomeProjetoAnt = NomeReduzidoPRJ.Text

    Trata_Projeto = SUCESSO

    Exit Function

Erro_Trata_Projeto:

    Trata_Projeto = gErr

    Select Case gErr
    
        Case 182982, 182983

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182958)

    End Select

    Exit Function

End Function

Function Carrega_Arvore(ByVal objProjeto As ClassProjetos) As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim objNode As Node
Dim lErro As Long
Dim sChaveTvw As String
Dim iIndicePai As Integer
Dim sTexto As String
Dim objEtapa As ClassPRJEtapas
Dim objEtapaAux As ClassPRJEtapas
Dim iProxChave As Integer

On Error GoTo Erro_Carrega_Arvore
    
    For Each objEtapa In objProjeto.colEtapas

        'Texto que identificará a nova Etapa que está sendo incluida
        sTexto = objEtapa.sNomeReduzido
        
        'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
        Call Calcula_Proxima_Chave(iProxChave)

        sChaveTvw = "X" & CStr(iProxChave)

        If objEtapa.lNumIntDocEtapaPaiOrg = 0 Then

            Set objNode = TvwOrg.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)

        Else

            For Each objEtapaAux In objProjeto.colEtapas
            
                If objEtapa.lNumIntDocEtapaPaiOrg = objEtapaAux.lNumIntDoc Then
                    iIndicePai = objEtapaAux.iIndiceTvw
                    Exit For
                End If

            Next

            Set objNode = TvwOrg.Nodes.Add(iIndicePai, tvwChild, sChaveTvw, sTexto)

        End If
                
        TvwOrg.Nodes.Item(objNode.Index).Expanded = True
        
        objEtapa.sCodigoAnt = objEtapa.sCodigo
        objEtapa.iIndiceTvw = objNode.Index
        objEtapa.sChaveTvw = sChaveTvw
        
        gcolEtapas.Add objEtapa, sChaveTvw
        gcolEtapasOriginal.Add objEtapa, sChaveTvw
        
        objNode.Tag = sChaveTvw
        
    Next

    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182959)

    End Select

    Exit Function

End Function

Private Sub BotaoSobe_Click()

Dim lErro As Long
Dim objNode As Node
Dim objNode1 As Node
Dim objEtapa As ClassPRJEtapas
Dim colAux As New Collection

On Error GoTo Erro_BotaoSobe_Click

    If TvwOrg.SelectedItem Is Nothing Then gError 182984
    
    lErro = Move_Tela_Memoria(colAux)
    If lErro <> SUCESSO Then gError 182985

    Set objEtapa = gcolEtapas.Item(TvwOrg.SelectedItem.Key)
    Set objNode = TvwOrg.SelectedItem

    'Verifica se tem irmão posicionado anteriormente
    If objNode.Previous Is Nothing Then gError 182986

    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.Previous, tvwPrevious)
    If lErro <> SUCESSO Then gError 182987

    Exit Sub

Erro_BotaoSobe_Click:

    Select Case gErr

        Case 182984
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", gErr)
            
        Case 182985, 182987

        Case 182986
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ACIMA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182960)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoDesce_Click()

Dim lErro As Long
Dim objNode As Node
Dim objEtapa As ClassPRJEtapas
Dim colAux As New Collection

On Error GoTo Erro_BotaoDesce_Click

    If TvwOrg.SelectedItem Is Nothing Then gError 182988

    Set objEtapa = gcolEtapas.Item(TvwOrg.SelectedItem.Key)

    'Salva dados do nó corrente do grid e da árvore nos obj's
    lErro = Move_Tela_Memoria(colAux)
    If lErro <> SUCESSO Then gError 182989

    Set objNode = TvwOrg.SelectedItem

    'Verifica se tem irmão posicionado posteriormente
    If objNode.Next Is Nothing Then gError 182990
    
    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.Next, tvwNext)
    If lErro <> SUCESSO Then gError 44610

    Exit Sub

Erro_BotaoDesce_Click:

    Select Case gErr

        Case 182988
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", gErr)
            
        Case 182989, 182991

        Case 182990
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ABAIXO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182961)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFundo_Click()

Dim lErro As Long
Dim objNode As Node
Dim objEtapa As ClassPRJEtapas
Dim colAux As New Collection

On Error GoTo Erro_BotaoFundo_Click

    If TvwOrg.SelectedItem Is Nothing Then gError 182992

    Set objEtapa = gcolEtapas.Item(TvwOrg.SelectedItem.Key)

    'Salva dados do nó corrente do grid e da árvore nos obj's
    lErro = Move_Tela_Memoria(colAux)
    If lErro <> SUCESSO Then gError 182993

    Set objNode = TvwOrg.SelectedItem

    'Verifica se tem irmão posicionado posteriormente
    If objNode.Next Is Nothing Then gError 182994
    
    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.LastSibling, tvwNext)
    If lErro <> SUCESSO Then gError 182995

    Exit Sub

Erro_BotaoFundo_Click:

    Select Case gErr

        Case 182992
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", gErr)
            
        Case 182993, 182995

        Case 182994
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ABAIXO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182962)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTopo_Click()

Dim lErro As Long
Dim objNode As Node
Dim objEtapa As ClassPRJEtapas
Dim colAux As New Collection

On Error GoTo Erro_BotaoTopo_Click

    If TvwOrg.SelectedItem Is Nothing Then gError 182996

    Set objEtapa = gcolEtapas.Item(TvwOrg.SelectedItem.Key)

    'Salva dados do nó corrente do grid e da árvore nos obj's
    lErro = Move_Tela_Memoria(colAux)
    If lErro <> SUCESSO Then gError 182997

    Set objNode = TvwOrg.SelectedItem

    'Verifica se tem irmão posicionado posteriormente
    If objNode.Previous Is Nothing Then gError 182998
    
    'executa a movimentação do elemento
    lErro = Executa_Movimentacao(objNode, objNode.FirstSibling, tvwPrevious)
    If lErro <> SUCESSO Then gError 182999

    Exit Sub

Erro_BotaoTopo_Click:

    Select Case gErr

        Case 182996
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_NAO_SELECIONADO_MOV_ARV", gErr)
            
        Case 182997, 182999

        Case 182998
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_SELECIONADO_NAO_MOV_ACIMA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182963)

    End Select

    Exit Sub

End Sub

Private Function Executa_Movimentacao(objNode As Node, objNodeParente As Node, iRelacao As Integer) As Long
'executa a movimentação do nó objNode para o lado de objNodeParente. Se vai ficar acima ou abaixo de objNodeParente depende do valor de iRelacao que pode ser tvwPrevious ou tvwNext
    
Dim lErro As Long
Dim objNodeNovo As Node
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_Executa_Movimentacao

    lErro = Move_Posicoes_Arvore
    If lErro <> SUCESSO Then gError 185816

    Set objEtapa = gcolEtapas(objNode.Key)

    TvwOrg.Nodes.Remove (objNode.Key)
    
    lErro = Reposiciona_No_Arvore(objEtapa, objNodeParente, iRelacao)
    If lErro <> SUCESSO Then gError 185000
    
    Executa_Movimentacao = SUCESSO

    Exit Function

Erro_Executa_Movimentacao:

    Executa_Movimentacao = gErr
    
    Select Case gErr
    
        Case 185000, 185816
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182964)
    
    End Select
    
    Exit Function
    
End Function

Private Function Reposiciona_No_Arvore(objEtapa As ClassPRJEtapas, objNodeParente As Node, iRelacao As Integer) As Long
'carrega a arvore com o nó e seus descentes a partir da no passado como parametro. Se é posicionado acima ou abaixo do no depende do valor do parametro iRelacao.

Dim lErro As Long
Dim objEtapa1 As ClassPRJEtapas
Dim colPais As New Collection
Dim iNivel As Integer
Dim iPosicao As Integer
Dim objNode1 As Node
Dim objEtapaAux As ClassPRJEtapas

On Error GoTo Erro_Reposiciona_No_Arvore

    'coloca o nó que está sendo movido na nova posicao
    Set objNode1 = TvwOrg.Nodes.Add(objNodeParente, iRelacao, objEtapa.sChaveTvw, objEtapa.sNomeReduzido)
    objEtapa.iIndiceTvw = objNode1.Index
    objNode1.Expanded = True

    objNode1.Selected = True

    'proxima posicao a ser pesquisada
    iPosicao = objEtapa.iPosicao + 1
    For Each objEtapa1 In gcolEtapas
        If objEtapa1.iPosicao = iPosicao Then Exit For
    Next
    
    colPais.Add objNode1
    
    If Not (objEtapa1 Is Nothing) Then
    
        'enquanto os elementos forem descendentes do elemento que está sendo movido
        Do While Obtem_Nivel(objEtapa1.sCodigo) > Obtem_Nivel(objEtapa.sCodigo)
        
            Set objNode1 = TvwOrg.Nodes.Add(colPais.Item(Obtem_Nivel(objEtapa1.sCodigo) - Obtem_Nivel(objEtapa.sCodigo)), tvwChild, objEtapa1.sChaveTvw, objEtapa1.sNomeReduzido)
            objEtapa1.iIndiceTvw = objNode1.Index
            objNode1.Expanded = True
    
            For iNivel = (objEtapa1.iNivel - objEtapa.iNivel + 1) To colPais.Count
                colPais.Remove (objEtapa1.iNivel - objEtapa.iNivel + 1)
            Next
    
            colPais.Add objNode1
    
            'vai tratar a proxima posicao
            iPosicao = iPosicao + 1
            'procura o elemento na proxima posicao
            For Each objEtapa1 In gcolEtapas
                If objEtapa1.iPosicao = iPosicao Then Exit For
            Next
            
            If objEtapa1 Is Nothing Then Exit Do
    
        Loop

    End If
    
    lErro = Move_Posicoes_Arvore
    If lErro <> SUCESSO Then gError 185817

    Reposiciona_No_Arvore = SUCESSO

    Exit Function

Erro_Reposiciona_No_Arvore:

    Reposiciona_No_Arvore = gErr

    Select Case gErr
    
        Case 185817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182965)

    End Select

    Exit Function

End Function

Private Sub TvwOrg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then

        If gcolEtapas.Count > 0 Then

            'Seta objTela como a Tela de Baixas a Receber
            Set PopUpMenuPRJ.objTela = Me
            
            Set gobjEtapaAtual = gcolEtapas.Item(TvwOrg.SelectedItem.Key)
    
            'Chama o Menu PopUp
            PopUpMenuPRJ.PopupMenu PopUpMenuPRJ.mnuGrid, vbPopupMenuRightButton
    
            'Limpa o objTela
            Set PopUpMenuPRJ.objTela = Nothing
            
        End If

    End If
    
End Sub

Public Function mnuTvwAbrirEtapa_Click() As Long

Dim lErro As Long
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_mnuTvwAbrirEtapa_Click

    If Len(Trim(gobjEtapaAtual.sCodigoAnt)) = 0 Then gError 185013
    
    objEtapa.lNumIntDocPRJ = gobjEtapaAtual.lNumIntDocPRJ
    objEtapa.sCodigo = gobjEtapaAtual.sCodigoAnt

    Call Chama_Tela("EtapaPRJ", objEtapa)
    
    mnuTvwAbrirEtapa_Click = SUCESSO
    
    Exit Function

Erro_mnuTvwAbrirEtapa_Click:

    mnuTvwAbrirEtapa_Click = gErr

    Select Case gErr
    
        Case 185013
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO3", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185011)

    End Select

    Exit Function
    
End Function

Public Function mnuTvwDocRelacs_Click() As Long

Dim lErro As Long
Dim objProjetoInfo As New ClassProjetoInfo
Dim colSelecao As New Collection

On Error GoTo Erro_mnuTvwDocRelacs_Click

    If Len(Trim(gobjEtapaAtual.sCodigoAnt)) = 0 Then gError 185020
    
    colSelecao.Add gobjEtapaAtual.lNumIntDocPRJ
    colSelecao.Add gobjEtapaAtual.lNumIntDoc

    Call Chama_Tela("ProjetoInfoLista", colSelecao, objProjetoInfo, Nothing, "NumIntDocPRJ = ? AND NumIntDocEtapa = ?")
    
    mnuTvwDocRelacs_Click = SUCESSO
    
    Exit Function

Erro_mnuTvwDocRelacs_Click:

    mnuTvwDocRelacs_Click = gErr

    Select Case gErr
        
        Case 185020
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO3", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185012)

    End Select

    Exit Function
    
End Function

Public Function Obtem_Nivel(ByVal sMascara As String) As Long

Dim lErro As Long
Dim iPos As Integer
Dim iNivel As Integer

    iNivel = 0

    Do While True
    
        iPos = InStr(iPos + 1, sMascara, ".")
    
        If iPos = 0 Then Exit Do
            
        iNivel = iNivel + 1
    
    Loop

    Obtem_Nivel = iNivel
    
End Function

Public Function Obtem_Seq(ByVal sMascara As String) As Long

Dim lErro As Long
Dim iPos As Integer
Dim iSeq As Integer
Dim iPosAnt As Integer
    
    Do While True
    
        iPosAnt = iPos
    
        iPos = InStr(iPos + 1, sMascara, ".")
    
        If iPos = 0 Then Exit Do
    
    Loop
    
    iSeq = StrParaInt(right(sMascara, Len(sMascara) - iPosAnt))

    Obtem_Seq = iSeq
    
End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objProjetos As New ClassProjetos
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Projetos"

    lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
    If lErro <> SUCESSO Then gError 189078

    objProjetos.sCodigo = sProjeto
    objProjetos.iFilialEmpresa = giFilialEmpresa
    
    'Le
    lErro = CF("Projetos_Le", objProjetos)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181320

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objProjetos.sCodigo, STRING_PRJ_CODIGO, "Codigo"
    'Filtros para o Sistema de Setas
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 181320, 189078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181593)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objProjetos As New ClassProjetos

On Error GoTo Erro_Tela_Preenche

    objProjetos.sCodigo = colCampoValor.Item("Codigo").vValor
    objProjetos.iFilialEmpresa = giFilialEmpresa

    lErro = Retorno_Projeto_Tela(Projeto, objProjetos.sCodigo)
    If lErro <> SUCESSO Then gError 189115
    
    Call Projeto_Validate(bSGECancelDummy)

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr
    
        Case 189115

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181594)

    End Select

    Exit Function

End Function

Private Function Move_Posicoes_Arvore() As Long

Dim lErro As Long
Dim objNode As Node
Dim objNode1 As Node
Dim iPosicao As Integer

On Error GoTo Erro_Move_Posicoes_Arvore

    If TvwOrg.Nodes.Count > 0 Then

        Set objNode = TvwOrg.Nodes.Item(1)

        If Not (objNode.Root Is Nothing) Then Set objNode = objNode.Root

        If Not (objNode.FirstSibling Is Nothing) Then Set objNode = objNode.FirstSibling

        lErro = Armazena_Posicao_Arvore(objNode, iPosicao, 1)
        If lErro <> SUCESSO Then gError 185820

    End If

    Move_Posicoes_Arvore = SUCESSO

    Exit Function

Erro_Move_Posicoes_Arvore:

    Move_Posicoes_Arvore = gErr

    Select Case gErr

        Case 185820

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185821)

    End Select

    Exit Function

End Function

Private Function Armazena_Posicao_Arvore(objNode As Node, iPosicao As Integer, ByVal iNivel As Integer) As Long
'descobre a posicao do no na arvore e armazena-a e pesquisa o proximo no (seja um filho ou um irmao)

Dim lErro As Long
Dim objEtapa As ClassPRJEtapas

On Error GoTo Erro_Armazena_Posicao_Arvore

    Do While Not (objNode Is Nothing)

        iPosicao = iPosicao + 1
    
        Set objEtapa = gcolEtapas.Item(objNode.Key)
    
        objEtapa.iPosicao = iPosicao
        objEtapa.iNivel = iNivel
        objEtapa.sNomeReduzido = objNode.Text
    
        If Not (objNode.Child Is Nothing) Then
    
            lErro = Armazena_Posicao_Arvore(objNode.Child, iPosicao, iNivel + 1)
            If lErro <> SUCESSO Then gError 185818
    
        End If
            
        Set objNode = objNode.Next

    Loop

    Armazena_Posicao_Arvore = SUCESSO

    Exit Function

Erro_Armazena_Posicao_Arvore:

    Armazena_Posicao_Arvore = gErr

    Select Case gErr

        Case 185818

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185819)

    End Select

    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is NomeReduzidoPRJ Then
            Call LabelNomeRedPRJ_Click
        End If

    End If
    
End Sub
