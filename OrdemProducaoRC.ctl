VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl OrdemProducaoRC 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   3405
      Picture         =   "OrdemProducaoRC.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5355
      Width           =   1005
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   5175
      Picture         =   "OrdemProducaoRC.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5370
      Width           =   1005
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todos"
      Height          =   645
      Index           =   0
      Left            =   2115
      Picture         =   "OrdemProducaoRC.ctx":025C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4575
      Width           =   1830
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todas"
      Height          =   645
      Left            =   210
      Picture         =   "OrdemProducaoRC.ctx":143E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4575
      Width           =   1830
   End
   Begin MSComctlLib.TreeView OPs 
      Height          =   3915
      Left            =   225
      TabIndex        =   0
      Top             =   600
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   6906
      _Version        =   393217
      LabelEdit       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComCtl2.UpDown UpDownItemOP 
      Height          =   315
      Left            =   6180
      TabIndex        =   7
      Top             =   135
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox ItemOP 
      Height          =   315
      Left            =   5655
      TabIndex        =   8
      Top             =   135
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "Item OP:"
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
      Left            =   4860
      TabIndex        =   9
      Top             =   180
      Width           =   765
   End
   Begin VB.Label OP 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2025
      TabIndex        =   6
      Top             =   135
      Width           =   2250
   End
   Begin VB.Label Label1 
      Caption         =   "Ordem de Produção:"
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
      Left            =   210
      TabIndex        =   5
      Top             =   195
      Width           =   1800
   End
End
Attribute VB_Name = "OrdemProducaoRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjOP As New ClassOrdemDeProducao
Dim gobjOPOriginal As ClassOrdemDeProducao
Dim iItemOPAnt As Integer
Dim colComponentes As Collection

Private iIndiceNo As Integer

Public iAlterado As Integer

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object
    Set Form_Load_Ocx = Me
    Caption = "Requisições de Compra para Ordens de Produção"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "OrdemProducaoRC"
End Function

Public Sub Show()
'    Me.Show
'    Parent.SetFocus
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

Private Sub OPs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lErro As Long
Dim objNode As Node
Dim objProdutoKitInfo As ClassProdutoKitInfo

On Error GoTo Erro_Marca_Desmarca

    If iIndiceNo <> 0 Then
    
        Set objNode = OPs.Nodes.Item(iIndiceNo)
    
        For Each objProdutoKitInfo In colComponentes
    
            If objNode.Index = objProdutoKitInfo.iPosicaoArvore Then
    
                If objProdutoKitInfo.objProduto.iCompras = PRODUTO_PRODUZIVEL Or objProdutoKitInfo.objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    objNode.Checked = False
                End If
    
                Exit For
            End If
    
        Next
    
        iIndiceNo = 0
    
    End If

    Exit Sub

Erro_Marca_Desmarca:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181881)

    End Select

    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'''
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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

Public Sub Form_Load()

    giRetornoTela = vbAbort
       
    'Sinaliza que o Form_Loas ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set colComponentes = Nothing
    Set gobjOP = Nothing
    Set gobjOPOriginal = Nothing

End Sub

Function Trata_Parametros(objOP As ClassOrdemDeProducao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjOPOriginal = objOP
    
    OP.Caption = objOP.sCodigo
    
    'Cria cópia com outro endereço de memória para manter dados originais caso aja
    'o cancelamento dos dados alterados
    Call OP_Copia(objOP, gobjOP)
    
    lErro = Traz_OP_Tela(gobjOP)
    If lErro <> SUCESSO Then gError 181856
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 181856
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181857)
    
    End Select
    
    Exit Function
        
End Function

Private Sub BotaoCancela_Click()
    
    'Nao mexer no obj da tela
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 181858
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 181858
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181859)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_OP_Memoria()
    If lErro <> SUCESSO Then gError 181860
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 181860
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181861)

    End Select

    Exit Function

End Function

Function Traz_OP_Tela(objOP As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objProdutoKitInfo As ClassProdutoKitInfo
Dim iSeq As Integer

On Error GoTo Erro_Traz_OP_Tela

    For Each objItemOP In objOP.colItens

        If objItemOP.colProdutoKitInfo.Count = 0 Then
        
            Set objProdutoKitInfo = New ClassProdutoKitInfo
            
            objProdutoKitInfo.sProduto = objItemOP.sProduto
            objProdutoKitInfo.sVersao = objItemOP.sVersao
            objProdutoKitInfo.iNivel = KIT_NIVEL_RAIZ
            objProdutoKitInfo.iSeq = 1
            
            iSeq = 0
        
            lErro = Gera_Dados_ItemOP(objItemOP, objProdutoKitInfo, PRODUTO_PRODUZIVEL, iSeq)
            If lErro <> SUCESSO Then gError 181862
        
        End If
    
    Next
    
    If objOP.colItens.Count <> 0 Then
    
        ItemOP.PromptInclude = False
        ItemOP.Text = "1"
        ItemOP.PromptInclude = True
        
        lErro = Traz_ItemOP_Tela()
        If lErro <> SUCESSO Then gError 181863
        
    End If
    
    Traz_OP_Tela = SUCESSO

    Exit Function

Erro_Traz_OP_Tela:

    Traz_OP_Tela = gErr
    
    Select Case gErr
    
        Case 181860, 181863
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181864)
    
    End Select
    
    Exit Function
    
End Function

Function Traz_ItemOP_Tela() As Long

Dim lErro As Long

On Error GoTo Erro_Traz_ItemOP_Tela

    lErro = Carrega_Arvore(gobjOP.colItens.Item(StrParaInt(ItemOP.Text)))
    If lErro <> SUCESSO Then gError 181865
    
    iItemOPAnt = StrParaInt(ItemOP.Text)

    Traz_ItemOP_Tela = SUCESSO

    Exit Function

Erro_Traz_ItemOP_Tela:

    Traz_ItemOP_Tela = gErr
    
    Select Case gErr
    
        Case 181865
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181866)
    
    End Select
    
    Exit Function
    
End Function

Function Gera_Dados_ItemOP(ByVal objItemOP As ClassItemOP, ByVal objProdutoKitInfoPai As ClassProdutoKitInfo, ByVal iCompras As Integer, iSeq As Integer) As Long

Dim lErro As Long
Dim objKit As New ClassKit
Dim objProduto As ClassProduto
Dim objProdutoKit As ClassProdutoKit
Dim objProdutoKitInfo As ClassProdutoKitInfo
Dim colProdutos As New Collection
Dim bAchou As Boolean

On Error GoTo Erro_Gera_Dados_ItemOP

    If objProdutoKitInfoPai.iNivel >= 20 Then Exit Function

    objKit.sProdutoRaiz = objProdutoKitInfoPai.sProduto
    objKit.sVersao = objProdutoKitInfoPai.sVersao
    
    If objKit.sVersao = "" Then
    
        lErro = CF("Kit_Le_Padrao", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then gError 181867
    
    End If

    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then gError 181868
    
    For Each objProdutoKit In objKit.colComponentes
    
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Or objProdutoKitInfoPai.iNivel = KIT_NIVEL_RAIZ Then
        
            Set objProdutoKitInfo = New ClassProdutoKitInfo
            
            iSeq = iSeq + 1
        
            bAchou = False
            For Each objProduto In colProdutos
                If objProduto.sCodigo = objProdutoKit.sProduto Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                Set objProduto = New ClassProduto
                
                objProduto.sCodigo = objProdutoKit.sProduto
            
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 181869
        
            End If
            
            Set objProdutoKitInfo.objProdutoKit = objProdutoKit
            Set objProdutoKitInfo.objProduto = objProduto
            
            If objProdutoKitInfoPai.iNivel = KIT_NIVEL_RAIZ And objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                objProdutoKitInfo.iSeqPai = 0
            Else
                objProdutoKitInfo.iSeqPai = objProdutoKitInfoPai.iSeq
            End If
            objProdutoKitInfo.iSeq = iSeq
            objProdutoKitInfo.iSeqNivel = objProdutoKit.iSeq
            objProdutoKitInfo.iNivel = objProdutoKit.iNivel + objProdutoKitInfoPai.iNivel
            objProdutoKitInfo.sProduto = objProdutoKit.sProduto
            objProdutoKitInfo.sVersao = objProdutoKit.sVersaoKitComp
            objProdutoKitInfo.sProdutoDesc = objProduto.sDescricao
            
            If iCompras = PRODUTO_PRODUZIVEL And objProduto.iCompras = PRODUTO_COMPRAVEL Then
                objProdutoKitInfo.iSelecionado = MARCADO
            Else
                objProdutoKitInfo.iSelecionado = DESMARCADO
            End If
            
            objItemOP.colProdutoKitInfo.Add objProdutoKitInfo
        
            If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
                lErro = Gera_Dados_ItemOP(objItemOP, objProdutoKitInfo, objProduto.iCompras, iSeq)
                If lErro <> SUCESSO Then gError 181870
            End If
        End If
    
    Next

    Gera_Dados_ItemOP = SUCESSO

    Exit Function

Erro_Gera_Dados_ItemOP:

    Gera_Dados_ItemOP = gErr
    
    Select Case gErr
    
        Case 181867 To 181870
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181871)
    
    End Select
    
    Exit Function
    
End Function

Function Move_OP_Memoria() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemOP As ClassItemOP

On Error GoTo Erro_Move_OP_Memoria

    For Each objItemOP In gobjOPOriginal.colItens
    
        iIndice = iIndice + 1
        
        Set objItemOP.colProdutoKitInfo = gobjOP.colItens.Item(iIndice).colProdutoKitInfo
    
    Next
    
    Move_OP_Memoria = SUCESSO

    Exit Function

Erro_Move_OP_Memoria:

    Move_OP_Memoria = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 181872)
    
    End Select
    
    Exit Function
    
End Function

Function Carrega_Arvore(objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim objNode As Node
Dim objProdutoKitInfo As ClassProdutoKitInfo
Dim objProdutoKitInfoPai As ClassProdutoKitInfo
   
On Error GoTo Erro_Carrega_Arvore

    Set colComponentes = New Collection
    
    OPs.Nodes.Clear

    'Para cada Item nó Pai Insere
    For Each objProdutoKitInfo In objItemOP.colProdutoKitInfo
            
        If objProdutoKitInfo.iNivel = 0 Then
        
            Set objNode = OPs.Nodes.Add(, tvwFirst, "X" & CStr(objProdutoKitInfo.iSeq), "PRODUTO: " & objProdutoKitInfo.sProduto & SEPARADOR & objProdutoKitInfo.sProdutoDesc & " VERSÃO :" & objProdutoKitInfo.sVersao)

            OPs.Nodes.Item(objNode.Index).Expanded = True
            colComponentes.Add objProdutoKitInfo, "X" & objProdutoKitInfo.iSeq
            objNode.Tag = "X" & objProdutoKitInfo.iSeq

            objProdutoKitInfo.iPosicaoArvore = objNode.Index
            
            OPs.Nodes.Item(objNode.Index).Expanded = True

        Else
            
            'Encontra o Pai
            For Each objProdutoKitInfoPai In objItemOP.colProdutoKitInfo
                
                If True And _
                    objProdutoKitInfoPai.iSeq = objProdutoKitInfo.iSeqPai Then
                    Exit For
                End If
            Next

            Set objNode = OPs.Nodes.Add(objProdutoKitInfoPai.iPosicaoArvore, tvwChild, "X" & objProdutoKitInfo.iSeq, "PRODUTO: " & objProdutoKitInfo.sProduto & SEPARADOR & objProdutoKitInfo.sProdutoDesc & " VERSÃO :" & objProdutoKitInfo.sVersao)
            colComponentes.Add objProdutoKitInfo, "X" & objProdutoKitInfo.iSeq
            objNode.Tag = "X" & objProdutoKitInfo.iSeq

            OPs.Nodes.Item(objNode.Index).Expanded = True
            
            If objProdutoKitInfo.iSelecionado = MARCADO Then
                OPs.Nodes.Item(objNode.Index).Checked = True
            Else
                OPs.Nodes.Item(objNode.Index).Checked = False
            End If

            objProdutoKitInfo.iPosicaoArvore = objNode.Index
        
        End If
    
    Next

    Carrega_Arvore = SUCESSO
    
    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181873)

    End Select


    Exit Function
    
End Function

Function OP_Copia(ByVal objOP As ClassOrdemDeProducao, ByVal objOPCopia As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objItemOPCopia As ClassItemOP
Dim objProdutoKitInfo As ClassProdutoKitInfo
Dim objProdutoKitInfoCopia As ClassProdutoKitInfo
   
On Error GoTo Erro_OP_Copia

    For Each objItemOP In objOP.colItens
    
        Set objItemOPCopia = New ClassItemOP
    
        objItemOPCopia.sProduto = objItemOP.sProduto
        objItemOPCopia.sVersao = objItemOP.sVersao
        
        
        For Each objProdutoKitInfo In objItemOP.colProdutoKitInfo
        
            Set objProdutoKitInfoCopia = New ClassProdutoKitInfo
            
            objProdutoKitInfoCopia.iNivel = objProdutoKitInfo.iNivel
            objProdutoKitInfoCopia.iPosicaoArvore = objProdutoKitInfo.iPosicaoArvore
            objProdutoKitInfoCopia.iSelecionado = objProdutoKitInfo.iSelecionado
            objProdutoKitInfoCopia.iSeq = objProdutoKitInfo.iSeq
            objProdutoKitInfoCopia.iSeqNivel = objProdutoKitInfo.iSeqNivel
            objProdutoKitInfoCopia.iSeqPai = objProdutoKitInfo.iSeqPai
            objProdutoKitInfoCopia.sProduto = objProdutoKitInfo.sProduto
            objProdutoKitInfoCopia.sProdutoDesc = objProdutoKitInfo.sProdutoDesc
            objProdutoKitInfoCopia.sVersao = objProdutoKitInfo.sVersao
            
            Set objProdutoKitInfoCopia.objProduto = objProdutoKitInfo.objProduto
            Set objProdutoKitInfoCopia.objProdutoKit = objProdutoKitInfo.objProdutoKit
        
            objItemOPCopia.colProdutoKitInfo.Add objProdutoKitInfoCopia
        
        Next
    
        objOPCopia.colItens.Add objItemOPCopia
    
    Next

    OP_Copia = SUCESSO
    
    Exit Function

Erro_OP_Copia:

    OP_Copia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 181874)

    End Select

    Exit Function
    
End Function

Private Sub ItemOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_ItemOP_Validate

    'Verifica se ItemOP está preenchida
    If Len(Trim(ItemOP.ClipText)) <> 0 Then

        'Critica o ItemOP
        lErro = Inteiro_Critica(ItemOP.Text)
        If lErro <> SUCESSO Then gError 181875
        
        iItem = StrParaInt(ItemOP.Text)
        
        If iItemOPAnt <> iItem Then
        
            'Se o valor estiver fora do range do grid... Erro
            If iItem < 1 Or iItem > gobjOP.colItens.Count Then gError 181876
            
            'Então, mostra a nova arvore
            lErro = Traz_ItemOP_Tela
            If lErro <> SUCESSO Then gError 181877
            
        End If

    End If

    Exit Sub

Erro_ItemOP_Validate:

    Cancel = True

    Select Case gErr

        Case 181875, 181877
            'erros tratados nas rotinas chamadas
            
        Case 181876
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181878)

    End Select

    Exit Sub

End Sub

Private Sub OPs_NodeCheck(ByVal Node As MSComctlLib.Node)
'Marca e desmarca descendentes (Recursivo)

Dim iIndice As Integer
Dim objProdutoKitInfo As ClassProdutoKitInfo

On Error GoTo Erro_OPs_NodeCheck
            
    For Each objProdutoKitInfo In colComponentes
       
        If Node.Index = objProdutoKitInfo.iPosicaoArvore Then
        
            If Node.Checked Then
                If objProdutoKitInfo.objProduto.iCompras = PRODUTO_PRODUZIVEL Then gError 181882
                
                If objProdutoKitInfo.objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 181883
            End If
        
            If Node.Checked Then
                objProdutoKitInfo.iSelecionado = MARCADO
            Else
                objProdutoKitInfo.iSelecionado = DESMARCADO
            End If
        
            Exit For
        End If

    Next
    
    Exit Sub
    
Erro_OPs_NodeCheck:

    iIndiceNo = Node.Index

    Select Case gErr
    
        Case 181882
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, objProdutoKitInfo.objProduto.sCodigo)
                    
        Case 181883
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProdutoKitInfo.objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181884)

    End Select

    Exit Sub

End Sub

Private Sub UpDownItemOP_DownClick()

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_UpDownItemOP_DownClick

    ItemOP.SetFocus

    If Len(Trim(ItemOP.ClipText)) > 0 Then

        iItem = StrParaInt(ItemOP.Text)
        iItem = iItem - 1
        
        If iItem < 1 Then
            iItem = 1
        End If
    
    Else
        iItem = 1
    End If

    ItemOP.PromptInclude = False
    ItemOP.Text = CStr(iItem)
    ItemOP.PromptInclude = True
    
    Call ItemOP_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownItemOP_DownClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181879)

    End Select

    Exit Sub

End Sub

Private Sub UpDownItemOP_UpClick()

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_UpDownItemOP_UpClick

    ItemOP.SetFocus

    If Len(ItemOP.ClipText) > 0 Then

        iItem = StrParaInt(ItemOP.Text)
        iItem = iItem + 1
        
        If iItem > gobjOP.colItens.Count Then
            iItem = gobjOP.colItens.Count
        End If
        
    Else
        iItem = 1
    End If

    ItemOP.PromptInclude = False
    ItemOP.Text = CStr(iItem)
    ItemOP.PromptInclude = True
    
    Call ItemOP_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownItemOP_UpClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181880)

    End Select

    Exit Sub

End Sub

Private Sub Marca_Desmarca(ByVal bFlag As Boolean)

Dim lErro As Long
Dim objNode As Node
Dim objProdutoKitInfo As ClassProdutoKitInfo

On Error GoTo Erro_Marca_Desmarca

    For Each objNode In OPs.Nodes
    
        objNode.Checked = bFlag

        For Each objProdutoKitInfo In colComponentes
       
            If objNode.Index = objProdutoKitInfo.iPosicaoArvore Then
                
                If objProdutoKitInfo.objProduto.iCompras = PRODUTO_PRODUZIVEL Or objProdutoKitInfo.objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    objNode.Checked = False
                End If
            
                If objNode.Checked Then
                    objProdutoKitInfo.iSelecionado = MARCADO
                Else
                    objProdutoKitInfo.iSelecionado = DESMARCADO
                End If
            
                Exit For
            End If
    
        Next
        
    Next
    
    Exit Sub

Erro_Marca_Desmarca:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181881)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click(Index As Integer)
    Call Marca_Desmarca(False)
End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Marca_Desmarca(True)
End Sub

