VERSION 5.00
Begin VB.UserControl SubstProdutoNFOcx 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   6735
   Begin VB.Frame Frame1 
      Caption         =   "Substituir"
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   735
      Width           =   6375
      Begin VB.Frame Frame3 
         Caption         =   "Quantidade Disponível"
         Height          =   1065
         Left            =   2385
         TabIndex        =   14
         Top             =   645
         Width           =   3825
         Begin VB.Label QuantDisponivelPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   18
            Top             =   255
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado Padrão:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label QuantDisponivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   16
            Top             =   645
            Width           =   1440
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1590
            TabIndex        =   15
            Top             =   675
            Width           =   510
         End
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   2970
         TabIndex        =   23
         Top             =   270
         Width           =   3225
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   990
         TabIndex        =   22
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Unidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   780
         Width           =   780
      End
      Begin VB.Label UnidadeMedida 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   990
         TabIndex        =   19
         Top             =   735
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Por"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   2820
      Width           =   6375
      Begin VB.Frame Frame4 
         Caption         =   "Quantidade Disponível"
         Height          =   1065
         Left            =   2370
         TabIndex        =   4
         Top             =   660
         Width           =   3795
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1590
            TabIndex        =   8
            Top             =   675
            Width           =   510
         End
         Begin VB.Label QuantDisponivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   7
            Top             =   645
            Width           =   1440
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado Padrão:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   6
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label QuantDisponivelPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   5
            Top             =   255
            Width           =   1440
         End
      End
      Begin VB.ComboBox ProdutoSubstituto 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label UnidadeMedida 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Unidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   2925
         TabIndex        =   9
         Top             =   270
         Width           =   3225
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   4680
      Picture         =   "SubstProdutoNFOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   5670
      Picture         =   "SubstProdutoNFOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "SubstProdutoNFOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Declaração de variáveis globais

Dim gobjSubstProdutoNF As ClassSubstProdutoNF
Dim gsProdutoSubst(0 To 1) As String
Dim giCodigoTipoDocInfo As Integer

Private Function Verifica_Substituto(sSubstituto1 As String, sSubstituto2 As String)
'Recebe os códigos dos produtos substitutos
'Verifica se eles são válidos para operação de substituição
'e os coloca na combo

Dim lErro As Long
Dim iIndice As Integer
Dim iVerificaProd1 As Integer
Dim iVerificaProd2 As Integer
Dim sProdutoMascarado As String
Dim objProdutoSubstituto1 As New ClassProduto
Dim objProdutoSubstituto2 As New ClassProduto

On Error GoTo Erro_Verifica_Substituto
    
    'Se os códigos Substitutos estiverem vazios --> Erro
    If Len(Trim(sSubstituto1)) = 0 And Len(Trim(sSubstituto2)) = 0 Then Error 23783
    
    'Lê no BD os produtos substitutos
    If Len(Trim(sSubstituto1)) > 0 Then
        objProdutoSubstituto1.sCodigo = sSubstituto1
        
        lErro = CF("Produto_Le",objProdutoSubstituto1)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 23817
            
    End If

    If Len(Trim(sSubstituto2)) > 0 Then
        objProdutoSubstituto2.sCodigo = sSubstituto2
        
        lErro = CF("Produto_Le",objProdutoSubstituto2)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 23818
            
    End If
    
    'Verificam se os produtos substitutos já se encontram no -->
    '--> NF de venda, se estão inativos ou não participam do faturamento
    iVerificaProd1 = 0
    iVerificaProd2 = 0

    'Verifica se os produtos substitutos já estão na coleção ou estão inativos ou tem faturamento igual a zero
    For iIndice = 1 To gobjSubstProdutoNF.colOutrosProdutosNF.Count
        If objProdutoSubstituto1.sCodigo = gobjSubstProdutoNF.colOutrosProdutosNF.Item(iIndice) Or objProdutoSubstituto1.iAtivo = Inativo Or objProdutoSubstituto1.iFaturamento = 0 Then
            iVerificaProd1 = 1
        End If

        If objProdutoSubstituto2.sCodigo = gobjSubstProdutoNF.colOutrosProdutosNF.Item(iIndice) Or objProdutoSubstituto2.iAtivo = Inativo Or objProdutoSubstituto2.iFaturamento = 0 Then
            iVerificaProd2 = 1
        End If

        If iVerificaProd1 = 1 And iVerificaProd2 = 1 Then Error 23784

    Next
    
    'Preenche combo ProdutoSustituto com os Códigos
    'O produto não deve pertencer ao NF de venda e estar preenchido
    If iVerificaProd1 = 0 And Len(Trim(sSubstituto1)) > 0 Then
        
        lErro = Mascara_MascararProduto(sSubstituto1, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 23803
                
        ProdutoSubstituto.AddItem sProdutoMascarado
        gsProdutoSubst(0) = sSubstituto1
        ProdutoSubstituto.ItemData(0) = 0
        
    End If
    
    If iVerificaProd2 = 0 And Len(Trim(sSubstituto2)) > 0 Then
            
        lErro = Mascara_MascararProduto(sSubstituto2, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 23804
        
        ProdutoSubstituto.AddItem sProdutoMascarado
        gsProdutoSubst(ProdutoSubstituto.NewIndex) = sSubstituto2
        ProdutoSubstituto.ItemData(ProdutoSubstituto.NewIndex) = ProdutoSubstituto.NewIndex
        
    End If

    Verifica_Substituto = SUCESSO
    
    Exit Function
    
Erro_Verifica_Substituto:
    
    Verifica_Substituto = Err
    
    Select Case Err
    
        Case 23783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_SUBSTITUTOS", Err)

        Case 23784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOS_SUBSTITUTOS_INVALIDOS1", Err)
                
        Case 23803, 23804, 23817, 23818
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174454)
            
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoCancela_Click()

    giRetornoTela = vbCancel
    Unload Me

End Sub

Private Sub BotaoOK_Click()
    
    gobjSubstProdutoNF.sProdutoSubstituto = gsProdutoSubst(ProdutoSubstituto.ItemData(ProdutoSubstituto.ListIndex))
    giRetornoTela = vbOK
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giRetornoTela = vbCancel

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174455)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjSubstProdutoNF = Nothing

End Sub

Function Trata_Parametros(objSubstProduto As ClassSubstProdutoNF, iCodigo As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAlmoxarifadoPadrao As Integer
Dim colEstoque As New colEstoqueProduto
Dim objItem As New ClassItemNF
Dim objProduto As New ClassProduto

On Error GoTo Erro_Trata_Parametros

    Set gobjSubstProdutoNF = objSubstProduto
    giCodigoTipoDocInfo = iCodigo

    objProduto.sCodigo = gobjSubstProdutoNF.sProduto

    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 23782

    If lErro = 28030 Then Error 23808
            
    'Verificar se os produtos substitutos são válidos
    lErro = Verifica_Substituto(objProduto.sSubstituto1, objProduto.sSubstituto2)
    If lErro <> SUCESSO Then Error 23816
        
    'Lê as posições de Estoque do produto do Ítem
    lErro = CF("EstoquesProduto_Le_Filial",gobjSubstProdutoNF.sProduto, colEstoque)
    If lErro <> SUCESSO Then Error 23785

    'Lê almoxarifado padrão desta FilialEmpresa
    lErro = CF("AlmoxarifadoPadrao_Le",giFilialEmpresa, gobjSubstProdutoNF.sProduto, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then Error 23786
    
    'Não existe almoxarifado padrão para este produto
    If lErro = 23796 Then iAlmoxarifadoPadrao = 0
    
    lErro = Preenche_Tela(objProduto, colEstoque, iAlmoxarifadoPadrao, 0)
    If lErro <> SUCESSO Then Error 23788

    'Mostra o primeiro produto substituto buscado no BD
    ProdutoSubstituto.ListIndex = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    giRetornoTela = vbCancel

    Select Case Err

        Case 23782, 23785, 23786, 23788

        Case 23808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
        
        Case 23816
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174456)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If UnloadMode <> vbFormCode Then giRetornoTela = vbCancel

End Sub

Private Sub ProdutoSubstituto_Click()

Dim lErro As Long
Dim iAlmoxarifadoPadrao As Integer
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colEstoque As New colEstoqueProduto

On Error GoTo Erro_ProdutoSubstituto_Click

    'Se nenhum produto foi selecionado --> sai
    If ProdutoSubstituto.ListIndex = -1 Then Exit Sub

    'Lê produto no BD a partir do seu código
    objProduto.sCodigo = gsProdutoSubst(ProdutoSubstituto.ItemData(ProdutoSubstituto.ListIndex))

    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 23789
    
    If lErro = 28030 Then Error 23806
    
    'Lê as posições de Estoque do produto
    lErro = CF("EstoquesProduto_Le_Filial",objProduto.sCodigo, colEstoque)
    If lErro <> SUCESSO Then Error 23790

    'Lê Almoxarifado Padrão do produto desta FilialEmpresa
    lErro = CF("AlmoxarifadoPadrao_Le",giFilialEmpresa, objProduto.sCodigo, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then Error 23791
    
    'Não existe almoxarifado padrão para este produto
    If lErro = 23796 Then iAlmoxarifadoPadrao = 0
        
    'Preenche os dados do produto substituto selecionado
    lErro = Preenche_Tela(objProduto, colEstoque, iAlmoxarifadoPadrao, 1)
    If lErro <> SUCESSO Then Error 23792

    Exit Sub

Erro_ProdutoSubstituto_Click:

    Select Case Err

        Case 23789, 23790, 23791, 23792
        
        Case 23806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174457)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Tela(objProduto As ClassProduto, colEstoque As colEstoqueProduto, iAlmoxarifado As Integer, iIndice As Integer) As Long

Dim lErro As Long
Dim sProdutoMascarado As String
Dim dTotalDisponivel As Double
Dim iIndice1 As Integer

On Error GoTo Erro_Preenche_Tela

    If iIndice = 0 Then
        lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 43828
                
        Produto.Caption = sProdutoMascarado
    End If
    
    Descricao(iIndice).Caption = objProduto.sDescricao
    UnidadeMedida(iIndice).Caption = objProduto.sSiglaUMEstoque
    
    'Verifica em colEstoque qual almoxarifado corresponde ao almoxarifado padrão
    dTotalDisponivel = 0
    QuantDisponivelPadrao(iIndice).Caption = Formata_Estoque(0)
    
    For iIndice1 = 1 To colEstoque.Count
        If colEstoque(iIndice1).iAlmoxarifado = iAlmoxarifado Then
            
            If giCodigoTipoDocInfo = DOCINFO_NFISPC Or giCodigoTipoDocInfo = DOCINFO_NFFISPC Then
                'Preenche quantidade consignada padrão
                QuantDisponivelPadrao(iIndice).Caption = Formata_Estoque(colEstoque(iIndice1).dQuantConsig)
                        
            ElseIf giCodigoTipoDocInfo = DOCINFO_NFISRMB3PV Then 'Incluido por Leo em 15/01/02
                'Preenche com a quantidade benef. de terceiros
                QuantDisponivelPadrao(iIndice).Caption = Formata_Estoque(colEstoque(iIndice1).dQuantBenef3)
                   
            Else
                'Preenche quantidade disponível padrão
                QuantDisponivelPadrao(iIndice).Caption = Formata_Estoque(colEstoque(iIndice1).dQuantDisponivel)
            End If
            
        End If

        If giCodigoTipoDocInfo = DOCINFO_NFISPC Or giCodigoTipoDocInfo = DOCINFO_NFFISPC Then
            'Acumula o total consignado nos almoxarifados em dTotalDisponivel
            dTotalDisponivel = dTotalDisponivel + colEstoque(iIndice1).dQuantConsig
        
        ElseIf giCodigoTipoDocInfo = DOCINFO_NFISRMB3PV Then 'por Leo em 15/01/02
            'Preenche com a quantidade benef. de terceiros
            dTotalDisponivel = dTotalDisponivel + colEstoque(iIndice1).dQuantBenef3
            
        Else
            'Acumula o total disponível nos almoxarifados em dTotalDisponivel
            dTotalDisponivel = dTotalDisponivel + colEstoque(iIndice1).dQuantDisponivel
        
        End If
    
    Next

    'Mostra o total Disponível
    For iIndice = 0 To 1
        QuantDisponivel(iIndice).Caption = Formata_Estoque(dTotalDisponivel)
    Next
    
    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = Err

    Select Case Err
    
        Case 43828

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174458)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_SUBST_PRODUTO_NF
    Set Form_Load_Ocx = Me
    Caption = "Substituição de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "SubstProdutoNF"
    
End Function

Public Sub Show()
'    Parent.Show
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




Private Sub QuantDisponivelPadrao_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(QuantDisponivelPadrao(Index), Source, X, Y)
End Sub

Private Sub QuantDisponivelPadrao_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivelPadrao(Index), Button, Shift, X, Y)
End Sub

Private Sub QuantDisponivel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(QuantDisponivel(Index), Source, X, Y)
End Sub

Private Sub QuantDisponivel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivel(Index), Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Descricao(Index), Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao(Index), Button, Shift, X, Y)
End Sub

Private Sub UnidadeMedida_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(UnidadeMedida(Index), Source, X, Y)
End Sub

Private Sub UnidadeMedida_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida(Index), Button, Shift, X, Y)
End Sub


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Produto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Produto, Source, X, Y)
End Sub

Private Sub Produto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Produto, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

