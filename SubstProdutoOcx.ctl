VERSION 5.00
Begin VB.UserControl SubstProdutoOcx 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   6645
   Begin VB.Frame Frame1 
      Caption         =   "Substituir"
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   705
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
         Left            =   195
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
         Left            =   1005
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
      Top             =   2790
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
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
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   5640
      Picture         =   "SubstProdutoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   4665
      Picture         =   "SubstProdutoOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "SubstProdutoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Declaração de variáveis globais
Dim gobjSubstProduto As ClassSubstProduto
Dim gsProdutoSubst(0 To 1) As String

Private Function Verifica_Substituto(sSubstituto1 As String, sSubstituto2 As String)
'Recebe os códigos dos produtos substitutos
'Verifica se eles são válidos para operação de substituição
'e os inclui na combo

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
    '--> pedido de venda, se estão inativos ou não participam do faturamento
    iVerificaProd1 = 0
    iVerificaProd2 = 0

    If objProdutoSubstituto1.iAtivo = Inativo Or objProdutoSubstituto1.iFaturamento = 0 Then
        iVerificaProd1 = 1
    End If
    
    If objProdutoSubstituto2.iAtivo = Inativo Or objProdutoSubstituto2.iFaturamento = 0 Then
        iVerificaProd2 = 1
    End If
    
    'Verifica se os produtos substitutos já estão na coleção ou estão inativos ou tem faturamento igual a zero
    For iIndice = 1 To gobjSubstProduto.ColItemPedido.Count
        
        If objProdutoSubstituto1.sCodigo = gobjSubstProduto.ColItemPedido(iIndice).sProduto Then
            iVerificaProd1 = 1
        End If

        If objProdutoSubstituto2.sCodigo = gobjSubstProduto.ColItemPedido(iIndice).sProduto Then
            iVerificaProd2 = 1
        End If

        If iVerificaProd1 = 1 And iVerificaProd2 = 1 Then Error 23784

    Next
    
    'Preenche combo ProdutoSustituto com os Códigos
    'O produto não deve pertencer ao pedido de venda e estar preenchido
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOS_SUBSTITUTOS_INVALIDOS", Err)
                
        Case 23803, 23804, 23817, 23818
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174459)
            
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoCancela_Click()

    giRetornoTela = vbCancel
    Unload Me

End Sub

Private Sub BotaoOK_Click()
    
    gobjSubstProduto.sCodProdutoSubstituto = gsProdutoSubst(ProdutoSubstituto.ItemData(ProdutoSubstituto.ListIndex))
    giRetornoTela = vbOK
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174460)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objSubstProduto As ClassSubstProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAlmoxarifadoPadrao As Integer
Dim colEstoque As New colEstoqueProduto
Dim colReservaItemBD As New colReservaItem
Dim objItemPedido As New ClassItemPedido
Dim objProduto As New ClassProduto

On Error GoTo Erro_Trata_Parametros

    Set gobjSubstProduto = objSubstProduto

    objProduto.sCodigo = gobjSubstProduto.sCodProduto

    lErro = CF("Produto_Le",objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 23782

    If lErro = 28030 Then gError 23808
            
    'Verificar se os produtos substitutos são válidos
    lErro = Verifica_Substituto(objProduto.sSubstituto1, objProduto.sSubstituto2)
    If lErro <> SUCESSO Then gError 23816
        
    'Lê as posições de Estoque do produto do Ítem
    lErro = CF("EstoquesProduto_Le",gobjSubstProduto.sCodProduto, colEstoque)
    If lErro <> SUCESSO And lErro <> 30100 Then gError 23785

    'Lê almoxarifado padrão desta FilialEmpresa
    lErro = CF("AlmoxarifadoPadrao_Le",giFilialEmpresa, gobjSubstProduto.sCodProduto, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 23786
    
    'Não existe almoxarifado padrão para este produto
    If lErro = 23796 Then iAlmoxarifadoPadrao = 0
    
    'Lê as reservas do Ítem
    For iIndice = 1 To gobjSubstProduto.ColItemPedido.Count
    
        If gobjSubstProduto.sCodProduto = gobjSubstProduto.ColItemPedido(iIndice).sProduto Then
            objItemPedido.sProduto = gobjSubstProduto.ColItemPedido(iIndice).sProduto
            objItemPedido.lCodPedido = gobjSubstProduto.ColItemPedido(iIndice).lCodPedido

            lErro = CF("ReservasItem_Le",objItemPedido, colReservaItemBD)
            If lErro <> SUCESSO And lErro <> 30099 Then gError 23787

            Exit For

        End If

    Next

    lErro = Preenche_Tela(objProduto, colEstoque, colReservaItemBD, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO Then gError 23788

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 23782, 23785, 23786, 23787, 23788

        Case 23808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 23816
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174461)

    End Select

    Exit Function

End Function

Private Function Preenche_Tela(objProduto As ClassProduto, colEstoque As colEstoqueProduto, colReservaItemBD As colReservaItem, iAlmoxarifado As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objEstoqueProduto As ClassEstoqueProduto
Dim dQuantReservadaPedido As Double
Dim objReservaItem As ClassReservaItem
Dim dTotalDisponivel As Double
Dim sProdutoEnxuto As String

On Error GoTo Erro_Preenche_Tela

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then Error 23805
    
    Produto.Caption = sProdutoEnxuto
    Descricao(0).Caption = objProduto.sDescricao
    UnidadeMedida(0).Caption = objProduto.sSiglaUMEstoque

    'Localiza em colReservaItemBD a quantidade reservada no almoxarifado padrão se houver
    For iIndice = 1 To colReservaItemBD.Count
        If colReservaItemBD(iIndice).iAlmoxarifado = iAlmoxarifado Then
            dQuantReservadaPedido = colReservaItemBD(iIndice).dQuantidade
            Exit For
        End If
    Next

    'Localiza em colEstoque o Almoxarifado que corresponde a iAlmoxarifado
    QuantDisponivelPadrao(0).Caption = Formata_Estoque(0)
    For iIndice = 1 To colEstoque.Count
        If colEstoque(iIndice).iAlmoxarifado = iAlmoxarifado Then
            'Preenche a quantidade disponível padrão
            QuantDisponivelPadrao(0).Caption = Formata_Estoque(colEstoque(iIndice).dQuantDisponivel + dQuantReservadaPedido)

            Exit For

        End If

    Next

    'Buscar o total disponivel nos Almoxarifados
    'Para cada Almoxarifado de colEstoque -->
    dTotalDisponivel = 0
    For Each objEstoqueProduto In colEstoque

        dQuantReservadaPedido = 0
        '--> Verifica se existe quantidade reservada no BD
        For Each objReservaItem In colReservaItemBD
            If objEstoqueProduto.iAlmoxarifado = objReservaItem.iAlmoxarifado Then
                dQuantReservadaPedido = dQuantReservadaPedido + objReservaItem.dQuantidade

            End If
        Next

        'Acumula o Total Disponível em dTotalDisponivel
        dTotalDisponivel = dTotalDisponivel + objEstoqueProduto.dQuantDisponivel + dQuantReservadaPedido

    Next

    'Preenche a Quantidade total disponível
    QuantDisponivel(0).Caption = Formata_Estoque(dTotalDisponivel)

    'Mostra o primeiro produto substituto buscado no BD
    If ProdutoSubstituto.ListCount > 0 Then ProdutoSubstituto.ListIndex = 0

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = Err

    Select Case Err
        
        Case 23805
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174462)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If UnloadMode <> vbFormCode Then giRetornoTela = vbCancel

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjSubstProduto = Nothing

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
    
    Descricao(1).Caption = objProduto.sDescricao
    QuantDisponivel(1).Caption = Formata_Estoque(0)
    QuantDisponivelPadrao(1) = Formata_Estoque(0)
    
    'Lê as posições de Estoque do produto
    lErro = CF("EstoquesProduto_Le",objProduto.sCodigo, colEstoque)
    If lErro <> SUCESSO And lErro <> 30100 Then Error 23790
    
    'Lê Almoxarifado Padrão do produto desta FilialEmpresa
    lErro = CF("AlmoxarifadoPadrao_Le",giFilialEmpresa, objProduto.sCodigo, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then Error 23791
    
    'Não existe almoxarifado padrão para este produto
    If lErro = 23796 Then iAlmoxarifadoPadrao = 0

    'Preenche os dados do produto substituto selecionado
    lErro = Preenche_Tela2(objProduto, colEstoque, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO Then Error 23792

    Exit Sub

Erro_ProdutoSubstituto_Click:

    Select Case Err

        Case 23789, 23790, 23791, 23792
        
        Case 23806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174463)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Tela2(objProduto As ClassProduto, colEstoque As colEstoqueProduto, iAlmoxarifado As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dTotalDisponivel As Double

On Error GoTo Erro_Preenche_Tela2

    Descricao(1).Caption = objProduto.sDescricao
    UnidadeMedida(1).Caption = objProduto.sSiglaUMEstoque
    
    'Verifica em colEstoque qual almoxarifado corresponde ao almoxarifado padrão
    dTotalDisponivel = 0
    QuantDisponivelPadrao(1).Caption = Formata_Estoque(0)
    For iIndice = 1 To colEstoque.Count
        If colEstoque(iIndice).iAlmoxarifado = iAlmoxarifado Then
            'Preenche quantidade disponível padrão
            QuantDisponivelPadrao(1).Caption = Formata_Estoque(colEstoque(iIndice).dQuantDisponivel)
            
        End If

        'Acumula o total disponível nos almoxarifados em dTotalDisponivel
        dTotalDisponivel = dTotalDisponivel + colEstoque(iIndice).dQuantDisponivel

    Next

    'Mostra o total Disponível
    QuantDisponivel(1).Caption = Formata_Estoque(dTotalDisponivel)

    Preenche_Tela2 = SUCESSO

    Exit Function

Erro_Preenche_Tela2:

    Preenche_Tela2 = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174464)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_SUBSTITUICAO_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Substituição de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "SubstProduto"
    
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

