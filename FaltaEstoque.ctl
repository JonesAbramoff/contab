VERSION 5.00
Begin VB.UserControl FaltaEstoque 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   ScaleHeight     =   4950
   ScaleWidth      =   6375
   Begin VB.CommandButton BotaoParaTodos 
      Caption         =   "Aplicar em todos os itens"
      Height          =   525
      Left            =   180
      Picture         =   "FaltaEstoque.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4260
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.ListBox Tratamento 
      Height          =   1230
      ItemData        =   "FaltaEstoque.ctx":015A
      Left            =   165
      List            =   "FaltaEstoque.ctx":0170
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2850
      Width           =   6060
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2265
      Picture         =   "FaltaEstoque.ctx":0216
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4275
      Width           =   1005
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   3420
      Picture         =   "FaltaEstoque.ctx":0370
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4275
      Width           =   990
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1380
      TabIndex        =   17
      Top             =   660
      Width           =   4830
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1380
      TabIndex        =   16
      Top             =   165
      Width           =   1410
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
      Left            =   570
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   195
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quant. a Reservar:"
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
      Left            =   3015
      TabIndex        =   14
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Label QuantReservar 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4740
      TabIndex        =   13
      Top             =   1170
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quant. Disponível:"
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
      Left            =   3045
      TabIndex        =   12
      Top             =   1710
      Width           =   1620
   End
   Begin VB.Label QuantDisponivel 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4740
      TabIndex        =   11
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "U.M.:"
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
      Left            =   825
      TabIndex        =   10
      Top             =   1215
      Width           =   480
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1380
      TabIndex        =   9
      Top             =   1170
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tratamento"
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
      Left            =   150
      TabIndex        =   8
      Top             =   2610
      Width           =   1005
   End
   Begin VB.Label Label6 
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
      Height          =   195
      Left            =   375
      TabIndex        =   7
      Top             =   690
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifado:"
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
      Left            =   150
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Almoxarifado 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1380
      TabIndex        =   5
      Top             =   1650
      Width           =   1410
   End
   Begin VB.Label QuantDisponivelTodos 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4740
      TabIndex        =   4
      Top             =   2205
      Width           =   1440
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade Disponível em Todos os Almoxarifados:"
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
      Left            =   270
      TabIndex        =   3
      Top             =   2250
      Width           =   4395
   End
End
Attribute VB_Name = "FaltaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'responsavel: Jones
'revisada em:05/11/98
'pendencias:

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Declaração das variáveis globais
Dim gobjItemPedido As ClassItemPedido
Dim gColItemPedido As colItemPedido

Private Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Private Sub BotaoParaTodos_Click()
    If Tratamento.ListIndex <> -1 Then
        gobjItemPedido.iTratamentoFaltaEstoque = Tratamento.ItemData(Tratamento.ListIndex)
    End If
    Call BotaoOK_Click
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim dFator As Double
Dim dQuantReservar As Double
Dim dQuantDisponivel As Double
Dim objSubstProduto As New ClassSubstProduto
Dim objItemOP As New ClassItemOP
Dim dQuantCancelada As Double 'Necessária para nao subtrair a quant cancelada q já foi subtraida

On Error GoTo Erro_BotaoOK_Click

    If Tratamento.ListIndex = -1 Then Error 30095
    
    Select Case Tratamento.ItemData(Tratamento.ListIndex)
    
        Case TR_CANCELA
            If Len(Trim(QuantReservar.Caption)) <> 0 Then dQuantReservar = CDbl(QuantReservar.Caption)
            
            lErro = CF("UM_Conversao_Trans", gobjItemPedido.iClasseUM, gobjItemPedido.sUMEstoque, gobjItemPedido.sUnidadeMed, dFator)
            If lErro <> SUCESSO Then Error 30096
            
            gobjItemPedido.dQuantCancelada = gobjItemPedido.dQuantCancelada + (dQuantReservar) * dFator
            
            giRetornoTela = vbOK
            Unload Me
            
        Case TR_CANCELA_ULTR_DISP
            If Len(Trim(QuantReservar.Caption)) <> 0 Then dQuantReservar = CDbl(QuantReservar.Caption)
            If Len(Trim(QuantDisponivel.Caption)) <> 0 Then dQuantDisponivel = CDbl(QuantDisponivel.Caption)
            
            If dQuantDisponivel > QTDE_ESTOQUE_DELTA Then
               
               lErro = CF("UM_Conversao_Trans", gobjItemPedido.iClasseUM, gobjItemPedido.sUMEstoque, gobjItemPedido.sUnidadeMed, dFator)
               If lErro <> SUCESSO Then Error 30097
        
               dQuantCancelada = (dQuantReservar - dQuantDisponivel) * dFator
                                      
               gobjItemPedido.dQuantReservada = StrParaDbl(Formata_Estoque((dQuantReservar * dFator) - dQuantCancelada))
                                      
               gobjItemPedido.dQuantCancelada = gobjItemPedido.dQuantCancelada + dQuantCancelada
               
               gobjItemPedido.ColReserva.Add 0, 0, "", Almoxarifado.Tag, 0, 0, 0, dQuantDisponivel, DATA_NULA, DATA_NULA, "", RESERVA_AUTO_RESP, 0, Almoxarifado.Caption
            
            End If
            
            giRetornoTela = vbOK
            
            Unload Me
            
        Case TR_ALOC_MANUAL
            Call Chama_Tela_Modal("AlocacaoProduto1", gobjItemPedido)
            
            If giRetornoTela = vbOK Then Unload Me
            
        Case TR_NAO_RESERVA
            giRetornoTela = vbOK
            
            Unload Me
            
        Case TR_RESERVA_EST
        
            dQuantDisponivel = StrParaDbl(QuantDisponivel.Caption)
            
            'If dQuantDisponivel = 0 Then Error 30099
            
            If dQuantDisponivel > QTDE_ESTOQUE_DELTA Then
            
                gobjItemPedido.ColReserva.Add 0, 0, "", Almoxarifado.Tag, 0, 0, 0, dQuantDisponivel, DATA_NULA, DATA_NULA, "", RESERVA_AUTO_RESP, 0, Almoxarifado.Caption
            
                lErro = CF("UM_Conversao", gobjItemPedido.iClasseUM, gobjItemPedido.sUMEstoque, gobjItemPedido.sUnidadeMed, dFator)
                If lErro <> SUCESSO Then Error 51465
                
                '????
                gobjItemPedido.dQuantReservada = StrParaDbl(Formata_Estoque(dQuantDisponivel * dFator))
                
            End If
            
            giRetornoTela = vbOK
            
            Unload Me
            
        Case TR_SUBSTITUI
            'Testa se já foi faturado (parcialmente). Se foi, erro.
            If gobjItemPedido.dQuantFaturada > 0 Then Error 30098
            
            'Testa se está vinculado a OP. Se estiver erro.
            If gobjItemPedido.lNumIntDoc <> 0 Then
            
                lErro = CF("ItemOP_Le_ItemPV", objItemOP, gobjItemPedido)
                If lErro <> SUCESSO And lErro <> 46074 Then Error 24459
                If lErro = SUCESSO Then Error 24460
            
            End If
            
            objSubstProduto.sCodProduto = gobjItemPedido.sProduto
            Set objSubstProduto.colItemPedido = gColItemPedido
            
            Call Chama_Tela_Modal("SubstProduto", objSubstProduto)
            
            If giRetornoTela = vbOK Then
                gobjItemPedido.sProduto = objSubstProduto.sCodProdutoSubstituto
                
                Unload Me
                
            End If
            
    End Select
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case Err
    
        Case 24459, 30096, 30097, 51465
        
        Case 24460
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_VINCULADO_ITEM_OP", Err, gobjItemPedido.iItem, objItemOP.lNumIntDoc)
             
        Case 30095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRATAMENTO_NAO_INFORMADO", Err)
        
        Case 30098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_FATURADA_MAIORZERO", Err, gobjItemPedido.dQuantFaturada)
        
        Case 30099 'ERRO_SALDO_MAT_DISPONIVEL
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DISPONIVEL", Err, Produto.Caption, Almoxarifado.Caption, QuantDisponivel.Caption)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159900)
            
    End Select
    
    Exit Sub
        
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giRetornoTela = vbCancel
    
    lErro = CF("PV_FaltaEstoque_Preenche_Trat_Cust", Tratamento)
    If lErro <> SUCESSO Then gError 198690
        
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 198690
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 159901)
            
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros(Optional objItemPedido As ClassItemPedido, Optional colItemPedido As colItemPedido, Optional dQuantReservarEstoque As Double, Optional objAlmoxarifadoPadrao As ClassAlmoxarifado, Optional dSaldoAlmoxarifadoPadrao As Double) As Long

Dim lErro As Long
Dim dQuantReservadaPedido As Double
Dim dSaldoTodosAlmoxarifados As Double
Dim colEstoque As New ColEstoqueProduto
Dim colReservaItemBD As New colReservaItem
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_Trata_Parametros

    Set gobjItemPedido = objItemPedido
    Set gColItemPedido = colItemPedido
    
    'Quantidade disponivel do ítem em todos os Almoxarifados
    lErro = CF("EstoquesProduto_Le", gobjItemPedido.sProduto, colEstoque)
    If lErro <> SUCESSO And lErro <> 30100 Then Error 30092
    
    'Lê reservas do ítem no BD
    lErro = CF("ReservasItem_Le", gobjItemPedido, colReservaItemBD)
    If lErro <> SUCESSO And lErro <> 30099 Then Error 30093
    
    'Em cada Almoxarifado verifica se existe Reserva do ítem correspondente
    For Each objEstoqueProduto In colEstoque
        dQuantReservadaPedido = 0
        For Each objReservaItem In colReservaItemBD
            If objEstoqueProduto.iAlmoxarifado = objReservaItem.iAlmoxarifado Then
                dQuantReservadaPedido = objReservaItem.dQuantidade
                Exit For
            End If
        Next
        'Adiciona a quantidade reservada do ítem à quantidade disponível
        objEstoqueProduto.dSaldo = objEstoqueProduto.dQuantDisponivel + dQuantReservadaPedido
    Next
        
    'Calcula a soma das disponibilidades de todos os almoxarifados
    dSaldoTodosAlmoxarifados = 0
    For Each objEstoqueProduto In colEstoque
        dSaldoTodosAlmoxarifados = dSaldoTodosAlmoxarifados + objEstoqueProduto.dSaldo
    Next
    
    'Preenchimento de Tela
    lErro = Preenche_Tela(gobjItemPedido, dQuantReservarEstoque, objAlmoxarifadoPadrao, dSaldoAlmoxarifadoPadrao, dSaldoTodosAlmoxarifados)
    If lErro <> SUCESSO Then Error 30094
    
    'Se o produto for de grade -> Remove a opção de substituição
    If objItemPedido.iPossuiGrade = MARCADO Then Tratamento.RemoveItem (Tratamento.ListCount - 1)
    
    If objItemPedido.bTelaFaltaEstExibeBtnAplicarEmTodos Then BotaoParaTodos.Visible = True
    
    If gobjItemPedido.iTratamentoFaltaEstoque <> 0 Then
        Call Combo_Seleciona_ItemData(Tratamento, gobjItemPedido.iTratamentoFaltaEstoque)
        Call BotaoOK_Click
    End If
        
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case 30092, 30093, 30094
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159902)
            
    End Select
    
    Exit Function
        
End Function

Public Sub Form_Activate()

    If gobjItemPedido.iTratamentoFaltaEstoque <> 0 Then
        Call Combo_Seleciona_ItemData(Tratamento, gobjItemPedido.iTratamentoFaltaEstoque)
        Call BotaoOK_Click
    End If

End Sub

Function Preenche_Tela(objItemPedido As ClassItemPedido, dQuantReservarEstoque As Double, objAlmoxarifado As ClassAlmoxarifado, dSaldoAlmoxarifadoPadrao As Double, dSaldoTodosAlmoxarifados As Double) As Long

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Preenche_Tela

    lErro = Mascara_MascararProduto(objItemPedido.sProduto, sProduto)
    If lErro <> SUCESSO Then Error 23740
    
    Produto.Caption = sProduto
    Descricao.Caption = objItemPedido.sProdutoDescricao
    UnidadeMedida.Caption = objItemPedido.sUMEstoque
    
    QuantReservar.Caption = Formata_Estoque(dQuantReservarEstoque)
    
    Almoxarifado.Caption = objAlmoxarifado.sNomeReduzido
    Almoxarifado.Tag = objAlmoxarifado.iCodigo
    
    QuantDisponivel.Caption = Formata_Estoque(dSaldoAlmoxarifadoPadrao)
    QuantDisponivelTodos.Caption = Formata_Estoque(dSaldoTodosAlmoxarifados)
    
    Exit Function
    
Erro_Preenche_Tela:

    Select Case Err
    
        Case 23740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objItemPedido.sProduto)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159903)
            
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjItemPedido = Nothing
    Set gColItemPedido = Nothing

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FALTA_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Tratamento de Falta de Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FaltaEstoque"
    
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



Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub QuantReservar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantReservar, Source, X, Y)
End Sub

Private Sub QuantReservar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantReservar, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub QuantDisponivel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivel, Source, X, Y)
End Sub

Private Sub QuantDisponivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivel, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub UnidadeMedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidadeMedida, Source, X, Y)
End Sub

Private Sub UnidadeMedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Almoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Almoxarifado, Source, X, Y)
End Sub

Private Sub Almoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Almoxarifado, Button, Shift, X, Y)
End Sub

Private Sub QuantDisponivelTodos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QuantDisponivelTodos, Source, X, Y)
End Sub

Private Sub QuantDisponivelTodos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QuantDisponivelTodos, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

