VERSION 5.00
Begin VB.UserControl ParametrosPtoPedOcx 
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   ScaleHeight     =   1935
   ScaleWidth      =   6345
   Begin VB.CommandButton BotaoCalcular 
      Caption         =   "Calcula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3195
      Picture         =   "ParametrosPtoPedOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   225
      Width           =   1230
   End
   Begin VB.CommandButton BotaoFechar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4800
      Picture         =   "ParametrosPtoPedOcx.ctx":0172
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Fechar"
      Top             =   225
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parâmetros"
      Height          =   1605
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   2565
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Consumo Médio"
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
         Left            =   240
         TabIndex        =   4
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tempo de Ressuprimento"
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
         Left            =   240
         TabIndex        =   3
         Top             =   630
         Width           =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estoque de Segurança"
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
         Left            =   240
         TabIndex        =   2
         Top             =   930
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ponto de Pedido"
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
         Left            =   240
         TabIndex        =   1
         Top             =   1230
         Width           =   1425
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Último Cálculo:"
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
      Left            =   3090
      TabIndex        =   6
      Top             =   1485
      Width           =   1755
   End
   Begin VB.Label DataCalculoPtoPedido 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4980
      TabIndex        =   5
      Top             =   1440
      Width           =   1155
   End
End
Attribute VB_Name = "ParametrosPtoPedOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Quando esse cálculo é disparado, calculamos as quantidades recebidas e
'o tempo de ressup médio para cada mês envolvido (se já não estiver calculado no BD)
'e armazenamos no BD. Só pega mês fechado.

'**** inicio do trecho a ser copiado *****

Public Sub Form_Load()

Dim lErro As Long
Dim objComprasConfig As New ClassComprasConfig

On Error GoTo Erro_Form_Load

    'Ler a Data do Último Calculo
    objComprasConfig.sCodigo = COMPRAS_CONFIG_DATA_CALCULO_PTO_PEDIDO
    objComprasConfig.iFilialEmpresa = EMPRESA_TODA
    
    'Le o Conteudo de ComprasConfig para o Codigo passado
    lErro = CF("ComprasConfig_Le_Conteudo",objComprasConfig)
    If lErro <> SUCESSO Then Error 64257
    
    'Coloca a data na Tela
    If CDate(objComprasConfig.sConteudo) <> DATA_NULA Then DataCalculoPtoPedido.Caption = Format(objComprasConfig.sConteudo, "dd/mm/yyyy")
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case 64257
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164301)

    End Select

    Exit Sub

End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Parâmetros de Ponto de Pedido"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ParametrosPtoPed"

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

Private Sub BotaoCalcular_Click()

Dim lErro As Long
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoCalcular_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 69101
    
    lErro = CF("Rotina_CalculaPtoPedido",sNomeArqParam)
    If lErro <> SUCESSO Then gError 64258
    
    Unload Me
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoCalcular_Click:

    Select Case gErr
        
        Case 64258, 69101
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164302)
    
    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
        
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
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

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    
    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)
    
    Parent.Caption = New_Caption
    m_Caption = New_Caption

End Property

''**** fim do trecho a ser copiado *****


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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub DataCalculoPtoPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataCalculoPtoPedido, Source, X, Y)
End Sub

Private Sub DataCalculoPtoPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataCalculoPtoPedido, Button, Shift, X, Y)
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'Function ParametrosPtoPed_Calcula() As Long
''Calculas os Parametros de Ponto Pedido e atualiza a Tabela de Produto Filial
'
'Dim lErro As Long
'Dim tProdutoFilial As typeProdutoFilial
'Dim objProdutoFilial As New ClassProdutoFilial
'Dim lComando1 As Long
'Dim lComando2 As Long
'Dim lTransacao As Long
'Dim objComprasConfig As New ClassComprasConfig
'
'On Error GoTo Erro_ParametrosPtoPed_Calcula
'
'    'Inicia a Transacao
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then Error 64263
'
'    'Abertura comando
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then Error 64264
'
'    lComando2 = Comando_Abrir()
'    If lComando2 = 0 Then Error 64265
'
'    tProdutoFilial.sClasseABC = String(STRING_PRODUTOFILIAL_CLASSEABC, 0)
'    tProdutoFilial.sProduto = String(STRING_PRODUTO, 0)
'
'    'Pesquisa no BD ProdutoFilial
'    lErro = Comando_ExecutarPos(lComando1, "SELECT Produto, Almoxarifado, Fornecedor, FilialForn, VisibilidadeAlmoxarifados, EstoqueSeguranca, ESAuto, EstoqueMaximo, TemPtoPedido, PontoPedido, PPAuto, ClasseABC, LoteEconomico, IntRessup, TempoRessup, TRAuto, TempoRessupMax, ConsumoMedio, CMAuto, ConsumoMedioMax, MesesConsumoMedio FROM ProdutosFilial WHERE FilialEmpresa= ? ORDER BY Produto", 0, tProdutoFilial.sProduto, tProdutoFilial.iAlmoxarifado, tProdutoFilial.lFornecedor, tProdutoFilial.iFilialForn, tProdutoFilial.iVisibilidadeAlmoxarifados, tProdutoFilial.dEstoqueSeguranca, tProdutoFilial.iESAuto, tProdutoFilial.dEstoqueMaximo, tProdutoFilial.iTemPtoPedido, _
'    tProdutoFilial.dPontoPedido, tProdutoFilial.iPPAuto, tProdutoFilial.sClasseABC, tProdutoFilial.dLoteEconomico, tProdutoFilial.iIntRessup, tProdutoFilial.iTempoRessup, tProdutoFilial.iTRAuto, tProdutoFilial.dTempoRessupMax, tProdutoFilial.dConsumoMedio, tProdutoFilial.iCMAuto, tProdutoFilial.dConsumoMedioMax, tProdutoFilial.iMesesConsumoMedio, giFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then Error 64266
'
'    'Tenta selecionar Produto
'    lErro = Comando_BuscarPrimeiro(lComando1)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64267
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set objProdutoFilial = New ClassProdutoFilial
'
'        objProdutoFilial.sProduto = tProdutoFilial.sProduto
'        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
'        objProdutoFilial.iAlmoxarifado = tProdutoFilial.iAlmoxarifado
'        objProdutoFilial.lFornecedor = tProdutoFilial.lFornecedor
'        objProdutoFilial.iFilialForn = tProdutoFilial.iFilialForn
'        objProdutoFilial.iVisibilidadeAlmoxarifados = tProdutoFilial.iVisibilidadeAlmoxarifados
'        objProdutoFilial.dEstoqueSeguranca = tProdutoFilial.dEstoqueSeguranca
'        objProdutoFilial.iESCalculado = tProdutoFilial.iESAuto
'        objProdutoFilial.dEstoqueMaximo = tProdutoFilial.dEstoqueMaximo
'        objProdutoFilial.iTemPtoPedido = tProdutoFilial.iTemPtoPedido
'        objProdutoFilial.dPontoPedido = tProdutoFilial.dPontoPedido
'        objProdutoFilial.iPPCalculado = tProdutoFilial.iPPAuto
'        objProdutoFilial.sClasseABC = tProdutoFilial.sClasseABC
'        objProdutoFilial.dLoteEconomico = tProdutoFilial.dLoteEconomico
'        objProdutoFilial.iIntRessup = tProdutoFilial.iIntRessup
'        objProdutoFilial.iTempoRessup = tProdutoFilial.iTempoRessup
'        objProdutoFilial.iTRCalculado = tProdutoFilial.iTRAuto
'        objProdutoFilial.dTempoRessupMax = tProdutoFilial.dTempoRessupMax
'        objProdutoFilial.dConsumoMedio = tProdutoFilial.dConsumoMedio
'        objProdutoFilial.iCMCalculado = tProdutoFilial.iCMAuto
'        objProdutoFilial.dConsumoMedioMax = tProdutoFilial.dConsumoMedioMax
'        objProdutoFilial.iMesesConsumoMedio = tProdutoFilial.iMesesConsumoMedio
'
'        'Calcula o Consumo Médio para este Produto
'        lErro = Produto_Calcula_ConsumoMedio(objProdutoFilial)
'        If lErro <> SUCESSO Then Error 64268
'
'        'Calcula o Tempo de Ressuprimento para este Produto
'        lErro = Produto_Calcula_TempoRessuprimento(objProdutoFilial)
'        If lErro <> SUCESSO Then Error 64269
'
'        'Calcula o Estoque de Segurança para este Produto
'        lErro = Produto_Calcula_EstoqueSeguranca(objProdutoFilial)
'        If lErro <> SUCESSO Then Error 64270
'
'        'Calcula o Ponto de Pedido para este Produto
'        lErro = Produto_Calcula_PontoPedido(objProdutoFilial)
'        If lErro <> SUCESSO Then Error 64271
'
'        'Atualiza a Tabela ProdutoFilial para os valores calculados
'        'Utilizar o lComando2
'        lErro = Comando_ExecutarPos(lComando2, "UPDATE ProdutosFilial SET EstoqueSeguranca = ?, PontoPedido= ?, TempoRessup = ?, ConsumoMedio = ?", lComando1, objProdutoFilial.dEstoqueSeguranca, objProdutoFilial.dPontoPedido, objProdutoFilial.iTempoRessup, objProdutoFilial.dConsumoMedio)
'        If lErro <> AD_SQL_SUCESSO Then Error 64272
'
'        'Tenta selecionar ProdutoFilial
'        lErro = Comando_BuscarProximo(lComando1)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64273
'
'    Loop
'
'    objComprasConfig.sCodigo = COMPRAS_CONFIG_DATA_CALCULO_PTO_PEDIDO
'    objComprasConfig.iFilialEmpresa = EMPRESA_TODA
'    objComprasConfig.sConteudo = gdtDataAtual
'
'    'Atualiza a data de último cálculo no BD
'    lErro = ComprasConfig_Atualiza_Conteudo_Trans(objComprasConfig)
'    If lErro <> SUCESSO Then Error 64274
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    'Confirma a Transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then Error 64275
'
'    ParametrosPtoPed_Calcula = SUCESSO
'
'    Exit Function
'
'Erro_ParametrosPtoPed_Calcula:
'
'    ParametrosPtoPed_Calcula = Err
'
'    Select Case Err
'
'        Case 64263
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
'
'        Case 64264, 64265
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 64266, 64267, 64273
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL2", Err)
'
'        Case 64268, 64269, 64270, 64271, 64274 'Tratados nas rotinas chamadas
'
'        Case 64272
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", Err, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)
'
'        Case 64275
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164303)
'
'    End Select
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    Call Transacao_Rollback
'
'    Exit Function
'
'End Function
'
'Function Produto_Calcula_ConsumoMedio(objProdutoFilial As ClassProdutoFilial) As Long
''Calcula o Consumo Medio para o ProdutoFilial
'
'Dim lErro As Long
'Dim objComprasConfig As New ClassComprasConfig
'Dim iMesesConfig As Integer
'Dim tEstoqueMes As typeEstoqueMes
'Dim iContaMes As Integer
'Dim objSldMesEst As New ClassSldMesEst
'Dim dQuantidadeConsumidaTotal As Double
'Dim lComando1 As Long
'
'On Error GoTo Erro_Produto_Calcula_ConsumoMedio
'
'    If objProdutoFilial.iCMCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
'
'        'Abertura comando
'        lComando1 = Comando_Abrir()
'        If lComando1 = 0 Then Error 64276
'
'        '???? Porque você não usa o gobjCOM para não precisar ler. Inclua esse campo nele e inclua esse código no select
'        '???? Onde ele está sendo gravado? Na tela de Configuração. Por favor quando souber me avise.
'        objComprasConfig.sCodigo = COMPRAS_CONFIG_MESES_CONSUMO_MEDIO
'        objComprasConfig.iFilialEmpresa = EMPRESA_TODA
'
'        'Lê o número de meses que serão calculados o Consumo Médio
'        lErro = CF("ComprasConfig_Le_Conteudo",objComprasConfig)
'        If lErro <> SUCESSO Then Error 64277
'
'        iMesesConfig = CInt(objComprasConfig.sConteudo)
'
'        'Le os Ultimos Meses Fechados
'        lErro = Comando_Executar(lComando1, "SELECT FilialEmpresa, Ano, Mes FROM EstoqueMes WHERE Fechamento = ? AND FilialEmpresa = ? ORDER BY Ano DESC, Mes DESC", tEstoqueMes.iFilialEmpresa, tEstoqueMes.iAno, tEstoqueMes.iMes, ESTOQUEMES_FECHAMENTO_FECHADO, giFilialEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then Error 64278
'
'        lErro = Comando_BuscarPrimeiro(lComando1)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64279
'
'        'Para cada Mês
'        Do While (iContaMes < iMesesConfig And lErro = AD_SQL_SUCESSO)
'
'            iContaMes = iContaMes + 1
'
'            objSldMesEst.iAno = tEstoqueMes.iAno
'            objSldMesEst.sProduto = objProdutoFilial.sProduto
'            objSldMesEst.iFilialEmpresa = giFilialEmpresa
'
'            'Le a Quantidade Consumida e a Quantidade de Venda
'            lErro = CF("SldMesEst_Le",objSldMesEst)
'            If lErro <> SUCESSO And lErro <> 25429 Then Error 64280
'
'            dQuantidadeConsumidaTotal = dQuantidadeConsumidaTotal + (objSldMesEst.dQuantCons(tEstoqueMes.iMes) + objSldMesEst.dQuantVend(tEstoqueMes.iMes))
'
'            'Tenta selecionar ProdutoFilial
'            lErro = Comando_BuscarProximo(lComando1)
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64281
'
'        Loop
'
'        'Calcula o Consumo Médio
'        objProdutoFilial.dConsumoMedio = dQuantidadeConsumidaTotal / (iContaMes + 1)
'
'    End If
'
'    Call Comando_Fechar(lComando1)
'
'    Produto_Calcula_ConsumoMedio = SUCESSO
'
'    Exit Function
'
'Erro_Produto_Calcula_ConsumoMedio:
'
'    Produto_Calcula_ConsumoMedio = Err
'
'    Select Case Err
'
'        Case 64276
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 64277, 64280 'Tratados nas rotinas chamadas
'
'        Case 64278, 64279, 64281
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164304)
'
'    End Select
'
'    Call Comando_Fechar(lComando1)
'
'    Exit Function
'
'End Function
'
'Function Produto_Calcula_TempoRessuprimento(objProdutoFilial As ClassProdutoFilial) As Long
''Calcula o Tempo de Ressuprimento para o ProdutoFilial
'
'Dim lErro As Long
'Dim iMesesRessup As Integer
'Dim iNumeroComprasRessup As Integer
'Dim tEstoqueMes As typeEstoqueMes
'Dim dtLimiteMaior As Date
'Dim dtLimiteMenor As Date
'Dim iContaPedido As Integer
'Dim dQuantRecebidaPC As Double
'Dim dtDataEmissaoPC As Date
'Dim dtDataEntradaNF As Date
'Dim iDiasRecebimentoParcial As Integer
'Dim iDiasRecebimentoTotal As Integer
'Dim dValorParcial As Double
'Dim lComando1 As Long
'Dim lComando2 As Long
'Dim objComprasConfig As New ClassComprasConfig
'Dim iContaMes As Integer
'Dim dQuantidadeRecebidaTotal As Double
'Dim sSQL As String
'
'On Error GoTo Erro_Produto_Calcula_TempoRessuprimento
'
'    If objProdutoFilial.iTRCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
'
'        'Abertura comando
'        lComando1 = Comando_Abrir()
'        If lComando1 = 0 Then gError 64282
'
'        lComando2 = Comando_Abrir()
'        If lComando2 = 0 Then gError 64283
'
'        objComprasConfig.sCodigo = COMPRAS_CONFIG_MESES_MEDIA_TEMPO_RESSUP
'        objComprasConfig.iFilialEmpresa = EMPRESA_TODA
'
'        '??? Se você utilizar o gobjCOM, você não vai precisar ir ao BD nunca
'        '??? Inclui no obj e na leitura esse campo
'        'Lê o número de meses que seram calculados o Consumo Médio
'        lErro = CF("ComprasConfig_Le_Conteudo",objComprasConfig)
'        If lErro <> SUCESSO Then gError 64284
'
'        iMesesRessup = CInt(objComprasConfig.sConteudo)
'
'        '???? ONDE ESSAS INFORMAÇÕES SÃO GRAVADAS, em que tela?
'        '??? Se você utilizar o gobjCOM, você não vai precisar ir ao BD nunca
'        '??? Inclui no obj e na leitura esse campo
'        objComprasConfig.sCodigo = COMPRAS_CONFIG_NUM_COMPRAS_TEMPO_RESSUP
'        objComprasConfig.iFilialEmpresa = EMPRESA_TODA
'
'        'Lê o número de meses que seram calculados o Consumo Médio
'        lErro = CF("ComprasConfig_Le_Conteudo",objComprasConfig)
'        If lErro <> SUCESSO Then gError 64285
'
'        iNumeroComprasRessup = CInt(objComprasConfig.sConteudo)
'
'        'Le os Ultimos Meses Fechados
'        lErro = Comando_Executar(lComando1, "SELECT FilialEmpresa, Ano, Mes FROM EstoqueMes WHERE Fechamento = ? AND FilialEmpresa = ? ORDER BY Ano DESC, Mes DESC", tEstoqueMes.iFilialEmpresa, tEstoqueMes.iAno, tEstoqueMes.iMes, ESTOQUEMES_FECHAMENTO_FECHADO, giFilialEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then gError 64286
'
'        lErro = Comando_BuscarPrimeiro(lComando1)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64287
'
'        'Se não encontrou nenhum mês fechado, erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 67459
'
'        'Calcula a ultima data Limite (Ultimo dia do Mes Lido /MesLido/AnoLido)
'        If tEstoqueMes.iMes = 12 Then
'            dtLimiteMaior = CDate("31/" & tEstoqueMes.iMes & "/" & tEstoqueMes.iAno)
'        Else
'            dtLimiteMaior = CDate("01/" & (tEstoqueMes.iMes + 1) & "/" & tEstoqueMes.iAno) - 1
'        End If
'
'        Do While (iContaMes < iMesesRessup And lErro = AD_SQL_SUCESSO)
'
'            iContaMes = iContaMes + 1
'
'            'Tenta selecionar ProdutoFilial
'            lErro = Comando_BuscarProximo(lComando1)
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64288
'
'        Loop
'
'        'Calcula a Primeira data Limite (01/MesLido/AnoLido)
'        dtLimiteMenor = CDate("01/" & tEstoqueMes.iMes & "/" & tEstoqueMes.iAno)
'
'        'Comando SQL a executar abaixo
'        sSQL = "SELECT SUM(ItensPedCompra.QuantRecebida), MAX(PedidoCompra.DataEmissao), MAX(NFiscal.DataEntrada) FROM ItemNFItemPC, ItensPedCompra, ItensNFiscal, PedidoCompra, NFiscal WHERE ItensPedCompra.NumIntDoc = ItemNFItemPC.ItemPedCompra AND PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItemNFItemPC.ItemNFiscal = ItensNFiscal.NumIntDoc AND ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND ItensPedCompra.Produto = ? AND PedidoCompra.FilialEmpresa = ? AND PedidoCompra.DataEmissao >= ? AND  PedidoCompra.DataEmissao <= ? GROUP BY PedidoCompra.NumIntDoc ORDER BY PedidoCompra.NumIntDoc"
'
'        'Faz select em Todos os Pedidos de Compra (Baixados ou não)
'        lErro = Comando_Executar(lComando2, sSQL, dQuantRecebidaPC, dtDataEmissaoPC, dtDataEntradaNF, objProdutoFilial.sProduto, giFilialEmpresa, dtLimiteMenor, dtLimiteMaior)
'        If lErro <> AD_SQL_SUCESSO Then gError 64289
'
'        lErro = Comando_BuscarPrimeiro(lComando2)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64290
'
'        Do While iContaPedido < iNumeroComprasRessup And lErro = AD_SQL_SUCESSO
'
'            iContaPedido = iContaPedido + 1
'            iDiasRecebimentoParcial = dtDataEntradaNF - dtDataEmissaoPC
'            iDiasRecebimentoTotal = iDiasRecebimentoTotal + iDiasRecebimentoParcial
'            dQuantidadeRecebidaTotal = dQuantidadeRecebidaTotal + dQuantRecebidaPC
'            dValorParcial = dValorParcial + (dQuantRecebidaPC * iDiasRecebimentoParcial)
'
'            'Tenta selecionar ProdutoFilial
'            lErro = Comando_BuscarProximo(lComando2)
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64291
'
'        Loop
'
'        'Calcula o Tempo de Ressuprimento
'        If iDiasRecebimentoTotal > 0 Then objProdutoFilial.iTempoRessup = dValorParcial / dQuantidadeRecebidaTotal
'
'    End If
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    Produto_Calcula_TempoRessuprimento = SUCESSO
'
'    Exit Function
'
'Erro_Produto_Calcula_TempoRessuprimento:
'
'    Produto_Calcula_TempoRessuprimento = gErr
'
'    Select Case gErr
'
'        Case 64282, 64283
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 64284, 64285 'Tratados nas Rotinas chamadas
'
'        Case 64286, 64287, 64288
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", gErr)
'
'        Case 64289, 64290, 64291
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)
'
'        Case 67459
'            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_MES_FECHADO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164305)
'
'    End Select
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    Exit Function
'
'End Function
'
'Function Produto_Calcula_EstoqueSeguranca(objProdutoFilial As ClassProdutoFilial) As Long
''Calcula o Estoque de Segurança
'
'Dim lErro As Long
'Dim objTipoDeProduto As New ClassTipoDeProduto
'Dim objProduto As New ClassProduto
'Dim objComprasConfig As New ClassComprasConfig
'
'On Error GoTo Erro_Produto_Calcula_EstoqueSeguranca
'
'    'Se é Para calcular o Estoque de segurança
'    If objProdutoFilial.iESCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
'
'        'Verifica se já foi Lido o Consumo Médio MAX e o Tempo de Ressuprimento MAX
'        If objProdutoFilial.dConsumoMedioMax = 0 Or objProdutoFilial.dTempoRessupMax = 0 Then
'
'            objProduto.sCodigo = objProdutoFilial.sProduto
'
'            'Lê o Produto para Pegar o Tipo de Produto
'            lErro = CF("Produto_Le",objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then Error 64292
'
'            'Se não encontrou --> Erro
'            If lErro = 28030 Then Error 64293
'
'            If objProduto.iTipo > 0 Then
'
'                objTipoDeProduto.iTipo = objProduto.iTipo
'
'                'Le o Tipo de Produto para ver se o Consumo Medio Max e o Tempo de Ressuprimento Max estão Preenchidos
'                lErro = CF("TipoDeProduto_Le",objTipoDeProduto)
'                If lErro <> SUCESSO And lErro <> 22531 Then Error 64294
'
'                'Se não encontrou --> Erro
'                If lErro = 22531 Then Error 64295
'
'                objProdutoFilial.dConsumoMedioMax = objTipoDeProduto.dConsumoMedioMax
'                objProdutoFilial.dTempoRessupMax = objTipoDeProduto.dTempoRessupMax
'
'            End If
'
'            If objProdutoFilial.dConsumoMedioMax = 0 Or objProdutoFilial.dTempoRessupMax = 0 Then
'
'                objComprasConfig.sCodigo = COMPRAS_CONFIG_CONSUMO_MEDIO_MAX
'                objComprasConfig.iFilialEmpresa = EMPRESA_TODA
'
'                'Lê o número de meses que seram calculados o Consumo Médio
'                lErro = CF("ComprasConfig_Le_Conteudo",objComprasConfig)
'                If lErro <> SUCESSO Then Error 64296
'
'                objProdutoFilial.dConsumoMedioMax = CDbl(objComprasConfig.sConteudo)
'
'                objComprasConfig.sCodigo = COMPRAS_CONFIG_TEMPO_RESSUP_MAX
'                objComprasConfig.iFilialEmpresa = EMPRESA_TODA
'
'                'Lê o número de meses que seram calculados o Consumo Médio
'                lErro = CF("ComprasConfig_Le_Conteudo",objComprasConfig)
'                If lErro <> SUCESSO Then Error 64297
'
'                objProdutoFilial.dTempoRessupMax = CDbl(objComprasConfig.sConteudo)
'
'            End If
'
'        End If
'
'        'Calcula o Estoque de segurança
'        objProdutoFilial.dEstoqueSeguranca = ((objProdutoFilial.dConsumoMedioMax - objProdutoFilial.dConsumoMedio) * objProdutoFilial.iTempoRessup) + objProdutoFilial.dConsumoMedioMax * (objProdutoFilial.dTempoRessupMax - objProdutoFilial.iTempoRessup)
'
'    End If
'
'    Produto_Calcula_EstoqueSeguranca = SUCESSO
'
'    Exit Function
'
'Erro_Produto_Calcula_EstoqueSeguranca:
'
'    Produto_Calcula_EstoqueSeguranca = Err
'
'    Select Case Err
'
'        Case 64292, 64294, 64296, 64297 'Tratados nas Rotinas chamadas
'
'        Case 64293
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
'
'        Case 64295
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoDeProduto.iTipo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164306)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Produto_Calcula_PontoPedido(objProdutoFilial As ClassProdutoFilial) As Long
''Calcula o Ponto de Pedido
'
'On Error GoTo Erro_Produto_Calcula_PontoPedido
'
'    'Verifica se é Para calcular o Ponto de Pedido
'    If objProdutoFilial.iPPCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
'
'        'Calcula o Ponto de Pedido
'        objProdutoFilial.dPontoPedido = (objProdutoFilial.dConsumoMedio * objProdutoFilial.iTempoRessup) + objProdutoFilial.dEstoqueSeguranca
'
'    End If
'
'    Produto_Calcula_PontoPedido = SUCESSO
'
'    Exit Function
'
'Erro_Produto_Calcula_PontoPedido:
'
'    Produto_Calcula_PontoPedido = Err
'
'    Select Case Err
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164307)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function ComprasConfig_Atualiza_Conteudo_Trans(objComprasConfig As ClassComprasConfig) As Long
''Atualiza o Conteudo na Tabela de Compras Config para o Codigo e Filial Passados
'
'Dim lErro As Long
'Dim sConteudo As String
'Dim lComando1 As Long
'Dim lComando2 As Long
'Dim lTransacao As Long
'
'On Error GoTo Erro_ComprasConfig_Atualiza_Conteudo_Trans
'
'    'Abertura comando
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then Error 64298
'
'    'Abertura comando
'    lComando2 = Comando_Abrir()
'    If lComando2 = 0 Then Error 64299
'
'    sConteudo = String(STRING_CONTEUDO, 0)
'
'    'Ler registo
'    lErro = Comando_ExecutarPos(lComando1, "SELECT Conteudo FROM ComprasConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then Error 64300
'
'    'Lê o primeiro registro
'    lErro = Comando_BuscarPrimeiro(lComando1)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64301
'
'    'Se não encontrou o registro
'    If lErro = AD_SQL_SEM_DADOS Then Error 64302
'
'    If sConteudo <> objComprasConfig.sConteudo Then
'
'        'Atualiza o conteudo do código passado
'        lErro = Comando_ExecutarPos(lComando2, "UPDATE ComprasConfig SET Conteudo = ?", lComando1, objComprasConfig.sConteudo)
'        If lErro <> AD_SQL_SUCESSO Then Error 64303
'
'    End If
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    ComprasConfig_Atualiza_Conteudo_Trans = SUCESSO
'
'    Exit Function
'
'Erro_ComprasConfig_Atualiza_Conteudo_Trans:
'
'    ComprasConfig_Atualiza_Conteudo_Trans = Err
'
'    Select Case Err
'
'        Case 64298, 64299
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 64300, 64301
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRASCONFIG", Err, objComprasConfig.sCodigo)
'
'        Case 64302
'            Call Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO", Err, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
'
'        Case 64303
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPRASCONFIG", Err, objComprasConfig.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164308)
'
'    End Select
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    Exit Function
'
'End Function
'
