VERSION 5.00
Begin VB.UserControl RelOpCustoCOMCustoFPOcx 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   ScaleHeight     =   2295
   ScaleWidth      =   6870
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
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
      Left            =   2528
      Picture         =   "RelOpCustoCOMCustoFP.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condições"
      Height          =   1275
      Left            =   195
      TabIndex        =   1
      Top             =   150
      Width           =   6480
      Begin VB.CheckBox PrecoMaior 
         Caption         =   "Preço do último pedido de Compras maior que o utilizado para formação de preços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   300
         TabIndex        =   2
         Top             =   285
         Value           =   1  'Checked
         Width           =   5940
      End
   End
End
Attribute VB_Name = "RelOpCustoCOMCustoFPOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 123128
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 123128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 123129
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167943)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
  
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167944)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 123139
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 123139

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167945)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long, lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 106962
         
    If bExecutar Then
    
        lErro = CF("RelCustoCOMCustoFP_Prepara", giFilialEmpresa, lNumIntRel)
        If lErro <> SUCESSO Then gError 106963
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 106964
    
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 106962, 106963, 106964
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167946)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Preços de Compra vs Custo para Formação de Preços"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpCustoCOMCustoFP"
    
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

Public Sub Unload(objme As Object)
    
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

'??? transferir p/fatgrava
Function RelCustoCOMCustoFP_Prepara(ByVal iFilialEmpresa As Integer, lNumIntRel As Long) As Long
'Insere registros em RElCustoComCustoFP referentes a produtos atributos para efeito de formacao de precos
'comparados com o preço utilizado no ultimo pedido de compras enviado.

Dim alComando(1 To 3) As Long
Dim lErro As Long
Dim lTransacao As Long
Dim sProduto As String, iClasseUM As Integer
Dim lNumIntItemPedCompra As Long
Dim sFPUM As String, dFPCusto As Double, dtFPDataAtualizacao As Date, dFPAliquotaICMS As Double, iFPCondicaoPagto As Integer
Dim iIndice As Integer, dFator As Double
Dim lNumIntDoc As Long, dtData As Date, sUM As String, iMoeda As Integer, dTaxa As Double, dAliquotaICMS As Double, iCondicaoPagto As Integer, dPrecoUnitario As Double
Dim objCotacao As ClassCotacaoMoeda
Dim objCotacaoAnterior As ClassCotacaoMoeda

On Error GoTo Erro_RelCustoCOMCustoFP_Prepara
    
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 123211
    
    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 123212
    Next
    
    'Obtem o próximo NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_CUSTOFPCOM", lNumIntRel)
    If lErro <> SUCESSO Then gError 123217
        
    'inicializa as strings
    sProduto = String(STRING_PRODUTO, 0)
    sFPUM = String(STRING_UM_SIGLA, 0)
    sUM = String(STRING_UM_SIGLA, 0)
    
    'Busca a maior data daquele produto
    lErro = Comando_Executar(alComando(1), "SELECT Produtos.ClasseUM, Produtos.SiglaUMEstoque, Produto, Custo, DataAtualizacao, AliquotaICMS, CondicaoPagto FROM CustoEmbMP, Produtos WHERE CustoEmbMP.Produto = Produtos.Codigo AND FilialEmpresa = ? ORDER BY Produto", _
        iClasseUM, sFPUM, sProduto, dFPCusto, dtFPDataAtualizacao, dFPAliquotaICMS, iFPCondicaoPagto, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 123213
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 123214
    
    Do While lErro = AD_SQL_SUCESSO
    
        '??? depois ver necessidade de considerar desconto
    
        'Busca informacoes sobre o ultimo pedido de compras enviado referente ao produto.
        lErro = Comando_Executar(alComando(2), "SELECT ItensPedCompraN.UM, ItensPedCompraN.NumIntDoc, ItensPedCompraN.PrecoUnitario, ItensPedCompraN.AliquotaICMS, PedidoCompraN.CondicaoPagto, PedidoCompraN.Moeda, PedidoCompraN.Taxa FROM ItensPedCompraN, PedidoCompraN WHERE ItensPedCompraN.PedCompra = PedidoCompraN.NumIntDoc AND PedidoCompraN.FilialEmpresa = ? AND PedidoCompraN.DataEnvio <> ? AND ItensPedCompraN.Produto = ? ORDER BY PedidoCompraN.Data DESC, PedidoCompraN.Codigo DESC", _
            sUM, lNumIntDoc, dPrecoUnitario, dAliquotaICMS, iCondicaoPagto, iMoeda, dTaxa, iFilialEmpresa, DATA_NULA, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 123215
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 123216
    
        If lErro = AD_SQL_SUCESSO Then
        
            'Verifica se a moeda não é a Real
            If iMoeda <> MOEDA_REAL Then
                
                'Verifica se a taxa foi preenchida e realiza a conversão
                If dTaxa = 0 Then
                
                    objCotacao.dtData = gdtDataAtual
                    objCotacao.iMoeda = iMoeda
                    lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
                    If lErro <> SUCESSO Then gError 106960
                    
                    If objCotacao.dValor <> 0 Then
                        dFPCusto = dFPCusto * objCotacao.dValor
                    Else
                        If objCotacaoAnterior.dValor <> 0 Then
                            dFPCusto = dFPCusto * objCotacaoAnterior.dValor
                        End If
                    End If
                
                Else
                    
                    dFPCusto = dFPCusto * dTaxa
                    
                End If
                
            End If
            
            If sFPUM <> sUM Then
                
                lErro = CF("UM_Conversao", iClasseUM, sFPUM, sUM, dFator)
                If lErro <> SUCESSO Then gError 106961
                
                dPrecoUnitario = dPrecoUnitario * dFator
                
            End If
            
            '??? depois incluir diferencas de aliqicms, condpgto
            
            If (dPrecoUnitario - dFPCusto) > DELTA_VALORMONETARIO Then
            
                'Realiza a gravação
                lErro = Comando_Executar(alComando(3), "INSERT INTO RelCustoCOMCustoFP (NumIntRel, NumIntItemPedCompra, FPCusto, FPDataAtualizacao, FPAliquotaICMS, FPCondicaoPagto) VALUES (?,?,?,?,?,?)", lNumIntRel, lNumIntItemPedCompra, dFPCusto, dtFPDataAtualizacao, dFPAliquotaICMS, iFPCondicaoPagto)
                If lErro <> AD_SQL_SUCESSO Then gError 123218
        
            End If
    
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 123219
        
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 123220
    
    RelCustoCOMCustoFP_Prepara = SUCESSO
            
    Exit Function
    
Erro_RelCustoCOMCustoFP_Prepara:

    RelCustoCOMCustoFP_Prepara = gErr
    
    Select Case gErr
        
        Case 123211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 123212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 123213, 123214, 123215, 123216, 123219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTOEMBMP", gErr)
            
        Case 123217, 106960, 106961
        
        Case 123218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCUSTOCOMCUSTOFP", gErr)
            
        Case 123220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167947)
        
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function



