VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpFlCx 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   ScaleHeight     =   1605
   ScaleWidth      =   4305
   Begin VB.Frame Frame1 
      Caption         =   "Fluxo de Caixa"
      Height          =   1290
      Left            =   165
      TabIndex        =   4
      Top             =   105
      Width           =   2505
      Begin MSMask.MaskEdBox NumDeDias 
         Height          =   300
         Left            =   1710
         TabIndex        =   6
         Top             =   487
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Número de Dias:"
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
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   525
         Width           =   1470
      End
   End
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
      Left            =   2805
      Picture         =   "RelOpFlCxOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   810
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   2820
      ScaleHeight     =   495
      ScaleWidth      =   1170
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Width           =   1230
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpFlCxOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   645
         Picture         =   "RelOpFlCxOcx.ctx":0634
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
End
Attribute VB_Name = "RelOpFlCx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const RELFLCX_PARC_REC = 1
Const RELFLCX_BX_PARC_REC = 2
Const RELFLCX_PARC_PAG = 3
Const RELFLCX_BX_PARC_PAG = 4
Const RELFLCX_SAQUE = 5
Const RELFLCX_DEPOSITO = 6
Const RELFLCX_APLICACAO = 7
Const RELFLCX_RESGATE = 8

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long, lNumIntRel As Long
Dim sNumDeDias As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sNumDeDias)
    If lErro <> SUCESSO Then gError 123210
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 123202
    
    'Inclui o número de dias
    lErro = objRelOpcoes.IncluirParametro("NDIAS", sNumDeDias)
    If lErro <> AD_BOOL_TRUE Then gError 123203

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 123204

    If bExecutar Then
    
        lErro = RelOpFlCx_Prepara(giFilialEmpresa, StrParaInt(sNumDeDias), lNumIntRel)
        If lErro <> SUCESSO Then gError 33333
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 123203
    
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 123202 To 123204, 123210

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sNumDeDias As String) As Long
'E critica o Número de dias

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica o número de dias
    If NumDeDias.ClipText = "" Then gError 123209
    
    sNumDeDias = CStr(NumDeDias.ClipText)
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 123209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DIAS_NAO_PREENCHIDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 123205

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 123205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 123206
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 123206

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela(Me)
        
    'Realiza a chamada da define padrao
    lErro = Define_Padrao
    If lErro <> SUCESSO Then gError 123207
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
    
        Case 123207

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Realiza a chamada da define padrao
    lErro = Define_Padrao
    If lErro <> SUCESSO Then gError 123208

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 123208
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Function Define_Padrao() As Long
'Define os campos a serem preenchidos como default

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    'Preenche o número de dias
    NumDeDias.PromptInclude = False
    NumDeDias.Text = "14"
    NumDeDias.PromptInclude = True
    
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSCLI
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpFlCx"

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

Private Sub NumDeDias_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumDeDias)
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

Function RelOpFlCx_Prepara(ByVal iFilialEmpresa As Integer, ByVal iDias As Integer, lNumIntRel As Long) As Long
'preenche tabelas para execucao do relatorio e retorna lNumIntRel

Dim lErro As Long, lTransacao As Long, alComando(1 To 22) As Long, alComando1(1 To 4) As Long, iIndice As Integer, iDia As Integer, dtDataFluxo As Date, dtDataRef As Date
Dim dValor As Double, dAcum As Double, dtData As Date, colRelItem As New Collection, lTrans As Long
Dim objRelFluxoCx As ClassRelFluxoCx, iSequencial As Integer
Dim dSaldoInicial As Double, dSaldoFinal As Double, dSaldoAplicacoes As Double, dSaldoFinalTotal As Double
Dim iCodCCI As Integer, dtDataBase As Date, dSaldoInicialAnterior As Double, dSaldoAplicacoesAnterior As Double

On Error GoTo Erro_RelOpFlCx_Prepara

    dtDataBase = gdtDataAtual
    
    lTrans = Transacao_Abrir
    If lTrans = 0 Then gError 33333
    
    'obter lNumIntRel
    lErro = CF("Config_ObterNumInt", "TESConfig", "NUM_PROX_REL_FLCXMIGUEZ", lNumIntRel)
    If lErro <> SUCESSO Then gError 33333
    
    'fechar transacao
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    'abrir comandos e transacao na conexao rel
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
        If alComando(iIndice) = 0 Then gError 33333
    Next
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
        If alComando1(iIndice) = 0 Then gError 33333
    Next
    
    lTransacao = Transacao_AbrirExt(GL_lConexaoRel)
    If lTransacao = 0 Then gError 33333
    
    'cria uma colecao com itens para cada um dos n dias úteis a partir de dtDataBase
    dtDataRef = dtDataBase
    
    For iDia = 0 To iDias - 1
    
        'obter 1o dia util a partir da data sendo analisada
        lErro = CF("DataVencto_Real", dtDataRef, dtDataFluxo)
        If lErro <> SUCESSO Then gError 33333
        
        Set objRelFluxoCx = New ClassRelFluxoCx
        
        objRelFluxoCx.lNumIntRel = lNumIntRel
        objRelFluxoCx.dtData = dtDataFluxo
        
        colRelItem.Add objRelFluxoCx
        
        'incrementa a proxima data a ser analisada
        dtDataRef = dtDataFluxo + 1
        
        If colRelItem.Count >= (iDias / 7 * 5) Then Exit For
        
    Next
    
    'obter saldo inicial (das contas correntes)
    lErro = Comando_Executar(alComando(22), "SELECT Codigo FROM ContasCorrentesInternas WHERE FilialEmpresa = ?", iCodCCI, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(22))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = CF("CCI_ObterRelTes", iCodCCI, dtDataBase, dValor, alComando1)
        If lErro <> SUCESSO Then gError 33333
    
        dSaldoInicial = dSaldoInicial + Round(dValor, 0)
    
        lErro = Comando_BuscarProximo(alComando(22))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333
    
    Loop
    
    'inicializa sequencial para identificar detalhamento de pagtos e recebimentos
    iSequencial = 1
    
    'preencher dados de parcelas a receber
    For Each objRelFluxoCx In colRelItem
    
        'ignorar as parcelas do dia corrente pois as baixas já terao sido lancadas
        If dtDataBase <> objRelFluxoCx.dtData Then
        
            lErro = RelOpFlCx_Prepara1(objRelFluxoCx, iFilialEmpresa, iSequencial, alComando)
            If lErro <> SUCESSO Then gError 33333
                
        End If
        
    Next
    
    'preencher dados de baixas rec de titulos baixados com movto no dia corrente
    lErro = RelOpFlCx_Prepara6(colRelItem.Item(1), iFilialEmpresa, iSequencial, alComando)
    If lErro <> SUCESSO Then gError 33333
    
    'preencher dados de baixas rec de titulos abertos com movto no dia corrente
    lErro = RelOpFlCx_Prepara7(colRelItem.Item(1), iFilialEmpresa, iSequencial, alComando)
    If lErro <> SUCESSO Then gError 33333
    
    'preencher dados de parcelas a pagar
    For Each objRelFluxoCx In colRelItem
    
        'ignorar as parcelas do dia corrente pois as baixas já terao sido lancadas
        If dtDataBase <> objRelFluxoCx.dtData Then
        
            lErro = RelOpFlCx_Prepara2(objRelFluxoCx, iFilialEmpresa, iSequencial, alComando)
            If lErro <> SUCESSO Then gError 33333
            
        End If
        
    Next
    
    'preencher dados de baixas pag de titulos baixados com movto no dia corrente
    lErro = RelOpFlCx_Prepara8(colRelItem.Item(1), iFilialEmpresa, iSequencial, alComando)
    If lErro <> SUCESSO Then gError 33333
    
    'preencher dados de baixas pag de titulos abertos com movto no dia corrente
    lErro = RelOpFlCx_Prepara9(colRelItem.Item(1), iFilialEmpresa, iSequencial, alComando)
    If lErro <> SUCESSO Then gError 33333
    
    'preencher dados de saques e depositos com movto no dia corrente
    lErro = RelOpFlCx_Prepara3(colRelItem.Item(1), iFilialEmpresa, iSequencial, alComando)
    If lErro <> SUCESSO Then gError 33333
    
    'preencher dados de aplicacoes com movto no dia corrente
    lErro = RelOpFlCx_Prepara4(colRelItem.Item(1), iFilialEmpresa, iSequencial, alComando)
    If lErro <> SUCESSO Then gError 33333
    
    'preencher dados de resgates com movto no dia corrente
    lErro = RelOpFlCx_Prepara5(colRelItem.Item(1), iFilialEmpresa, iSequencial, alComando)
    If lErro <> SUCESSO Then gError 33333
    
    'obter saldo atual de aplicacoes
    lErro = Comando_Executar(alComando(21), "SELECT SUM(SaldoAplicado) FROM Aplicacoes WHERE FilialEmpresa = ? AND Status = ?", dValor, iFilialEmpresa, STATUS_LANCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(21))
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    dValor = Round(dValor, 0)

    'atualizar acumulados diarios
    dSaldoInicialAnterior = dSaldoInicial
    dSaldoAplicacoesAnterior = dValor
    
    For Each objRelFluxoCx In colRelItem
    
        objRelFluxoCx.dSaldoInicial = dSaldoInicialAnterior
        objRelFluxoCx.dSaldoFinal = dSaldoInicialAnterior + objRelFluxoCx.dTotalRec - objRelFluxoCx.dTotalPag
        objRelFluxoCx.dSaldoAplicacoes = dSaldoAplicacoesAnterior
        objRelFluxoCx.dSaldoFinalTotal = objRelFluxoCx.dSaldoFinal + objRelFluxoCx.dSaldoAplicacoes - objRelFluxoCx.dTotalAplic
        
        dSaldoInicialAnterior = objRelFluxoCx.dSaldoFinal
        dSaldoAplicacoesAnterior = objRelFluxoCx.dSaldoAplicacoes - objRelFluxoCx.dTotalAplic
        
    Next
    
    'gravar dados consolidados por dia
    For Each objRelFluxoCx In colRelItem
    
        With objRelFluxoCx
            lErro = Comando_Executar(alComando(5), "INSERT INTO RelOpFluxoCxMiguez (NumIntRel, Data, RecDespesas, TotalRec, DespBanc, PagDesp, TotalPag, SaldoInicial, SaldoFinal, SaldoAplicacoes, TotalAplic, SaldoFinalTotal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntRel, .dtData, .dRecDespesas, .dTotalRec, .dDespBanc, .dPagDesp, .dTotalPag, .dSaldoInicial, .dSaldoFinal, .dSaldoAplicacoes, .dTotalAplic, .dSaldoFinalTotal)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 33333

    Next
    
    'fechar transacao e comandos
    lErro = Transacao_CommitExt(lTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
        
    'Libera os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    RelOpFlCx_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara:

    RelOpFlCx_Prepara = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Transacao_RollbackExt(lTransacao)
    
    'Libera os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    Exit Function

End Function

Function RelOpFlCx_Prepara1(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'obter dados acumulados do dia referentes a recebimentos e gravar dados detalhados

Dim lErro As Long, dtVctoReal As Date
Dim dValor As Double, dRecDespesas As Double, dTotalRec As Double
Dim iFluxoCaixa As Integer, sNatureza As String, sNomeRedCli As String
Dim iNumParcelas As Integer, iNumParcela As Integer, lNumIntDoc As Long

On Error GoTo Erro_RelOpFlCx_Prepara1
    
    'obter dados de parcelasrec com status pendente ou previsao, da filialempresa
    'por enquanto: colocar D+1, inclusive p/facilitar tratamento de previsoes
    'mais correto: do cobrador/carteira obter dias de retencao.
    
    dRecDespesas = 0
    dTotalRec = 0
    
    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sNomeRedCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    
    'obter 1o dia util anterior a data que vai sair no fluxo (D+1)
    lErro = CF("Dias_Uteis_Antes", objRelFluxoCx.dtData, 1, dtVctoReal)
    If lErro <> SUCESSO Then gError 33333
    
    '########################################
    'Alterado por Wagner - (Dar preferência a data de previsão e depois a data do vencimento real -1)
    lErro = Comando_Executar(alComando(1), "SELECT TitRecNatMov.NumParcelas, ParcelasRec.NumParcela, ParcelasRec.NumIntDoc, TitRecNatMov.FluxoCaixa, TitRecNatMov.Natureza, Clientes.NomeReduzido, ParcelasRec.Saldo FROM ParcelasRec, TitRecNatMov, Clientes WHERE ParcelasRec.NumIntTitulo = TitRecNatMov.NumIntDoc AND TitRecNatMov.Cliente = Clientes.Codigo AND TitRecNatMov.FilialEmpresa = ? AND (ParcelasRec.Status = ? OR ParcelasRec.Status = ?) AND ((ParcelasRec.DataVencimentoReal = ? AND ParcelasRec.DataPrevisao = ?) OR (ParcelasRec.DataPrevisao = ?)) ORDER BY TitRecNatMov.NumTitulo, ParcelasRec.NumParcela", _
        iNumParcelas, iNumParcela, lNumIntDoc, iFluxoCaixa, sNatureza, sNomeRedCli, dValor, iFilialEmpresa, STATUS_ABERTO, STATUS_PREVISAO, dtVctoReal, DATA_NULA, objRelFluxoCx.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    '########################################
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        Select Case iFluxoCaixa
        
            Case 3 'outras
                dRecDespesas = dRecDespesas + dValor
                
            Case Else
                'insere registro
                lErro = Comando_Executar(alComando(2), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_RECEBIMENTO, iFluxoCaixa, sNomeRedCli, dValor, iNumParcelas, iNumParcela, RELFLCX_PARC_REC, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 33333
                            
                iSequencial = iSequencial + 1
                
        End Select
        
        'acumula valor do dia
        dTotalRec = dTotalRec + dValor
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
                
    objRelFluxoCx.dTotalRec = objRelFluxoCx.dTotalRec + dTotalRec
    objRelFluxoCx.dRecDespesas = objRelFluxoCx.dRecDespesas + dRecDespesas
    
    RelOpFlCx_Prepara1 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara1:

    RelOpFlCx_Prepara1 = gErr
     
    Select Case gErr
    
        Case 33333
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function RelOpFlCx_Prepara2(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'obter dados acumulados do dia referentes a pagamentos e gravar dados detalhados

Dim lErro As Long
Dim dValor As Double, dDespBanc As Double, dPagDesp As Double, dTotalPag As Double
Dim iFluxoCaixa As Integer, sNatureza As String, sObservacao As String, sNomeReduzido As String
Dim iNumParcelas As Integer, iNumParcela As Integer, lNumIntDoc As Long

On Error GoTo Erro_RelOpFlCx_Prepara2
    
    'obter dados de parcelaspag com status pendente ou previsao, da filialempresa
    
    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    sNomeReduzido = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT TitPagNatMov.NumParcelas, ParcelasPag.NumParcela, ParcelasPag.NumIntDoc, TitPagNatMov.FluxoCaixa, TitPagNatMov.Natureza, TitPagNatMov.Observacao, ParcelasPag.Saldo, Fornecedores.NomeReduzido FROM ParcelasPag, TitPagNatMov, Fornecedores WHERE TitPagNatMov.Fornecedor = Fornecedores.Codigo AND ParcelasPag.NumIntTitulo = TitPagNatMov.NumIntDoc AND TitPagNatMov.FilialEmpresa = ? AND (ParcelasPag.Status = ? OR ParcelasPag.Status = ?) AND ParcelasPag.DataVencimentoReal = ? ORDER BY TitPagNatMov.NumTitulo, ParcelasPag.NumParcela", _
        iNumParcelas, iNumParcela, lNumIntDoc, iFluxoCaixa, sNatureza, sObservacao, dValor, sNomeReduzido, iFilialEmpresa, STATUS_ABERTO, STATUS_PREVISAO, objRelFluxoCx.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        Select Case iFluxoCaixa
        
            Case 2 'despesas bancarias
                dDespBanc = dDespBanc + dValor
            
            Case 3 'outras
                dPagDesp = dPagDesp + dValor
                
            Case Else
                'insere registro
                lErro = Comando_Executar(alComando(4), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_PAGAMENTO, iFluxoCaixa, IIf(Len(Trim(sObservacao)) = 0, sNomeReduzido, sObservacao), dValor, iNumParcelas, iNumParcela, RELFLCX_PARC_PAG, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 33333
                            
                iSequencial = iSequencial + 1
                
        End Select
        
        'acumula valor do dia
        dTotalPag = dTotalPag + dValor
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
                
    objRelFluxoCx.dDespBanc = objRelFluxoCx.dDespBanc + dDespBanc
    objRelFluxoCx.dPagDesp = objRelFluxoCx.dPagDesp + dPagDesp
    objRelFluxoCx.dTotalPag = objRelFluxoCx.dTotalPag + dTotalPag
    
    RelOpFlCx_Prepara2 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara2:

    RelOpFlCx_Prepara2 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function
    
Function RelOpFlCx_Prepara3(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'obter dados acumulados do dia referentes a saques e depositos e gravar dados detalhados
    
Dim lErro As Long, iFluxoCaixa As Integer, sNatureza As String, sHistorico As String, dValor As Double, iTipo As Integer, lNumMovto As Long
Dim dDespBanc As Double, dPagDesp As Double, dTotalPag As Double
Dim dRecDespesas As Double, dTotalRec As Double

On Error GoTo Erro_RelOpFlCx_Prepara3
    
    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sHistorico = String(STRING_HISTORICOMOVCONTA, 0)
    
    lErro = Comando_Executar(alComando(6), "SELECT MovCCINatMov.NumMovto, MovCCINatMov.Tipo, MovCCINatMov.FluxoCaixa, MovCCINatMov.Natureza, MovCCINatMov.Historico, MovCCINatMov.Valor FROM MovCCINatMov WHERE MovCCINatMov.NumRefInterna = 0 AND MovCCINatMov.FilialEmpresa = ? AND MovCCINatMov.Excluido = 0 AND MovCCINatMov.DataMovimento = ? AND MovCCINatMov.Tipo IN (0,1)", _
        lNumMovto, iTipo, iFluxoCaixa, sNatureza, sHistorico, dValor, iFilialEmpresa, objRelFluxoCx.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        If iTipo = MOVCCI_SAQUE Then
        
            Select Case iFluxoCaixa
                
                Case 2 'despesas bancarias
                    dDespBanc = dDespBanc + dValor
                
                Case 3 'outras
                    dPagDesp = dPagDesp + dValor
                    
                Case Else
                    'insere registro
                    lErro = Comando_Executar(alComando(7), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                        objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_PAGAMENTO, iFluxoCaixa, sHistorico, dValor, 0, 0, RELFLCX_SAQUE, lNumMovto)
                    If lErro <> AD_SQL_SUCESSO Then gError 33333
                                
                    iSequencial = iSequencial + 1
                    
            End Select
            
            'acumula valor do dia
            dTotalPag = dTotalPag + dValor
        
        Else
        
            Select Case iFluxoCaixa
            
                Case 3 'outras
                    dRecDespesas = dRecDespesas + dValor
                    
                Case Else
                    'insere registro
                    lErro = Comando_Executar(alComando(8), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                        objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_RECEBIMENTO, iFluxoCaixa, sHistorico, dValor, 0, 0, RELFLCX_DEPOSITO, lNumMovto)
                    If lErro <> AD_SQL_SUCESSO Then gError 33333
                                
                    iSequencial = iSequencial + 1
                    
            End Select
            
            'acumula valor do dia
            dTotalRec = dTotalRec + dValor
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
                
    objRelFluxoCx.dTotalRec = objRelFluxoCx.dTotalRec + dTotalRec
    objRelFluxoCx.dRecDespesas = objRelFluxoCx.dRecDespesas + dRecDespesas
    
    objRelFluxoCx.dDespBanc = objRelFluxoCx.dDespBanc + dDespBanc
    objRelFluxoCx.dPagDesp = objRelFluxoCx.dPagDesp + dPagDesp
    objRelFluxoCx.dTotalPag = objRelFluxoCx.dTotalPag + dTotalPag
    
    RelOpFlCx_Prepara3 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara3:

    RelOpFlCx_Prepara3 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function RelOpFlCx_Prepara6(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'obter dados acumulados do dia referentes a baixasrec de titulos baixados ligadas a movtos de cta corrente na data e gravar dados detalhados

Dim lErro As Long
Dim dValor As Double, dRecDespesas As Double, dTotalRec As Double
Dim iFluxoCaixa As Integer, sNatureza As String, sNomeRedCli As String
Dim iNumParcelas As Integer, iNumParcela As Integer, lNumIntDoc As Long

On Error GoTo Erro_RelOpFlCx_Prepara6
    
    dRecDespesas = 0
    dTotalRec = 0
    
    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sNomeRedCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(9), "SELECT TitRecBxNatMov.NumParcelas, ParcelasRecBaixadas.NumParcela, BaixasParcRec.NumIntDoc, TitRecBxNatMov.FluxoCaixa, TitRecBxNatMov.Natureza, Clientes.NomeReduzido, BaixasParcRec.ValorRecebido FROM MovimentosContaCorrente, BaixasRec, BaixasParcRec, ParcelasRecBaixadas, TitRecBxNatMov, Clientes WHERE MovimentosContaCorrente.NumMovto = BaixasRec.NumMovCta AND BaixasRec.NumIntBaixa = BaixasParcRec.NumIntBaixa AND BaixasParcRec.Status = ? AND BaixasParcRec.NumIntParcela = ParcelasRecBaixadas.NumIntDoc AND ParcelasRecBaixadas.NumIntTitulo = TitRecBxNatMov.NumIntDoc AND TitRecBxNatMov.Cliente = Clientes.Codigo AND TitRecBxNatMov.FilialEmpresa = ? AND MovimentosContaCorrente.DataMovimento = ?", _
        iNumParcelas, iNumParcela, lNumIntDoc, iFluxoCaixa, sNatureza, sNomeRedCli, dValor, iFilialEmpresa, STATUS_LANCADO, objRelFluxoCx.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        Select Case iFluxoCaixa
        
            Case 3 'outras
                dRecDespesas = dRecDespesas + dValor
                
            Case Else
                'insere registro
                lErro = Comando_Executar(alComando(10), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_RECEBIMENTO, iFluxoCaixa, sNomeRedCli, dValor, iNumParcelas, iNumParcela, RELFLCX_BX_PARC_REC, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 33333
                            
                iSequencial = iSequencial + 1
                
        End Select
        
        'acumula valor do dia
        dTotalRec = dTotalRec + dValor
        
        lErro = Comando_BuscarProximo(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
                
    objRelFluxoCx.dTotalRec = objRelFluxoCx.dTotalRec + dTotalRec
    objRelFluxoCx.dRecDespesas = objRelFluxoCx.dRecDespesas + dRecDespesas
    
    RelOpFlCx_Prepara6 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara6:

    RelOpFlCx_Prepara6 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function RelOpFlCx_Prepara7(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'obter dados acumulados do dia referentes a baixasrec de titulos abertos ligadas a movtos de cta corrente na data e gravar dados detalhados

Dim lErro As Long
Dim dValor As Double, dRecDespesas As Double, dTotalRec As Double
Dim iFluxoCaixa As Integer, sNatureza As String, sNomeRedCli As String
Dim iNumParcelas As Integer, iNumParcela As Integer, lNumIntDoc As Long

On Error GoTo Erro_RelOpFlCx_Prepara7
    
    dRecDespesas = 0
    dTotalRec = 0
    
    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sNomeRedCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(11), "SELECT TitRecNatMov.NumParcelas, ParcelasRec.NumParcela, BaixasParcRec.NumIntDoc, TitRecNatMov.FluxoCaixa, TitRecNatMov.Natureza, Clientes.NomeReduzido, BaixasParcRec.ValorRecebido FROM MovimentosContaCorrente, BaixasRec, BaixasParcRec, ParcelasRec, TitRecNatMov, Clientes WHERE MovimentosContaCorrente.NumMovto = BaixasRec.NumMovCta AND BaixasRec.NumIntBaixa = BaixasParcRec.NumIntBaixa AND BaixasParcRec.Status = ? AND BaixasParcRec.NumIntParcela = ParcelasRec.NumIntDoc AND ParcelasRec.NumIntTitulo = TitRecNatMov.NumIntDoc AND TitRecNatMov.Cliente = Clientes.Codigo AND TitRecNatMov.FilialEmpresa = ? AND MovimentosContaCorrente.DataMovimento = ?", _
        iNumParcelas, iNumParcela, lNumIntDoc, iFluxoCaixa, sNatureza, sNomeRedCli, dValor, iFilialEmpresa, STATUS_LANCADO, objRelFluxoCx.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(11))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        Select Case iFluxoCaixa
        
            Case 3 'outras
                dRecDespesas = dRecDespesas + dValor
                
            Case Else
                'insere registro
                lErro = Comando_Executar(alComando(12), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_RECEBIMENTO, iFluxoCaixa, sNomeRedCli, dValor, iNumParcelas, iNumParcela, RELFLCX_BX_PARC_REC, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 33333
                            
                iSequencial = iSequencial + 1
                
        End Select
        
        'acumula valor do dia
        dTotalRec = dTotalRec + dValor
        
        lErro = Comando_BuscarProximo(alComando(11))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
                
    objRelFluxoCx.dTotalRec = objRelFluxoCx.dTotalRec + dTotalRec
    objRelFluxoCx.dRecDespesas = objRelFluxoCx.dRecDespesas + dRecDespesas
    
    RelOpFlCx_Prepara7 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara7:

    RelOpFlCx_Prepara7 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function RelOpFlCx_Prepara8(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'obter dados acumulados do dia referentes a baixaspag de titulos baixados ligadas a movtos de cta corrente na data e gravar dados detalhados

Dim lErro As Long
Dim dValor As Double, dDespBanc As Double, dPagDesp As Double, dTotalPag As Double
Dim iFluxoCaixa As Integer, sNatureza As String, sObservacao As String, sNomeReduzido As String
Dim iNumParcelas As Integer, iNumParcela As Integer, lNumIntDoc As Long

On Error GoTo Erro_RelOpFlCx_Prepara8
    
    'obter dados de parcelaspag com status pendente ou previsao, da filialempresa
    
    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    sNomeReduzido = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(13), "SELECT TitPagBxNatMov.NumParcelas, ParcelasPagBaixadas.NumParcela, BaixasParcPag.NumIntDoc, TitPagBxNatMov.FluxoCaixa, TitPagBxNatMov.Natureza, TitPagBxNatMov.Observacao, (BaixasParcPag.ValorBaixado+BaixasParcPag.ValorMulta+BaixasParcPag.ValorJuros-BaixasParcPag.ValorDesconto), Fornecedores.NomeReduzido FROM MovimentosContaCorrente, BaixasPag, BaixasParcPag, ParcelasPagBaixadas, TitPagBxNatMov, Fornecedores WHERE TitPagBxNatMov.Fornecedor = Fornecedores.Codigo AND MovimentosContaCorrente.NumMovto = BaixasPag.NumMovCta AND BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.Status = ? AND BaixasParcPag.NumIntParcela = ParcelasPagBaixadas.NumIntDoc AND ParcelasPagBaixadas.NumIntTitulo = TitPagBxNatMov.NumIntDoc AND TitPagBxNatMov.FilialEmpresa = ? AND MovimentosContaCorrente.DataMovimento = ?", _
        iNumParcelas, iNumParcela, lNumIntDoc, iFluxoCaixa, sNatureza, sObservacao, dValor, sNomeReduzido, iFilialEmpresa, STATUS_LANCADO, objRelFluxoCx.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(13))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        Select Case iFluxoCaixa
        
            Case 2 'despesas bancarias
                dDespBanc = dDespBanc + dValor
            
            Case 3 'outras
                dPagDesp = dPagDesp + dValor
                
            Case Else
                'insere registro
                lErro = Comando_Executar(alComando(14), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_PAGAMENTO, iFluxoCaixa, IIf(Len(Trim(sObservacao)) = 0, sNomeReduzido, sObservacao), dValor, iNumParcelas, iNumParcela, RELFLCX_BX_PARC_PAG, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 33333
                            
                iSequencial = iSequencial + 1
                
        End Select
        
        'acumula valor do dia
        dTotalPag = dTotalPag + dValor
        
        lErro = Comando_BuscarProximo(alComando(13))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
                
    objRelFluxoCx.dDespBanc = objRelFluxoCx.dDespBanc + dDespBanc
    objRelFluxoCx.dPagDesp = objRelFluxoCx.dPagDesp + dPagDesp
    objRelFluxoCx.dTotalPag = objRelFluxoCx.dTotalPag + dTotalPag
    
    RelOpFlCx_Prepara8 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara8:

    RelOpFlCx_Prepara8 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function
    
Function RelOpFlCx_Prepara9(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'obter dados acumulados do dia referentes a baixaspag de titulos abertos ligadas a movtos de cta corrente na data e gravar dados detalhados

Dim lErro As Long, sNomeReduzido As String
Dim dValor As Double, dDespBanc As Double, dPagDesp As Double, dTotalPag As Double
Dim iFluxoCaixa As Integer, sNatureza As String, sObservacao As String
Dim iNumParcelas As Integer, iNumParcela As Integer, lNumIntDoc As Long

On Error GoTo Erro_RelOpFlCx_Prepara9
    
    'obter dados de parcelaspag com status pendente ou previsao, da filialempresa
    
    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sObservacao = String(STRING_TITULO_OBSERVACAO, 0)
    sNomeReduzido = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(15), "SELECT TitPagNatMov.NumParcelas, ParcelasPag.NumParcela, BaixasParcPag.NumIntDoc, TitPagNatMov.FluxoCaixa, TitPagNatMov.Natureza, TitPagNatMov.Observacao, (BaixasParcPag.ValorBaixado+BaixasParcPag.ValorMulta+BaixasParcPag.ValorJuros-BaixasParcPag.ValorDesconto), Fornecedores.NomeReduzido FROM MovimentosContaCorrente, BaixasPag, BaixasParcPag, ParcelasPag, TitPagNatMov, Fornecedores WHERE TitPagNatMov.Fornecedor = Fornecedores.Codigo AND MovimentosContaCorrente.NumMovto = BaixasPag.NumMovCta AND BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.Status = ? AND BaixasParcPag.NumIntParcela = ParcelasPag.NumIntDoc AND ParcelasPag.NumIntTitulo = TitPagNatMov.NumIntDoc AND TitPagNatMov.FilialEmpresa = ? AND MovimentosContaCorrente.DataMovimento = ?", _
        iNumParcelas, iNumParcela, lNumIntDoc, iFluxoCaixa, sNatureza, sObservacao, dValor, sNomeReduzido, iFilialEmpresa, STATUS_LANCADO, objRelFluxoCx.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(15))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        Select Case iFluxoCaixa
        
            Case 2 'despesas bancarias
                dDespBanc = dDespBanc + dValor
            
            Case 3 'outras
                dPagDesp = dPagDesp + dValor
                
            Case Else
                'insere registro
                lErro = Comando_Executar(alComando(16), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_PAGAMENTO, iFluxoCaixa, IIf(Len(Trim(sObservacao)) = 0, sNomeReduzido, sObservacao), dValor, iNumParcelas, iNumParcela, RELFLCX_BX_PARC_PAG, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 33333
                            
                iSequencial = iSequencial + 1
                
        End Select
        
        'acumula valor do dia
        dTotalPag = dTotalPag + dValor
        
        lErro = Comando_BuscarProximo(alComando(15))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
                
    objRelFluxoCx.dDespBanc = objRelFluxoCx.dDespBanc + dDespBanc
    objRelFluxoCx.dPagDesp = objRelFluxoCx.dPagDesp + dPagDesp
    objRelFluxoCx.dTotalPag = objRelFluxoCx.dTotalPag + dTotalPag
    
    RelOpFlCx_Prepara9 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara9:

    RelOpFlCx_Prepara9 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function
    
Function RelOpFlCx_Prepara4(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'preencher dados de aplicacoes com movto no dia corrente

Dim lErro As Long, dTotalAplicado As Double
Dim sNatureza As String, sHistorico As String, lNumMovto As Long, iTipo As Integer, iFluxoCaixa As Integer, dValor As Double

On Error GoTo Erro_RelOpFlCx_Prepara4

    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sHistorico = String(STRING_HISTORICOMOVCONTA, 0)
    
    lErro = Comando_Executar(alComando(17), "SELECT MovCCINatMov.NumMovto, MovCCINatMov.Tipo, MovCCINatMov.FluxoCaixa, MovCCINatMov.Historico, MovCCINatMov.Valor FROM MovCCINatMov WHERE MovCCINatMov.FilialEmpresa = ? AND MovCCINatMov.Excluido = 0 AND MovCCINatMov.DataMovimento = ? AND MovCCINatMov.Tipo = ?", _
        lNumMovto, iTipo, iFluxoCaixa, sHistorico, dValor, iFilialEmpresa, objRelFluxoCx.dtData, MOVCCI_APLICACAO)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(17))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        'insere registro
        lErro = Comando_Executar(alComando(18), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
            objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_PAGAMENTO, iFluxoCaixa, sHistorico, dValor, 0, 0, RELFLCX_APLICACAO, lNumMovto)
        If lErro <> AD_SQL_SUCESSO Then gError 33333
                    
        iSequencial = iSequencial + 1
    
        dTotalAplicado = dTotalAplicado + dValor
        
        lErro = Comando_BuscarProximo(alComando(17))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
        
    objRelFluxoCx.dTotalAplic = objRelFluxoCx.dTotalAplic + dTotalAplicado
    objRelFluxoCx.dTotalPag = objRelFluxoCx.dTotalPag + dTotalAplicado
    
    RelOpFlCx_Prepara4 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara4:

    RelOpFlCx_Prepara4 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function RelOpFlCx_Prepara5(ByVal objRelFluxoCx As ClassRelFluxoCx, ByVal iFilialEmpresa As Integer, iSequencial As Integer, alComando() As Long) As Long
'preencher dados de resgates com movto no dia corrente

Dim lErro As Long, dTotalResgatado As Double
Dim sNatureza As String, sHistorico As String, lNumMovto As Long, iTipo As Integer, iFluxoCaixa As Integer, dValor As Double

On Error GoTo Erro_RelOpFlCx_Prepara5

    sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    sHistorico = String(STRING_HISTORICOMOVCONTA, 0)
    
    lErro = Comando_Executar(alComando(19), "SELECT MovCCINatMov.NumMovto, MovCCINatMov.Tipo, MovCCINatMov.FluxoCaixa, MovCCINatMov.Historico, MovCCINatMov.Valor FROM MovCCINatMov WHERE MovCCINatMov.FilialEmpresa = ? AND MovCCINatMov.Excluido = 0 AND MovCCINatMov.DataMovimento = ? AND MovCCINatMov.Tipo = ?", _
        lNumMovto, iTipo, iFluxoCaixa, sHistorico, dValor, iFilialEmpresa, objRelFluxoCx.dtData, MOVCCI_RESGATE)
    If lErro <> AD_SQL_SUCESSO Then gError 33333
    
    lErro = Comando_BuscarProximo(alComando(19))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Do While lErro = AD_SQL_SUCESSO

        dValor = Round(dValor, 0)
        
        'insere registro
        lErro = Comando_Executar(alComando(20), "INSERT INTO RelOpFluxoCxItemMiguez (NumIntRel, Data, Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
            objRelFluxoCx.lNumIntRel, objRelFluxoCx.dtData, iSequencial, NATUREZA_TIPO_RECEBIMENTO, iFluxoCaixa, sHistorico, dValor, 0, 0, RELFLCX_APLICACAO, lNumMovto)
        If lErro <> AD_SQL_SUCESSO Then gError 33333
                    
        iSequencial = iSequencial + 1
        
        dTotalResgatado = dTotalResgatado + dValor
    
        lErro = Comando_BuscarProximo(alComando(19))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

    Loop
        
    objRelFluxoCx.dTotalAplic = objRelFluxoCx.dTotalAplic - dTotalResgatado
    objRelFluxoCx.dTotalRec = objRelFluxoCx.dTotalRec + dTotalResgatado
    
    RelOpFlCx_Prepara5 = SUCESSO
     
    Exit Function
    
Erro_RelOpFlCx_Prepara5:

    RelOpFlCx_Prepara5 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

