VERSION 5.00
Begin VB.UserControl CancelaCupom 
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ScaleHeight     =   2955
   ScaleWidth      =   4695
   Begin VB.CommandButton BotaoSelecionarVenda 
      Caption         =   "Selecionar Venda"
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
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   1830
   End
   Begin VB.TextBox DescricaoVenda 
      Height          =   1935
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   735
      Width           =   4410
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3375
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "CancelaCupom.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Enabled         =   0   'False
         Height          =   360
         Left            =   75
         Picture         =   "CancelaCupom.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "CancelaCupom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private gobjVenda As ClassVenda
Private giIndice As Integer

Private Function Alteracoes_CancelamentoCupom(objVenda As ClassVenda) As Long

Dim objMovCaixa As ClassMovimentoCaixa
Dim objCheque As ClassChequePre
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim objCarne As ClassCarne
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lSequencial As Long
Dim objAliquota As New ClassAliquotaICMS
Dim objItens As ClassItemCupomFiscal
Dim iIndice1 As Integer
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_Alteracoes_CancelamentoCupom

    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o Movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'se for um recebimento em cartão de crédito/Debito de TEF
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iTipoCartao = TIPO_TEF Then
            '''?????efetua caneclamento de TEF
        End If
    Next

    For Each objItens In objVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja - ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next

    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1
        Set objMovCaixa = gcolMovimentosCaixa.Item(iIndice)
        If objMovCaixa.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolMovimentosCaixa.Remove (iIndice)
    Next

    'Para cada movimento da venda
    For Each objMovCaixa In objVenda.colMovimentosCaixa

'??? 24/08/2016         If objMovCaixa.iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor

'??? 24/08/2016         If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor

        'Se for de cartao de crédito ou débito especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto <> 0 Then
            'Busca em gcolCartão a ocorrencia de Cartão nao especificado
            For iIndice = gcolCartao.Count To 1 Step -1
                Set objAdmMeioPagtoCondPagto = gcolCartao.Item(iIndice)
                'Se encontrou
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovCaixa.iTipoCartao Then
                    'Atualiza o saldo do cartão
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolCartao.Remove (iIndice)
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de crédito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de débito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
    Next

    'Para cada movimento
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for um recebimento em ticket
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada obj de ticket da coleção global de tickets
                For Each objAdmMeioPagtoCondPagto In gcolTicket
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo de não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Ticket da coleção global
                For iIndice1 = gcolTicket.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolTicket.Item(iIndice1)
                    'Se encontrou o ticket/parcelamento
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolTicket.Remove (iIndice1)
                        'Sinaliza que encontrou
                        Exit For
                    End If
                Next
            End If
        End If
    Next

    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto

    'Verifica se já existe movimentos de Outros\
    'Para cada MOvimento de Outros
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o MOvimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for do tipo outros
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada pagamento em outros na coleção global
                For Each objAdmMeioPagtoCondPagto In gcolOutros
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Pagamento em outros na col global
                For iIndice1 = gcolOutros.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolOutros.Item(iIndice1)
                    'Se for do mesmo tipo que o atual
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolOutros.Remove (iIndice1)
                        Exit For
                    End If
                Next
            End If
        End If
    Next

    'remove o Carne na col global
    If objVenda.objCarne.colParcelas.Count > 0 Then
        For iIndice = 1 To gcolCarne.Count
            Set objCarne = gcolCarne.Item(iIndice)
            If objCarne.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCarne.Remove (iIndice)
        Next
    End If

    'remove o Cheque na col global
    If objVenda.colCheques.Count > 0 Then
        For iIndice = gcolCheque.Count To 1 Step -1
            Set objCheque = gcolCheque.Item(iIndice)
            If objCheque.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCheque.Remove (iIndice)
        Next
    End If

    Alteracoes_CancelamentoCupom = SUCESSO

    Exit Function

Erro_Alteracoes_CancelamentoCupom:

    Alteracoes_CancelamentoCupom = gErr

    Select Case gErr

        Case 99901, 99953, 99952

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175678)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim lSequencial As Long
Dim lIntervaloTrans As Long
Dim sRetorno As String
Dim lTamanho As Long
Dim objObject As Object
Dim vbMsgRes As VbMsgBoxResult
Dim objOperador As New ClassOperador
Dim iCodGerente As Integer
Dim objMovCaixa As ClassMovimentoCaixa
Dim objMovCaixa1 As ClassMovimentoCaixa
Dim iCuponsVinculados As Integer
Dim colMeiosPag As New Collection

On Error GoTo Erro_Gravar_Registro

    'Envia aviso perguntando se deseja cancelar o cupom
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELA_CUPOM_TELA)

    If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

    'Se for Necessário a autorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM

        'Se Operador for Gerente
        iCodGerente = objOperador.iCodigo

    End If

    'verifica se tem mais de um cupom vinculado impresso
    'se tiver ==> nao pode cancelar por limitacao do ecf
    For Each objMovCaixa In gobjVenda.colMovimentosCaixa

        lErro = CF_ECF("Trata_MovCaixa", objMovCaixa, colMeiosPag)
        If lErro <> SUCESSO Then gError 214733

    Next

    iCuponsVinculados = 0

    For Each objMovCaixa1 In colMeiosPag

        If objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Or objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO Or objMovCaixa1.iTipo = TIPOMEIOPAGTOLOJA_TEF Then
            iCuponsVinculados = iCuponsVinculados + 1
        End If

    Next

    If iCuponsVinculados > 1 Then gError 214732

    'cancelar o Cupom de Venda
    lErro = AFRAC_CancelarCupom(Me, gobjVenda)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
    If lErro <> SUCESSO Then gError 99610

    'Fecha a Transação
    lErro = CF_ECF("Grava_CancelamentoCupom_Arquivo", gobjVenda)
    If lErro <> SUCESSO Then gError 112421

    lErro = Alteracoes_CancelamentoCupom(gobjVenda)
    If lErro <> SUCESSO Then gError 112078

    gcolVendas.Remove (giIndice)

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163636)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    Call Rotina_AvisoECF(vbOKOnly, AVISO_CANCELA_CUPOM_TELA_SUCESSO)
    
    Unload Me 'Não pode deixar na tela para cancelar mais de uma vez o mesmo cupom

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 109485, 207981

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163637)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    Call BotaoSelecionarVenda_Click

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163638)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163639)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Cancelar Cupom"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CancelaCupom"

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

Private Sub BotaoSelecionarVenda_Click()

Dim objCupomSelecionado As New ClassCupomFiscal
Dim lErro As Long
Dim objCupom As ClassCupomFiscal
Dim iIndice As Integer
Dim objVenda As ClassVenda

    Set gobjVenda = Nothing
    DescricaoVenda.Text = ""
    BotaoGravar.Enabled = False

    'chama o browser cupomFiscalLista passando objcupom
    Call Chama_TelaECF_Modal("CupomFiscalLista", objCupomSelecionado)

    'se objCupom tiver preenchido, chama pra tela..
    If giRetornoTela = vbOK Then

        For iIndice = gcolVendas.Count To 1 Step -1
    
            Set objVenda = gcolVendas.Item(iIndice)
    
            Set objCupom = objVenda.objCupomFiscal
    
            If ((objVenda.iTipo = OPTION_CF And objCupomSelecionado.lNumero = objCupom.lNumero) Or (objVenda.iTipo <> OPTION_CF And objCupomSelecionado.lNumero = objCupom.lNumOrcamento)) And Abs(objCupomSelecionado.dHoraEmissao - objCupom.dHoraEmissao) < 0.00001 And objCupomSelecionado.dtDataEmissao = objCupom.dtDataEmissao And objCupomSelecionado.dValorTotal = objCupom.dValorTotal Then
            
                Select Case objVenda.iTipo
        
                    Case OPTION_CF
        
                        If objCupom.iStatus = 0 Then
        
                            DescricaoVenda.Text = "O cancelamento de nfce deve ser feito na tela de venda."
        '                    permitir cancelar nfce emitida na ultima meia hora
        '                    Set gobjVenda = objVenda
        '                    Exit For
        
                        End If
        
        
                    Case OPTION_DAV, OPTION_ORCAMENTO
        
                        'se foi um orçamento que abriu gaveta
                        If objCupom.iStatus = 2 Then
        
                            Set gobjVenda = objVenda
                            Exit For
        
                        End If
        
                End Select
    
            End If
    
        Next

        If Not (gobjVenda Is Nothing) Then
    
            giIndice = iIndice
    
            'preencher o controle DescricaoVenda
            DescricaoVenda.Text = IIf(gobjVenda.iTipo = OPTION_CF, "NFCE: " & CStr(objCupom.lNumero), "DAV: " & CStr(objCupom.lNumOrcamento)) & " Data: " & Format(objCupom.dtDataEmissao, "dd/mm/yy") & " Hora: " & Format(objCupom.dHoraEmissao, "hh:mm:ss") & " Valor: R$ " & Format(objCupom.dValorTotal, "standard")
    
            BotaoGravar.Enabled = True
    
        End If

    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atravez da Tecla F2
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click

        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click

    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163640)

    End Select

    Exit Sub

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****


