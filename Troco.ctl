VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl Troco 
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   KeyPreview      =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   4755
   Begin VB.Frame Frame1 
      Caption         =   "Composição Troco"
      Height          =   4440
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4575
      Begin VB.CommandButton BotaoTicket 
         Caption         =   "(F2) ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3765
         TabIndex        =   14
         Top             =   960
         Width           =   765
      End
      Begin VB.CommandButton BotaoOk 
         Caption         =   "(F5)   Ok"
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
         Left            =   1590
         TabIndex        =   13
         Top             =   3915
         Width           =   1515
      End
      Begin MSMask.MaskEdBox Dinheiro 
         Height          =   360
         Left            =   1995
         TabIndex        =   1
         Top             =   435
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Ticket 
         Height          =   360
         Left            =   1995
         TabIndex        =   2
         Top             =   945
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ContraVale 
         Height          =   360
         Left            =   1995
         TabIndex        =   3
         Top             =   1500
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Falta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1995
         TabIndex        =   5
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Falta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   1245
         TabIndex        =   12
         Top             =   2670
         Width           =   705
      End
      Begin VB.Label SubTotal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1995
         TabIndex        =   4
         Top             =   2085
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SubTotal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   780
         TabIndex        =   11
         Top             =   2115
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contra Vale:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   435
         TabIndex        =   10
         Top             =   1515
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ticket:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1140
         TabIndex        =   9
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dinheiro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   855
         TabIndex        =   8
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label TrocoTotal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1995
         TabIndex        =   7
         Top             =   3195
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Troco Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   510
         TabIndex        =   6
         Top             =   3225
         Width           =   1440
      End
   End
End
Attribute VB_Name = "Troco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda
Dim gdValorTicketAnterior As Double

Function Trata_Parametros(objVenda As ClassVenda) As Long
    
Dim objMovimento As ClassMovimentoCaixa
Dim iTipo As Integer

    Set gobjVenda = objVenda
        
    'Os movimentos de caixa referentes a troco e colocar na tela
    'Joga na tela todos os Tickets
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_DINHEIRO
        End If
        'Se for do tipo dinheiro
        If objMovimento.iTipo = iTipo Then
            Dinheiro.Text = objMovimento.dValor
        End If
        
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_VALE
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_TICKET
        End If
        'Se for do tipo Ticket
        If objMovimento.iTipo = iTipo Then
            Ticket.Text = StrParaDbl(Ticket.Text) + objMovimento.dValor
        End If
        
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_CONTRAVALE
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_CONTRAVALE
        End If
        'Se for do tipo ContraVale
        If objMovimento.iTipo = iTipo Then
            ContraVale.Text = StrParaDbl(ContraVale.Text) + objMovimento.dValor
        End If
        
    Next
    
    'Joga o total do troco na tela
    TrocoTotal.Caption = Format(gobjVenda.objCupomFiscal.dValorTroco, "standard")
    
    'Atualiza o total do troco
    Call Recalcula_Valores
    
    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()
    
    giRetornoTela = vbOK
    
    lErro_Chama_Tela = SUCESSO

End Sub

Private Sub Dinheiro_LostFocus()

    If Dinheiro.Text = "" Then Dinheiro.Text = "0,00"
    
End Sub

Private Sub Dinheiro_GotFocus()
    
    If Dinheiro.Text = "0,00" Then Dinheiro.Text = ""
        
    If right(Dinheiro.Text, 3) = ",00" Then Dinheiro.Text = Format(Dinheiro.Text, "#,#")
    
    'Posiciona o cursor na frente
    Call MaskEdBox_TrataGotFocus(Dinheiro)
    
End Sub

Private Sub Dinheiro_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim iIndice As Integer
Dim objMovimento As ClassMovimentoCaixa
Dim iTipo As Integer

On Error GoTo Erro_Dinheiro_Validate
    
    If StrParaDbl(Dinheiro.Text) > 0 Then
    
        lErro = Valor_Positivo_Critica(Dinheiro.Text)
        If lErro <> SUCESSO Then gError 99628
        
    ElseIf StrParaDbl(Dinheiro.Text) < 0 Then
        gError 215030
    Else
        'Exclui todos os movimentos em dinheiro
        For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
            Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
            If gsNomePrinc = "SGEECF" Then
                iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO
            Else
                iTipo = MOVIMENTOCAIXA_CARNE_TROCO_DINHEIRO
            End If
            
            If objMovimento.iTipo = iTipo Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
        Next
    
    End If
        
    Call Recalcula_Valores
    
    Exit Sub
    
Erro_Dinheiro_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99628
        
        Case 215030
            Call Rotina_ErroECF(vbOKOnly, ERRO_TROCO_DINHEIRO_NEGATIVO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175615)

    End Select

    Exit Sub
    
End Sub

Private Sub Ticket_LostFocus()

    If Ticket.Text = "" Then Ticket.Text = "0,00"
    
End Sub

Private Sub Ticket_GotFocus()
    
    If Ticket.Text = "0,00" Then Ticket.Text = ""
        
    If right(Ticket.Text, 3) = ",00" Then Ticket.Text = Format(Ticket.Text, "#,#")
    
    'Posiciona o cursor na frente
    Call MaskEdBox_TrataGotFocus(Ticket)
    
    'Guarda o valor presente no campo Ticket
    gdValorTicketAnterior = StrParaDbl(Ticket.Text)
    
End Sub

Private Sub Ticket_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim objMovimento As ClassMovimentoCaixa
Dim vbMsgRes As VbMsgBoxResult
Dim bContinua As Boolean
Dim dValor As Double
Dim dTicket_Esp As Double

On Error GoTo Erro_Ticket_Validate
    
    If StrParaDbl(Ticket.Text) > 0 Then
    
        lErro = Valor_Positivo_Critica(Ticket.Text)
        If lErro <> SUCESSO Then gError 99629
        
    End If
       
    If StrParaDbl(Ticket.Text) < 0 Then
        gError 215032
    End If
       
    'Verifica se alterou a quantidade de ticket
    If gdValorTicketAnterior <> StrParaDbl(Ticket.Text) Then
        'Se alterou para mais
        If StrParaDbl(Ticket.Text) - gdValorTicketAnterior > 0.0001 Then
            dValor = StrParaDbl(Ticket.Text) - gdValorTicketAnterior
            'verificar se existe e update acrescentando senão --> cria novo
            Call Adiciona_Ticket(dValor, True)
        'Se alterou para menos
        Else
            'retorna o somatório do Tickets especificados
            Call Soma_Ticket_Especificado(dTicket_Esp)
            
            'Se o total dos tickets espec. for maior que o total atual
            If StrParaDbl(Ticket.Text) < dTicket_Esp Then
                'Envia aviso perguntando se realmente deseja apagar os dados do ticket
                vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELAR_TICKET)
        
                If vbMsgRes = vbYes Then
                    dValor = StrParaDbl(Ticket.Text)
                    'cancela os tickets
                    Call Cancela_Ticket
                    'verificar se existe senão cria novo
                    Call Adiciona_Ticket(dValor, False)
                Else
                    gError 99631
                End If
            Else
                dValor = StrParaDbl(Ticket.Text) - dTicket_Esp
                'verificar se existe senão cria novo
                Call Adiciona_Ticket(dValor, False)
            End If
         
        End If
    End If

    Call Recalcula_Valores
        
    Exit Sub
    
Erro_Ticket_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99629, 99631
        
        Case 215032
            Call Rotina_ErroECF(vbOKOnly, ERRO_TROCO_TICKET_NEGATIVO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175616)

    End Select

    Exit Sub
    
End Sub

Private Sub Adiciona_Ticket(dValor As Double, bAdd As Boolean)
    
Dim objMovimento As New ClassMovimentoCaixa
Dim bAchou As Boolean
Dim iTipo As Integer

    bAchou = False
    
    'Verifica se acha o registro
    For Each objMovimento In gobjVenda.colMovimentosCaixa
    'se achou --> update
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_VALE
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_TICKET
        End If
        
        If objMovimento.iTipo = iTipo And objMovimento.iAdmMeioPagto = 0 Then
            'Se é para ser adicionado
            If bAdd Then
                objMovimento.dValor = objMovimento.dValor + dValor
            'Se é para ser substituído
            Else
                objMovimento.dValor = dValor
            End If
            bAchou = True
         End If
    Next
    
    'Se não achou um campo --> cria novo
    If Not (bAchou) Then
        Set objMovimento = New ClassMovimentoCaixa

        'Insere um novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iTipo = iTipo
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = dValor
        objMovimento.dHora = CDbl(Time)
        objMovimento.iParcelamento = COD_A_VISTA
    
        gobjVenda.colMovimentosCaixa.Add objMovimento
    End If
    
    Exit Sub
    
End Sub

Private Sub Soma_Ticket_Especificado(dValor As Double)
    
Dim objMovimento As New ClassMovimentoCaixa
Dim iTipo As Integer

    'Somar o total de troco em ticket especificado
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_VALE
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_TICKET
        End If
        
        If objMovimento.iTipo = iTipo And objMovimento.iAdmMeioPagto <> 0 Then dValor = dValor + objMovimento.dValor
    Next
    
End Sub

Private Sub Cancela_Ticket()
    
Dim objMovimento As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim iTipo As Integer

    'exclui todos os tickets especificados
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_VALE
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_TICKET
        End If
        
        If objMovimento.iTipo = iTipo And objMovimento.iAdmMeioPagto <> 0 Then
            gobjVenda.colMovimentosCaixa.Remove (iIndice)
        End If
    Next
    
End Sub

Private Sub ContraVale_LostFocus()

    If ContraVale.Text = "" Then ContraVale.Text = "0,00"
    
End Sub

Private Sub ContraVale_GotFocus()
    
    If ContraVale.Text = "0,00" Then ContraVale.Text = ""
        
    If right(ContraVale.Text, 3) = ",00" Then ContraVale.Text = Format(ContraVale.Text, "#,#")
    
    'Posiciona o cursor na frente
    Call MaskEdBox_TrataGotFocus(ContraVale)
    
End Sub

Private Sub ContraVale_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim iIndice As Integer
Dim objMovimento As ClassMovimentoCaixa
Dim iTipo As Integer

On Error GoTo Erro_ContraVale_Validate
    
    If StrParaDbl(ContraVale.Text) > 0 Then
    
        lErro = Valor_Positivo_Critica(ContraVale.Text)
        If lErro <> SUCESSO Then gError 99630
        
    ElseIf StrParaDbl(ContraVale.Text) < 0 Then
        gError 215031
        
    Else
        
        'Exclui todos os movimentos em ContraVale
        For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
            Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
            If gsNomePrinc = "SGEECF" Then
                iTipo = MOVIMENTOCAIXA_TROCO_CONTRAVALE
            Else
                iTipo = MOVIMENTOCAIXA_CARNE_TROCO_CONTRAVALE
            End If
            
            If objMovimento.iTipo = iTipo Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
        Next
    
    End If
        
    Call Recalcula_Valores
        
    Exit Sub
    
Erro_ContraVale_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99630
        
        Case 215031
            Call Rotina_ErroECF(vbOKOnly, ERRO_TROCO_CONTRAVALE_NEGATIVO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175617)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoTicket_Click()
    
Dim iIndice As Integer
Dim objMovimento As ClassMovimentoCaixa
Dim dTotal As Double
Dim iTipo As Integer

    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(TrocoTotal.Caption)
    
    Call Chama_TelaECF_Modal("TrocoTicket", gobjVenda)
    
    If gsNomePrinc = "SGEECF" Then
        iTipo = MOVIMENTOCAIXA_TROCO_VALE
    Else
        iTipo = MOVIMENTOCAIXA_CARNE_TROCO_TICKET
    End If
    
    'Somar o total de troco em ticket
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = iTipo Then dTotal = dTotal + objMovimento.dValor
    Next
       
    'COLOCA NA TELA
    Ticket.Text = Format(dTotal, "standard")
    
    'Se ultrapassou o valor total
    If (StrParaDbl(Dinheiro.Text) + StrParaDbl(Ticket.Text) + StrParaDbl(ContraVale.Text)) - StrParaDbl(TrocoTotal.Caption) > 0.0001 Then
        'Apaga o valor de contra-vale
        ContraVale.Text = ""
        'o que faltar para completar o troco coloca em dinheiro
        Dinheiro.Text = Format(StrParaDbl(TrocoTotal.Caption) - StrParaDbl(Ticket.Text), "standard")
    End If
    
    Call Recalcula_Valores
    
End Sub

Private Sub BotaoOK_Click()
'Não Inclui os movimentos referentes ao Ticket pois já estão especificados
Dim iTipo As Integer
Dim dDinheiro As Double
Dim dVale As Double
Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207999
    
    If Not gobjVenda Is Nothing Then
    
    'Se o valor do troco difere do informado
    If StrParaDbl(Subtotal.Caption) <> StrParaDbl(TrocoTotal.Caption) Then gError 99689
    
    Call Dinheiro_Validate(False)
    Call ContraVale_Validate(False)
        
    If Len(Trim(Dinheiro.Text)) <> 0 Then dDinheiro = StrParaDbl(Dinheiro.Text)
    If Len(Trim(ContraVale.Text)) <> 0 Then dVale = StrParaDbl(ContraVale.Text)
    
    'Inclui os movimentos referentes ao dinheiro
    If dDinheiro > 0 Then
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_DINHEIRO
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_DINHEIRO
        End If
        Call Inclui_Movimento(StrParaDbl(Dinheiro.Text), MEIO_PAGAMENTO_DINHEIRO, iTipo)
    End If
    
    'Inclui os movimentos referentes ao contravale
    If dVale > 0 Then
        If gsNomePrinc = "SGEECF" Then
            iTipo = MOVIMENTOCAIXA_TROCO_CONTRAVALE
        Else
            iTipo = MOVIMENTOCAIXA_CARNE_TROCO_CONTRAVALE
        End If
        
        Call Inclui_Movimento(StrParaDbl(ContraVale.Text), MEIO_PAGAMENTO_CONTRAVALE, iTipo)
    End If
    
    giRetornoTela = vbOK
    
    Unload Me
    
    End If
    
    Exit Sub
    
Erro_BotaoOK_Click:
       
    Select Case gErr
        
        Case 99689
            Call Rotina_ErroECF(vbOKOnly, ERRO_TROCO_DIFERENTE, gErr)
            
        Case 207999
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175618)

    End Select

    Exit Sub
    
End Sub

Private Sub Inclui_Movimento(dValor As Double, iAdm As Integer, iTipo As Integer)

Dim objMovimento As New ClassMovimentoCaixa
Dim bAchou As Boolean
Dim iIndice As Integer
    
    bAchou = False
    
    'Verifica se existe algum movimento deste tipo
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = iTipo Then
            objMovimento.dValor = dValor
            bAchou = True
            Exit For
        End If
    Next

    If Not (bAchou) Then
        
        Set objMovimento = New ClassMovimentoCaixa
        
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iTipo = iTipo
        objMovimento.iParcelamento = COD_A_VISTA
        objMovimento.dHora = CDbl(Time)
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = dValor
        objMovimento.iAdmMeioPagto = iAdm
        
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    End If
    
End Sub

Private Sub Recalcula_Valores()
'Recalcula todos os valores

    Subtotal.Caption = Format(StrParaDbl(Dinheiro.Text) + StrParaDbl(Ticket.Text) + StrParaDbl(ContraVale.Text), "Standard")
    
    If Format(StrParaDbl(TrocoTotal.Caption) - StrParaDbl(Subtotal.Caption), "Standard") > 0 Then
        Falta.Caption = Format(StrParaDbl(TrocoTotal.Caption) - StrParaDbl(Subtotal.Caption), "Standard")
    Else
        Falta.Caption = ""
    End If
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referência da tela
    Set gobjVenda = Nothing
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    'Clique em F2
    If KeyCode = vbKeyF2 Then
        
        If Not TrocaFoco(Me, BotaoTicket) Then Exit Sub
        Call BotaoTicket_Click
    
    End If
        
    'Clique em f5
    If KeyCode = vbKeyF5 Then
    
        If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
        Call BotaoOK_Click
    
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Troco"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Troco"
    
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


