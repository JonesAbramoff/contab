VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BorderoChequesPre2Ocx 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   ScaleHeight     =   3285
   ScaleWidth      =   4290
   Begin MSFlexGridLib.MSFlexGrid GridParcelas 
      Height          =   495
      Left            =   -20000
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
   End
   Begin VB.CommandButton BotaoIntAtualiza 
      Caption         =   "Interromper Atualização"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   570
      TabIndex        =   0
      Top             =   2010
      Width           =   3105
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   705
      ScaleHeight     =   495
      ScaleWidth      =   2775
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2550
      Width           =   2835
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   90
         Picture         =   "BorderoChequesPre2Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2280
         Picture         =   "BorderoChequesPre2Ocx.ctx":075E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoAtualizar 
         Caption         =   "Atualizar"
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
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1050
      End
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   345
      Left            =   375
      TabIndex        =   4
      Top             =   1440
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label TotalCheques 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2670
      TabIndex        =   5
      Top             =   570
      Width           =   1365
   End
   Begin VB.Label ChequesProcessados 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2670
      TabIndex        =   6
      Top             =   990
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Baixas Processadas:"
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
      Left            =   780
      TabIndex        =   7
      Top             =   990
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total de Baixas de Parcelas:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   585
      Width           =   2460
   End
   Begin VB.Label Label4 
      Caption         =   "Atualização de Arquivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   615
      TabIndex        =   9
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "BorderoChequesPre2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giExecutando As Integer
Dim giCancelaBatch As Integer
Dim gobjBorderoChequePre As ClassBorderoChequePre

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloRec As ClassTituloReceber
Private gobjParcelaRec As ClassParcelaReceber
Private gobjBaixaParcRec As ClassBaixaParcRec
Private gobjBaixaReceber As ClassBaixaReceber
Private gsContaCtaCorrente As String 'conta contabil da conta corrente onde foram depositados os cheques
Private gsContaFilDep As String 'conta contabil da filial recebedora
Private giFilialEmpresaConta As Integer 'filial empresa possuidora da conta corrente utilizada p/o deposito

Private Sub BotaoAtualizar_Click()

Dim lErro As Long
Dim sMsgDebRecCli As String

On Error GoTo Erro_BotaoAtualizar_Click

    BotaoAtualizar.Enabled = False
    
    BotaoIntAtualiza.Enabled = True
        
    If giCancelaBatch <> CANCELA_BATCH Then
    
         giExecutando = ESTADO_ANDAMENTO
         
         gobjBorderoChequePre.iTipoBordero = TIPO_BORDERO_CHEQUEPRE
         gobjBorderoChequePre.iFilialEmpresa = giFilialEmpresa
         
         gobjBorderoChequePre.objTelaAtualizacao = Me
         lErro = CF("BorderoChequePre_Atualizar", gobjBorderoChequePre, sMsgDebRecCli)
         
         giExecutando = ESTADO_PARADO
         
         BotaoIntAtualiza.Enabled = False
         
         If lErro <> SUCESSO And lErro <> 59180 Then Error 59176
             
         If lErro = 59180 Then Error 59177 'interrompeu
         
         'se foram gerados creditos para clientes ==> exibe aviso
         If Len(sMsgDebRecCli) > 0 Then Call Rotina_Aviso(vbOKOnly, sMsgDebRecCli)
         
         'Chama a tela do passo seguinte
         Call Chama_Tela("BorderoChequesPre3", gobjBorderoChequePre)
                 
         'Fecha a tela
         Unload Me
                 
    End If
    
Exit Sub
    
Erro_BotaoAtualizar_Click:

    Select Case Err
        
        Case 59177
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")
            Unload Me
            
        Case 59176
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143640)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    If giExecutando = ESTADO_ANDAMENTO Then Error 59178
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoFechar_Click:
        
    Select Case Err
    
        Case 59178
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_IMPOSSIVEL_CONTINUAR")
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143641)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoIntAtualiza_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        
        giCancelaBatch = CANCELA_BATCH
        Exit Sub
    
    End If
    
    'Fecha a tela
    Unload Me

End Sub

Function Trata_Parametros(Optional objBorderoChequePre As ClassBorderoChequePre) As Long

Dim lErro As Long
Dim iCount As Integer

On Error GoTo Erro_Trata_Parametros

    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO
    
    Set gobjBorderoChequePre = objBorderoChequePre
    
    BarraProgresso.Min = 0
    BarraProgresso.Max = 100
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143642)
            
    End Select

    giCancelaBatch = CANCELA_BATCH
        
    Exit Function

End Function

Public Function Mostra_Evolucao(iCancela As Integer, iNumProc As Integer) As Long

Dim lErro As Long
Dim iEventos As Integer
Dim iProcessados As Integer
Dim iTotal As Integer

On Error GoTo Erro_Mostra_Evolucao

    iEventos = DoEvents()
    
    If giCancelaBatch = CANCELA_BATCH Then
        
        iCancela = CANCELA_BATCH
        giExecutando = ESTADO_PARADO
        
    Else
        'atualiza dados da tela ( registros atualizados e a barra )
        
        iProcessados = StrParaInt(ChequesProcessados.Caption)
        iTotal = gobjBorderoChequePre.iQuantParcelas
        If iTotal = 0 Then iTotal = 1
        TotalCheques.Caption = CStr(iTotal)
        
        iProcessados = iProcessados + iNumProc
        ChequesProcessados.Caption = CStr(iProcessados)
        
        BarraProgresso.Value = CInt((iProcessados / iTotal) * 100)
        
        giExecutando = ESTADO_ANDAMENTO
        
    End If
    
    Mostra_Evolucao = SUCESSO

    Exit Function
    
Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143643)
            
    End Select
        
    giCancelaBatch = CANCELA_BATCH
    
    Exit Function
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If giExecutando = ESTADO_ANDAMENTO Then
        If giCancelaBatch <> CANCELA_BATCH Then giCancelaBatch = CANCELA_BATCH
        Cancel = 1
    End If

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjContabAutomatica = Nothing
    Set gobjTituloRec = Nothing
    Set gobjParcelaRec = Nothing
    Set gobjBaixaParcRec = Nothing
    Set gobjBaixaReceber = Nothing
    gsContaCtaCorrente = ""
    gsContaFilDep = ""
    giFilialEmpresaConta = 0
    
    Set gobjBorderoChequePre = Nothing
    
End Sub

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, sContaContabil As String, dValor As Double, iConta As Integer
Dim objFilial As New ClassFilialCliente, sContaTela As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objInfoParcRec As ClassInfoParcRec, objCarteiraCobrador As New ClassCarteiraCobrador
Dim objBaixaParcRec As ClassBaixaParcRec
Dim dTotalRecebido As Double, sContaMascarada As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case "FilialCli_Conta_Ctb"
            For Each objInfoParcRec In gobjBorderoChequePre.colInfoParcRec

                objFilial.lCodCliente = objInfoParcRec.lCliente
                objFilial.iCodFilial = objInfoParcRec.iFilialCliente
                
                lErro = CF("FilialCliente_Le", objFilial)
                If lErro <> SUCESSO Then gError 186138
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 186139
                
                Else
                
                    sContaMascarada = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaMascarada
            Next
        
        
        'Início do trecho adicionado Rafael Menezes em 15/10/2002
        Case "Valor_Recebido_Total"
            objMnemonicoValor.colValor.Add gobjBorderoChequePre.dValorChequesSelecionados
        
        Case "Valor_Recebido_Loja"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                dTotalRecebido = dTotalRecebido + objBaixaParcRec.dValorRecebido
            Next
            objMnemonicoValor.colValor.Add (gobjBorderoChequePre.dValorChequesSelecionados - dTotalRecebido)
        
        'Fim do trecho adicionado Rafael Menezes em 15/10/2002
        Case "Valor_Recebido_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorRecebido
            Next
        
        Case "Valor_Baixar_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorBaixado
            Next
        
        Case "Valor_Desconto_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorDesconto
            Next
        
        Case "Valor_Juros_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorJuros
            Next
        
        Case "Valor_Multa_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorMulta
            Next
        
        Case "Numero_Titulo_Det"
            For Each objInfoParcRec In gobjBorderoChequePre.colInfoParcRec
                objMnemonicoValor.colValor.Add objInfoParcRec.lNumTitulo
            Next
        
        Case "Parcela_Det"
            For Each objInfoParcRec In gobjBorderoChequePre.colInfoParcRec
                objMnemonicoValor.colValor.Add objInfoParcRec.iNumParcela
            Next
        
        Case "Cliente_Codigo_Det"
            For Each objInfoParcRec In gobjBorderoChequePre.colInfoParcRec
                objMnemonicoValor.colValor.Add objInfoParcRec.lCliente
            Next
        
        Case "Cta_CartCobr_Det"
                        
            For Each objInfoParcRec In gobjBorderoChequePre.colInfoParcRec
        
                objCarteiraCobrador.iCobrador = objInfoParcRec.iCobrador
                objCarteiraCobrador.iCodCarteiraCobranca = objInfoParcRec.iCarteiraCobrador
                
                lErro = CartCobr_ObtemCtaTela(objCarteiraCobrador, sContaTela)
                If lErro <> SUCESSO Then Error 32273
                
                objMnemonicoValor.colValor.Add sContaTela
            
            Next
            
        Case "Valor_Pago"
        
            dValor = gobjBaixaParcRec.dValorRecebido
            objMnemonicoValor.colValor.Add dValor
        
        Case "Valor_Baixado"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorBaixado
        
        Case "Valor_Desconto"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorDesconto
        
        Case "Valor_Juros"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorJuros
        
        Case "Valor_Multa"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorMulta
        
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
            'calcula-la apenas uma vez e deixa-la guardada
                
            If gsContaCtaCorrente = "" Then
                
                iConta = gobjBorderoChequePre.iCodNossaConta
                lErro = CF("ContaCorrenteInt_Le", iConta, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 56546
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56547
                                        
                Else
                
                    sContaTela = ""
                    
                End If
                
                gsContaCtaCorrente = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add gsContaCtaCorrente
                
        Case "FilDep_Cta_Transf" 'conta de transferencia da filial do deposito

            If gsContaFilDep = "" Then
            
                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(giFilialEmpresaConta, sContaContabil)
                If lErro <> SUCESSO Then Error 56548
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56549
                    
                Else
                
                    sContaTela = ""
                    
                End If
            
                gsContaFilDep = sContaTela
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaFilDep
        
        Case "FilNaoDep_Cta_Transf" 'conta de transferencia da filial da parcela

                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloRec.iFilialEmpresa, sContaContabil)
                If lErro <> SUCESSO Then Error 56550
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56551
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
        
        Case "Num_Titulo"
        
            objMnemonicoValor.colValor.Add gobjTituloRec.lNumTitulo
            
        Case "Parcela"
            objMnemonicoValor.colValor.Add gobjParcelaRec.iNumParcela
        
        Case "Cliente_Codigo"
            
            objMnemonicoValor.colValor.Add gobjTituloRec.lCliente
                    
        Case "FilialCli_Codigo"
        
            objMnemonicoValor.colValor.Add gobjTituloRec.iFilial
        
        Case Else
        
            Error 56552

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 56546 To 56551, 32273
        
        Case 56552
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143644)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcrec e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iConta As Integer, dValorLivroAux As Double
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjBaixaParcRec = vParams(0)
    Set gobjParcelaRec = vParams(1)
    Set gobjTituloRec = vParams(2)
    Set gobjBaixaReceber = vParams(3)

    'se ainda nao obtive a filial empresa onde vai ser feito o deposito
    If giFilialEmpresaConta = 0 Then
    
        iConta = gobjBorderoChequePre.iCodNossaConta
    
        lErro = CF("ContaCorrenteInt_Le", iConta, objContasCorrentesInternas)
        If lErro <> SUCESSO Then Error 32243
    
        giFilialEmpresaConta = objContasCorrentesInternas.iFilialEmpresa
        
    End If
    
    'obtem numero de doc para a filial onde vai ser feito o deposito
    lErro = objContabAutomatica.Obter_Doc(lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32244
    
    'se contabiliza parcela p/parcela
    If gobjCR.iContabSemDet = 0 Then
    
        dValorLivroAux = Round(gobjBaixaParcRec.dValorRecebido + gobjBaixaParcRec.dValorDesconto - gobjBaixaParcRec.dValorJuros - gobjBaixaParcRec.dValorMulta, 2)
    
        'se a filial onde vai ser feito o deposito é diferente da do titulo
        'e a contabilidade é descentralizada por filiais
        If giFilialEmpresaConta <> gobjTituloRec.iFilialEmpresa And giContabCentralizada = 0 Then
                        
            'grava a contabilizacao na filial onde vai ser feito o deposito
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoChequePreFilDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
            If lErro <> SUCESSO Then Error 32245
        
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloRec.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 32246
        
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoChequePreFilNaoDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloRec.iFilialEmpresa, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32247
        
        Else
        
            'grava a contabilizacao na filial da cta (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoChequePre", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32248
        
        End If
    
    Else
    
        GridParcelas.Tag = gobjBaixaReceber.colBaixaParcRec.Count
    
        'grava a contabilizacao na filial da cta
        lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoChequePreRes", gobjBaixaReceber.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
        If lErro <> SUCESSO Then Error 32248
    
    End If
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32243 To 32248
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143645)
     
    End Select
     
    Exit Function

End Function

Public Sub Form_Load()
    
    lErro_Chama_Tela = SUCESSO
    
End Sub
    

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_DEPOSITO_CHEQUES_PRE_DATADOS2
    Set Form_Load_Ocx = Me
    Caption = "Borderô de Depósito de Cheques Pré-Datados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoChequesPre2"
    
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

Private Sub BotaoVoltar_Click()

    'Chama a tela do passo anterior
    Call Chama_Tela("BorderoChequesPre1", gobjBorderoChequePre)
    
    'Fecha a tela
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

'***** fim do trecho a ser copiado ******




Private Sub TotalCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCheques, Source, X, Y)
End Sub

Private Sub TotalCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCheques, Button, Shift, X, Y)
End Sub

Private Sub ChequesProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ChequesProcessados, Source, X, Y)
End Sub

Private Sub ChequesProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ChequesProcessados, Button, Shift, X, Y)
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

'??? já existe em ctbaixareccancelar.cls
Private Function CartCobr_ObtemCtaTela(objCarteiraCobrador As ClassCarteiraCobrador, sContaTela As String) As Long

Dim lErro As Long, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_CartCobr_ObtemCtaTela

    If objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA Then

        Select Case objCarteiraCobrador.iCodCarteiraCobranca

            Case CARTEIRA_CARTEIRA
                sCampoGlobal = "CtaReceberCarteira"

            Case CARTEIRA_CHEQUEPRE
                sCampoGlobal = "CtaRecChequePre"

            Case CARTEIRA_JURIDICO
                sCampoGlobal = "CtaJuridico"

            Case Else
                Error 56799

        End Select

        objMnemonico.sMnemonico = sCampoGlobal
        lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
        If lErro <> SUCESSO And lErro <> 39690 Then Error 56800
        If lErro <> SUCESSO Then Error 56801

        sContaTela = objMnemonico.sValor

    Else

        lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
        If lErro <> SUCESSO And lErro <> 23551 Then Error 49726
        If lErro <> SUCESSO Then Error 56797

        If objCarteiraCobrador.sContaContabil <> "" Then

            lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaContabil, sContaTela)
            If lErro <> SUCESSO Then Error 56526

        End If

    End If

    CartCobr_ObtemCtaTela = SUCESSO
     
    Exit Function
    
Erro_CartCobr_ObtemCtaTela:

    CartCobr_ObtemCtaTela = gErr
     
    Select Case gErr
          
        Case 49726, 56526, 56800
        
        Case 56797, 56799
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56801
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143646)
     
    End Select
     
    Exit Function

End Function



