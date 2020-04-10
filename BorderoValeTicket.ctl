VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoValeTicket 
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   KeyPreview      =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   4770
   Begin VB.ComboBox Parcelamento 
      Height          =   315
      Left            =   1530
      TabIndex        =   25
      ToolTipText     =   "Formas de Parcelamento"
      Top             =   2190
      Width           =   2625
   End
   Begin VB.CommandButton TestaLog 
      Caption         =   "LOG"
      Height          =   705
      Left            =   3180
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "Não Detalhado"
      Height          =   1245
      Left            =   120
      TabIndex        =   20
      Top             =   4065
      Width           =   4515
      Begin MSMask.MaskEdBox ValorEnviarN 
         Height          =   300
         Left            =   2100
         TabIndex        =   6
         Top             =   810
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   6
         Left            =   1455
         TabIndex        =   23
         Top             =   840
         Width           =   510
      End
      Begin VB.Label LabelTotalN 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2100
         TabIndex        =   22
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Caixa Central:"
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
         Index           =   5
         Left            =   480
         TabIndex        =   21
         Top             =   375
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhado"
      Height          =   1245
      Left            =   105
      TabIndex        =   16
      Top             =   2640
      Width           =   4515
      Begin MSMask.MaskEdBox ValorEnviar 
         Height          =   300
         Left            =   2100
         TabIndex        =   5
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Caixa Central:"
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
         Index           =   4
         Left            =   480
         TabIndex        =   19
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label LabelTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2100
         TabIndex        =   18
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   3
         Left            =   1455
         TabIndex        =   17
         Top             =   840
         Width           =   510
      End
   End
   Begin VB.ComboBox AdmMeioPagto 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   1725
      Width           =   2625
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2595
      Picture         =   "BorderoValeTicket.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   780
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   1950
      ScaleHeight     =   495
      ScaleWidth      =   2640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   2700
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1110
         Picture         =   "BorderoValeTicket.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   150
         Picture         =   "BorderoValeTicket.ctx":0274
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2115
         Picture         =   "BorderoValeTicket.ctx":0376
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "BorderoValeTicket.ctx":04F4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   600
         Picture         =   "BorderoValeTicket.ctx":0A26
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox DataEnvio 
      Height          =   300
      Left            =   1530
      TabIndex        =   3
      Top             =   1260
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataEnvio 
      Height          =   300
      Left            =   2490
      TabIndex        =   12
      Top             =   1260
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   765
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Parcelamento:"
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
      Index           =   9
      Left            =   210
      TabIndex        =   26
      ToolTipText     =   "Formas de Parcelamento"
      Top             =   2235
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ticket:"
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
      Index           =   1
      Left            =   825
      TabIndex        =   15
      Top             =   1770
      Width           =   615
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   780
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   14
      Top             =   810
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data de Envio:"
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
      Index           =   2
      Left            =   150
      TabIndex        =   13
      Top             =   1305
      Width           =   1290
   End
End
Attribute VB_Name = "BorderoValeTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iAlterado As Integer
Dim giAdmMeioPagtoVelho As Integer
Dim giParcelamentoVelho As Integer

Private WithEvents objEventoBorderoValeTicket As AdmEvento
Attribute objEventoBorderoValeTicket.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub BorderoValeTicket_Desmembra_Log(ByVal objBorderoValeTicket As ClassBorderoValeTicket, ByVal objLog As ClassLog)
'Função que recebe um log preenchido e um borderoValeticket em branco e o preenche com os dados do log

Dim iPosicao(0 To 2) As Integer
Dim iIndice As Integer

On Error GoTo Erro_BorderoValeTicket_Desmembra_Log

    'inicializa a posição de início da substring
    iPosicao(0) = 1
    
    'inicializa a posição de fim da substring
    iPosicao(1) = InStr(iPosicao(0), objLog.sLog, Chr(vbKeyEscape))
    
    'inicializa a posição de fim da string inteira
    iPosicao(2) = InStr(iPosicao(0), objLog.sLog, Chr(vbKeyEnd))
    
    'enquanto encontrar vbkeyescapes
    Do While iPosicao(1) <> 0
    
        'atualiza o índice de atributo a ser preenchido
        iIndice = iIndice + 1
        
        'preenche o atributo em questão
        Select Case iIndice
        
            Case 1: objBorderoValeTicket.iFilialEmpresa = StrParaInt(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
            Case 2: objBorderoValeTicket.lNumBordero = StrParaLong(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
            Case 3: objBorderoValeTicket.iAdmMeioPagto = StrParaInt(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
            Case 4: objBorderoValeTicket.iParcelamento = StrParaInt(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
            Case 5: objBorderoValeTicket.dtDataEnvio = StrParaDate(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
            Case 6: objBorderoValeTicket.dtDataImpressao = StrParaDate(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
            Case 7: objBorderoValeTicket.dtDataBackoffice = StrParaDate(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
            Case 8: objBorderoValeTicket.dValor = StrParaDbl(Mid(objLog.sLog, iPosicao(0), iPosicao(1) - iPosicao(0)))
        
        End Select
    
        'atualiza a posição de início da substring
        iPosicao(0) = iPosicao(1) + 1
        
        'atualiza a posição de fim da substring
        iPosicao(1) = InStr(iPosicao(0), objLog.sLog, Chr(vbKeyEscape))
        
        'se ainda não chegou ao fim e acabaram os vbkeyescape,i.e., se está prestes a extrair o último elemento da string
        'aponta para o fim da string
        If iPosicao(0) < iPosicao(2) And iPosicao(1) = 0 Then iPosicao(1) = iPosicao(2)
    
    Loop
    
    Exit Sub

Erro_BorderoValeTicket_Desmembra_Log:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143852)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objBordero As New ClassBorderoValeTicket

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o código estiver vazio-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 120046
    
    'se a data estiver em branco-> erro
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 120047
    
    'se a admmeiopagto estiver em branco-> erro
    If AdmMeioPagto.ListIndex = -1 Then gError 120048
    
    Call Move_Tela_Memoria(objBordero)
    'If lErro <> SUCESSO Then gError 120049

    lErro = CF("BorderoValeTicket_Le", objBordero)
    If lErro <> SUCESSO And lErro <> 107370 Then gError 120050
    
    If lErro = 107370 Then gError 120051
    
    '???? adaptar para bordero valetkt
    'ver expr. selecao, nome tsk, etc..
    'aguardando tsk ficar pronto....
    'lErro = objRelatorio.ExecutarDireto("Borderô Vale/Ticket", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVenda", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 120052

    'Limpa a Tela
    Call Limpa_Tela_BorderoValeTicket

    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 120046
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 120047
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 120048
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_SELECIONADO", gErr)
        
        Case 120049, 120050, 120052

        Case 120051
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROVALETICKET_NAOENCONTRADO", gErr, objBordero.iFilialEmpresa, objBordero.lNumBordero)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 143853)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub


End Sub

Private Sub Parcelamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcelamento_Click()

Dim lErro As Long

On Error GoTo Erro_Parcelamento_Click

    lErro = Testa_Alteracao_Parcelamento()
    If lErro <> SUCESSO Then gError 105802
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_Parcelamento_Click:

    Select Case gErr
        
        Case 105802
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143854)

    End Select
    
    Exit Sub

End Sub

Private Sub Parcelamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Parcelamento_Validate

    'se o parcelamento estiver prenchido
    If Len(Trim(Parcelamento.Text)) <> 0 Then
    
        'se o parcelamento for diferente do último selecionado
        If Parcelamento.ListIndex = -1 Then
            
            'tenta selecionar
            lErro = Combo_Seleciona(Parcelamento, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 105803
            
            'se não encontrar pelo código-> erro
            If lErro = 6730 Then gError 105804
            
            'se não encontrar pelo nomereduzido-> erro
            If lErro = 6731 Then gError 105805
        
        End If
        
        lErro = Testa_Alteracao_Parcelamento()
        If lErro <> SUCESSO Then gError 105806
    
    'se não estiver preenchido
    Else
    
        'limpa a label total
        LabelTotal.Caption = Format(0, "STANDARD")
        
    End If
    
    Exit Sub

Erro_Parcelamento_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 105803, 105806
        
        Case 105804
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAOENCONTRADO", gErr, iCodigo)
        
        Case 105805
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAOENCONTRADO", gErr, Parcelamento.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143855)

    End Select

    Exit Sub

End Sub

Private Sub TestaLog_Click()

Dim lErro As Long
Dim objLog As New ClassLog
Dim objBorderoValeTicket As New ClassBorderoValeTicket

On Error GoTo Erro_TestaLog_Click

    lErro = Log_Le(objLog)
    If lErro <> SUCESSO And lErro <> 108015 Then gError 108016
    
    If lErro = 108015 Then gError 108017
    
    Call BorderoValeTicket_Desmembra_Log(objBorderoValeTicket, objLog)
    
    Exit Sub

Erro_TestaLog_Click:

    Select Case gErr
    
        Case 108016
        
        Case 108017
            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAOENCONTRADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143856)

    End Select
    
    Exit Sub

End Sub

Private Function Log_Le(objLog As ClassLog) As Long

Dim lErro As Long
Dim tLog As typeLog
Dim lComando As Long

On Error GoTo Erro_Log_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108012

    'Inicializa o Buffer da Variáveis String
    tLog.sLog1 = String(STRING_CONCATENACAO, 0)
    tLog.sLog2 = String(STRING_CONCATENACAO, 0)
    tLog.sLog3 = String(STRING_CONCATENACAO, 0)
    tLog.sLog4 = String(STRING_CONCATENACAO, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log WHERE Operacao=37 ", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dtData, tLog.dHora)
    If lErro <> SUCESSO Then gError 108013

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108014

    If lErro = AD_SQL_SUCESSO Then

        'Carrega o objLog com as Infromações de bonco de dados
        objLog.lNumIntDoc = tLog.lNumIntDoc
        objLog.iOperacao = tLog.iOperacao
        objLog.sLog = tLog.sLog1 & tLog.sLog2 & tLog.sLog3 & tLog.sLog4
        objLog.dtData = tLog.dtData
        objLog.dHora = tLog.dHora

    End If

    If lErro = AD_SQL_SEM_DADOS Then gError 108015
    
    Log_Le = SUCESSO

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_Log_Le:

    Log_Le = gErr

   Select Case gErr

    Case gErr
    
        Case 108012
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108013, 108014
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)
    
        Case 108015
            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAO_EXISTENTE", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143857)

        End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objTMPLojaFilial As New ClassTMPLojaFilial

On Error GoTo Erro_Form_Load

    'carrega a combo de tickets
    lErro = Carrega_ValeTicket()
    If lErro <> SUCESSO Then gError 107366

    'preenche a data com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataEnvio.PromptInclude = True

    'preenche um tmplojafilial para ler o seu saldo
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_VALE_TICKET
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa

    'le o seu saldo
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 107365
    
    'preenche o total não especificado
    LabelTotalN.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    'preenche o total especificado
    LabelTotal.Caption = Format(0, "STANDARD")
    
    'instancia o objeto com eventos
    Set objEventoBorderoValeTicket = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 107365, 107366
        
        Case 107401
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTOLOJAFILIAL_NAOENCONTRADO", gErr, objTMPLojaFilial.iFilialEmpresa, objTMPLojaFilial.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143858)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoValeTicket As ClassBorderoValeTicket) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not objBorderoValeTicket Is Nothing Then
    
        'se o número do borderô estiver preenchido
        If objBorderoValeTicket.lNumBordero <> 0 Then
            
            'Transfere para o Trata_Parametros
            'lê o bordero
            lErro = CF("BorderoValeTicket_Le", objBorderoValeTicket)
            If lErro <> SUCESSO And lErro <> 107370 Then gError 107475
            
            'se não encontrou-> erro
            If lErro <> 107370 Then
        
                'busca o borderovaleticket
                lErro = Traz_BorderoValeTicket_Tela(objBorderoValeTicket)
                If lErro <> SUCESSO Then gError 107399
            
            End If
        
        Else
            
            'limpa a tela
            Call Limpa_Tela_BorderoValeTicket
            
            'preenche o codigo com o codigo buscado
            Codigo.Text = objBorderoValeTicket.lNumBordero
        
        End If

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr

        Case 107399, 107475

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143859)

    End Select

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objBorderoValeTicket As New ClassBorderoValeTicket
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCodigo_Click

    'se o código estiver preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then
    
        'preenche um borderovaleticket com os dados necessários para chamar o browser
        objBorderoValeTicket.iFilialEmpresa = giFilialEmpresa
        objBorderoValeTicket.lNumBordero = StrParaLong(Codigo.Text)
    
    End If

    Call Chama_Tela("BorderoValeTicketLista", colSelecao, objBorderoValeTicket, objEventoBorderoValeTicket)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143860)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoBorderoValeTicket_evSelecao(obj1 As Object)

Dim lErro As Long

Dim objBorderoValeTicket As ClassBorderoValeTicket

On Error GoTo Erro_objEventoBorderoValeTicket_evSelecao

    'seta o objBorderoValeticket com os dados do obj recebido por parâmetro
    Set objBorderoValeTicket = obj1
    
    'preenche a tela
    lErro = Traz_BorderoValeTicket_Tela(objBorderoValeTicket)
    If lErro <> SUCESSO Then gError 107402
    
    Me.Show
    
    Exit Sub

Erro_objEventoBorderoValeTicket_evSelecao:
    
    Select Case gErr
    
        Case 107402
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROVALETICKET_NAOENCONTRADO", gErr, objBorderoValeTicket.iFilialEmpresa, objBorderoValeTicket.lNumBordero)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143861)

    End Select
    
    Exit Sub

End Sub

Public Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim objBorderoValeTicket As New ClassBorderoValeTicket

On Error GoTo Erro_Tela_Extrai

    sTabela = "BorderoValeTicket"
    
    'preenche o obj com os dados da tela
    Call Move_Tela_Memoria(objBorderoValeTicket)
    
    'preenche a coleção de campos-valor
    colCampoValor.Add "FilialEmpresa", objBorderoValeTicket.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "NumBordero", objBorderoValeTicket.lNumBordero, 0, "NumBordero"
    colCampoValor.Add "AdmMeioPagto", objBorderoValeTicket.iAdmMeioPagto, 0, "AdmMeioPagto"
    colCampoValor.Add "Parcelamento", objBorderoValeTicket.iParcelamento, 0, "Parcelamento"
    colCampoValor.Add "DataEnvio", objBorderoValeTicket.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataImpressao", objBorderoValeTicket.dtDataImpressao, 0, "DataImpressao"
    colCampoValor.Add "DataBackoffice", objBorderoValeTicket.dtDataBackoffice, 0, "DataBackoffice"
    colCampoValor.Add "Valor", objBorderoValeTicket.dValor, 0, "Valor"
    colCampoValor.Add "NumIntDocCPR", objBorderoValeTicket.lNumIntDocCPR, 0, "NumIntDocCPR"

    'estabelece os filtros
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO
    
    Exit Function

Erro_Tela_Extrai:
    
    Tela_Extrai = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143862)

    End Select
    
    Exit Function

End Function

Public Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objBorderoValeTicket As New ClassBorderoValeTicket

On Error GoTo Erro_Tela_Preenche

    'preenche os dados necessários para um borderovaleticket ser encontrado
    objBorderoValeTicket.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objBorderoValeTicket.lNumBordero = colCampoValor.Item("NumBordero").vValor
    objBorderoValeTicket.dtDataBackoffice = colCampoValor.Item("DataBackoffice").vValor
    objBorderoValeTicket.dtDataEnvio = colCampoValor.Item("DataEnvio").vValor
    objBorderoValeTicket.dtDataImpressao = colCampoValor.Item("DataImpressao").vValor
    objBorderoValeTicket.dValor = colCampoValor.Item("Valor").vValor
    objBorderoValeTicket.iAdmMeioPagto = colCampoValor.Item("AdmMeioPagto").vValor
    objBorderoValeTicket.lNumIntDocCPR = colCampoValor.Item("NumIntDocCpr").vValor
    objBorderoValeTicket.iParcelamento = colCampoValor.Item("Parcelamento").vValor
    
    'traz o bordero para a tela
    lErro = Traz_BorderoValeTicket_Tela(objBorderoValeTicket)
    If lErro <> SUCESSO And lErro <> 103372 Then gError 107404
    
    If lErro = 103372 Then gError 107405
    
    Tela_Preenche = SUCESSO
    
    Exit Function
    
Erro_Tela_Preenche:

    Tela_Preenche = gErr
    
    Select Case gErr
    
        Case 107404
        
        Case 107405
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROVALETICKET_NAOENCONTRADO", gErr, objBorderoValeTicket.iFilialEmpresa, objBorderoValeTicket.lNumBordero)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143863)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()

Dim lCodigo As Long
Dim lErro As Long

On Error GoTo Erro_BotaoProxNum_Click

    'gera o próximo número de borderô
    lErro = BorderoValeTicket_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 107407
    
    'coloca o código na tela
    Codigo.Text = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:
    
    Select Case gErr
        
        Case 107407

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143864)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_Codigo_Validate

    'se o codigo estiver em branco-> sai
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub
    
    'critica o código digitado
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 107408
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 107408
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143865)

    End Select
    
    Exit Sub

End Sub

Private Sub DataEnvio_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvio, iAlterado)

End Sub

Private Sub DataEnvio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvio_Validate

    'se a data não estiver preenchida-> sai
    If Len(Trim(DataEnvio.ClipText)) = 0 Then Exit Sub
    
    'critica a data
    lErro = Data_Critica(DataEnvio.Text)
    If lErro <> SUCESSO Then gError 107409

    Exit Sub

Erro_DataEnvio_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 107409
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143866)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub UpDownDataEnvio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_DownClick

    'diminui a data
    lErro = Data_Up_Down_Click(DataEnvio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 107410
    
    Exit Sub
    
Erro_UpDownDataEnvio_DownClick:
    
    Select Case gErr
    
        Case 107410
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143867)

    End Select
    
    Exit Sub

End Sub

Private Sub UpDownDataEnvio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvio_UpClick

    'diminui a data
    lErro = Data_Up_Down_Click(DataEnvio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 107411
    
    Exit Sub
    
Erro_UpDownDataEnvio_UpClick:
    
    Select Case gErr
    
        Case 107411
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143868)

    End Select
    
    Exit Sub

End Sub

Private Sub AdmMeioPagto_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_AdmMeioPagto_Click
    
    iAlterado = REGISTRO_ALTERADO
    
    If AdmMeioPagto.ListIndex <> -1 Then
    
    'se o codigo atual for diferente do anterior
    If Codigo_Extrai(AdmMeioPagto.Text) <> giAdmMeioPagtoVelho Then
    
        'carrega a combo de parcelamento
        lErro = Carrega_Parcelamento(Codigo_Extrai(AdmMeioPagto.Text))
        If lErro <> SUCESSO Then gError 105807
        
        'guarda o código velho
        giAdmMeioPagtoVelho = Codigo_Extrai(AdmMeioPagto.Text)
    
        giParcelamentoVelho = 0
    
    End If
    
    
    End If

    Exit Sub

Erro_AdmMeioPagto_Click:

    Select Case gErr
    
        Case 105807
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143869)

    End Select
    
    Exit Sub

End Sub

Private Sub AdmMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AdmMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_AdmMeioPagto_Validate

    'se a combo está preenchida
    If Len(Trim(AdmMeioPagto.Text)) <> 0 Then
    
        'se o item não foi selecionado na lista
        If AdmMeioPagto.ListIndex = -1 Then
        
            'tenta selecionar na combo
            lErro = Combo_Seleciona(AdmMeioPagto, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 107389
            
            'se não encontrou pelo código-> erro
            If lErro = 6730 Then gError 107390
            
            'se não encontrou pelo nomereduzido-> erro
            If lErro = 6731 Then gError 107391
        
        End If
            
        'se o codigo atual for diferente do anterior
        If Codigo_Extrai(AdmMeioPagto.Text) <> giAdmMeioPagtoVelho Then
        
            'carrega a combo de parcelamento
            lErro = Carrega_Parcelamento(Codigo_Extrai(AdmMeioPagto.Text))
            If lErro <> SUCESSO Then gError 108081
            
            'guarda o código velho
            giAdmMeioPagtoVelho = Codigo_Extrai(AdmMeioPagto.Text)
        
        End If
            
    Else
        
        Parcelamento.Text = ""
        
        'limpa a combo de parcelamentos
        Parcelamento.Clear
    
        'limpa a label de total
        LabelTotal.Caption = Format(0, "STANDARD")
    
    End If
    
    Exit Sub

Erro_AdmMeioPagto_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 107389, 107394
        
        Case 107390
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, Codigo_Extrai(AdmMeioPagto.Text))
        
        Case 107391
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, AdmMeioPagto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143870)

    End Select
    
    Exit Sub

End Sub

Private Sub AdmMeioPagto_GotFocus()
    
    'guarda o código velho
    giAdmMeioPagtoVelho = Codigo_Extrai(AdmMeioPagto.Text)

End Sub

Private Function Testa_Alteracao_Parcelamento()

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objBorderoValeTicket As New ClassBorderoValeTicket

On Error GoTo Erro_Testa_Alteracao_Parcelamento

    'se estiver mudando o parcelamento
    If giParcelamentoVelho <> Codigo_Extrai(Parcelamento.Text) Then

        'preenche um admmeiopagtocondpagto para buscar seu saldo
        objAdmMeioPagtoCondPagto.iFilialEmpresa = giFilialEmpresa
        objAdmMeioPagtoCondPagto.iAdmMeioPagto = Codigo_Extrai(AdmMeioPagto.Text)
        objAdmMeioPagtoCondPagto.iParcelamento = Codigo_Extrai(Parcelamento.Text)

        'tenta buscar na tabela admmeiopagtocondpagto
        lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
        If lErro <> SUCESSO And lErro <> 107297 Then gError 108006

        'se não encontrar->erro
        If lErro = 107297 Then gError 108007

        'preenche o saldo
        LabelTotal.Caption = Format(objAdmMeioPagtoCondPagto.dSaldo, "STANDARD")

        'preenche um objBorderoValeTicket para a busca
        objBorderoValeTicket.iFilialEmpresa = giFilialEmpresa
        objBorderoValeTicket.lNumBordero = StrParaLong(Codigo.Text)

        'lê um borderovaleticket
        lErro = CF("BorderoValeTicket_Le", objBorderoValeTicket)
        If lErro <> SUCESSO And lErro <> 107370 Then gError 108011

        'se encontrou e o admmeiopagto e o parcelamento batem com os q estão na tela, atualizar o saldo
        If lErro = SUCESSO _
        And objBorderoValeTicket.iAdmMeioPagto = Codigo_Extrai(AdmMeioPagto.Text) _
        And objBorderoValeTicket.iParcelamento = Codigo_Extrai(Parcelamento.Text) _
        Then LabelTotal.Caption = Format(StrParaDbl(LabelTotal.Caption) + objBorderoValeTicket.dValor, "STANDARD")

        'guarda o parcelamento velho
        giParcelamentoVelho = Codigo_Extrai(Parcelamento.Text)

    End If

    Testa_Alteracao_Parcelamento = SUCESSO

    Exit Function

Erro_Testa_Alteracao_Parcelamento:

    Testa_Alteracao_Parcelamento = gErr

    Select Case gErr

        Case 108006

        Case 108007
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_ADMMEIOPAGTO_NAOENCONTRADO", gErr, objAdmMeioPagtoCondPagto.iParcelamento, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iAdmMeioPagto)

        Case 108011

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143871)

    End Select

    Exit Function

End Function

Private Sub ValorEnviar_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorEnviar_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorEnviar_Validate

    If Len(Trim(ValorEnviar.Text)) = 0 Then Exit Sub

    'Critica o Valor digitado
    lErro = Valor_NaoNegativo_Critica(ValorEnviar.Text)
    If lErro <> SUCESSO Then gError 107416
    
    Exit Sub

Erro_ValorEnviar_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 107416
    
        Case 107471
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORENVIAR_VALORDISPONIVEL", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143872)

    End Select

    Exit Sub

End Sub

Private Sub ValorEnviarN_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorEnviarN_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorEnviarN_Validate
    
    If Len(Trim(ValorEnviarN.Text)) = 0 Then Exit Sub

    'Critica o Valor digitado
    lErro = Valor_NaoNegativo_Critica(ValorEnviarN.Text)
    If lErro <> SUCESSO Then gError 107417
    
    Exit Sub

Erro_ValorEnviarN_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 107417
        
        Case 107470
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORENVIAR_VALORDISPONIVEL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143873)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a função de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 107418
    
    Call Limpa_Tela_BorderoValeTicket
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr
        
        Case 107418
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143874)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objBorderoValeTicket As New ClassBorderoValeTicket

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se estiver no bo->erro
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 107419
    
    'se o código estiver vazio-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 107420
    
    'se a data estiver em branco-> erro
    If Len(Trim(DataEnvio.ClipText)) = 0 Then gError 107421
    
    'se a admmeiopagto estiver em branco-> erro
    If AdmMeioPagto.ListIndex = -1 Then gError 107422
    
    'se o parcelamento estiver em branco-> erro
    If Parcelamento.ListIndex = -1 Then gError 107423
    
    'se a soma dos valores especificados e não especificados foir igual a 0-> erro
    If StrParaDbl(ValorEnviar.Text) + StrParaDbl(ValorEnviarN.Text) = 0 Then gError 107424
    
    'preenche o bordero com os dados da tela
    Call Move_Tela_Memoria(objBorderoValeTicket)
    
    'testa se é uma alteração
    lErro = Trata_Alteracao(objBorderoValeTicket, objBorderoValeTicket.iFilialEmpresa, objBorderoValeTicket.lNumBordero)
    If lErro <> SUCESSO Then gError 107426
    
    'grava o bordero de valeticket
    lErro = CF("BorderoValeTicket_Grava", objBorderoValeTicket)
    If lErro <> SUCESSO Then gError 107425

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 107419
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROVALETICKET_GRAVACAO_BACKOFFICE", gErr)
        
        Case 107420
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 107421
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 107422
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_SELECIONADO", gErr)
        
        Case 107423
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAO_SELECIONADO1", gErr)
        
        Case 107424
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDEROVALETICKET_ZERADO", gErr)
        
        Case 107425, 107426
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143875)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgResp As VbMsgBoxResult
Dim objBorderoValeTicket As New ClassBorderoValeTicket

On Error GoTo Erro_BotaoExcluir_Click

    'se o código não estiver preenchido-> erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 107481
    
    vbMsgResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_BORDEROVALETICKET", giFilialEmpresa, Codigo.Text)
    
    If vbMsgResp = vbYes Then
    
        'preenche os atributos necessários à exclusão do bordero
        objBorderoValeTicket.iFilialEmpresa = giFilialEmpresa
        objBorderoValeTicket.lNumBordero = StrParaLong(Codigo.Text)
        
        'exclui o borderÔ
        lErro = CF("BorderoValeTicket_Exclui", objBorderoValeTicket)
        If lErro <> SUCESSO Then gError 107482
        
        Call Limpa_Tela_BorderoValeTicket
        
        iAlterado = 0
    
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:
    
    Select Case gErr
    
        Case 107481
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 107482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143876)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro

On Error GoTo Erro_Botaolimpar_Click
    
    'testa se houve alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 107477
    
    Call Limpa_Tela_BorderoValeTicket
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 107478
    
    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:
    
    Select Case gErr
        
        Case 107477, 107478
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143877)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_BorderoValeTicket()

Dim lErro As Long
Dim objTMPLojaFilial As New ClassTMPLojaFilial

On Error GoTo Erro_Limpa_Tela_BorderoValeTicket

    Call Limpa_Tela(Me)

    LabelTotal.Caption = Format(0, "STANDARD")
    
    'preenche um tmplojafilial para ler o seu saldo
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_VALE_TICKET
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa

    'le o seu saldo
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 107479
    
    LabelTotalN.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    'preenche a data com a data atual
    DataEnvio.PromptInclude = False
    DataEnvio.Text = DataEnvio.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataEnvio.PromptInclude = True
    
    AdmMeioPagto.ListIndex = -1
    Parcelamento.Text = ""
    Parcelamento.Clear
    
    giAdmMeioPagtoVelho = 0
    giParcelamentoVelho = 0

    Exit Sub

Erro_Limpa_Tela_BorderoValeTicket:

    Select Case gErr
    
        Case 107479
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143878)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub form_unload(Cancel As Integer)

    'libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

    'libera a memória
    Set objEventoBorderoValeTicket = Nothing
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Function Carrega_Parcelamento(iCodigo As Integer) As Long

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_Parcelamento

    'preenche os atributos para buscar a admmeiopagtocondpagto da admmeiopagto
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
    objAdmMeioPagto.iCodigo = iCodigo

    'busca no BD e preenche colcondpagtoloja com os parcelamentos
    lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104086 Then gError 107392

    'se não encontrar-> erro
    If lErro = 104086 Then gError 107393

    Parcelamento.Text = ""

    'limpa a combo
    Parcelamento.Clear

    'preenche a combo com os novos valores
    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja

        Parcelamento.AddItem (objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento)
        Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento

    Next

    LabelTotal.Caption = ""

    Carrega_Parcelamento = SUCESSO

    Exit Function

Erro_Carrega_Parcelamento:

    Carrega_Parcelamento = gErr

    Select Case gErr

        Case 107392

        Case 107393
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTOS_ADMMEIOPAGTO_NAOENCONTRADOS", gErr, objAdmMeioPagto.iFilialEmpresa, objAdmMeioPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143879)

    End Select

    Exit Function

End Function

Public Function Traz_BorderoValeTicket_Tela(objBorderoValeTicket As ClassBorderoValeTicket) As Long

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objTMPLojaFilial As New ClassTMPLojaFilial

On Error GoTo Erro_Traz_BorderoValeTicket_Tela

    Call Limpa_Tela_BorderoValeTicket

    'preenche a tela
    Codigo.Text = objBorderoValeTicket.lNumBordero
    DataEnvio.PromptInclude = False
    DataEnvio.Text = Format(objBorderoValeTicket.dtDataEnvio, "dd/mm/yy")
    DataEnvio.PromptInclude = True
    
    AdmMeioPagto.Text = objBorderoValeTicket.iAdmMeioPagto
    Call AdmMeioPagto_Validate(bSGECancelDummy)
       
    Parcelamento.Text = objBorderoValeTicket.iParcelamento
    Call Parcelamento_Validate(bSGECancelDummy)
        
    ValorEnviar.Text = Format(objBorderoValeTicket.dValor, "STANDARD")
    
    'preenche os atributos necessários para buscar uma admmeiopagtocondpagto específica na referida tabela
    objAdmMeioPagtoCondPagto.iAdmMeioPagto = objBorderoValeTicket.iAdmMeioPagto
    objAdmMeioPagtoCondPagto.iParcelamento = objBorderoValeTicket.iParcelamento
    objAdmMeioPagtoCondPagto.iFilialEmpresa = objBorderoValeTicket.iFilialEmpresa
    
    'tenta encontrar o admmeiopagtocondpagto que atenda às condições acima
    lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
    If lErro <> SUCESSO And lErro <> 107297 Then gError 107395
    
    'se não encontrar-> erro
    If lErro = 107297 Then gError 107396
    
    'preenche o total
    LabelTotal.Caption = Format((objAdmMeioPagtoCondPagto.dSaldo + objBorderoValeTicket.dValor), "STANDARD")
    
    'preenche os atributos necessários para buscar um tmplojafilial
    objTMPLojaFilial.iFilialEmpresa = giFilialEmpresa
    objTMPLojaFilial.iTipo = TIPOMEIOPAGTOLOJA_VALE_TICKET
    
    'lê o saldo na tabela de não especificados
    lErro = CF("TipoMeioPagtoLojaFilial_Le", objTMPLojaFilial)
    If lErro <> SUCESSO Then gError 107397
    
    'preenche o total não especificado
    LabelTotalN.Caption = Format(objTMPLojaFilial.dSaldo, "STANDARD")
    
    iAlterado = 0

    Traz_BorderoValeTicket_Tela = SUCESSO

    Exit Function

Erro_Traz_BorderoValeTicket_Tela:

    Traz_BorderoValeTicket_Tela = gErr

    Select Case gErr
    
        Case 107395, 107397
        
        Case 107396
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_ADMMEIOPAGTO_NAOENCONTRADO", gErr, objAdmMeioPagtoCondPagto.iParcelamento, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iAdmMeioPagto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143880)

    End Select

    Exit Function

End Function

Private Function Carrega_ValeTicket() As Long

Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_Carrega_ValeTicket

    'le as admmeiopagto do tipomeiopagto ticket
    lErro = CF("AdmMeioPagto_Le_TipoMeioPagto", TIPOMEIOPAGTOLOJA_VALE_TICKET, colAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 107360 Then gError 107363
    
    'se não encontrar nenhuma -> erro
    If lErro = 107360 Then gError 107364

    'preenche a combo
    For Each objAdmMeioPagto In colAdmMeioPagto

        AdmMeioPagto.AddItem (objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome)
        AdmMeioPagto.ItemData(AdmMeioPagto.NewIndex) = objAdmMeioPagto.iCodigo

    Next

    Carrega_ValeTicket = SUCESSO

    Exit Function

Erro_Carrega_ValeTicket:

    Carrega_ValeTicket = gErr

    Select Case gErr

        Case 107363

        Case 107364
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_TIPOMEIOPAGTO_VAZIA", gErr, TIPOMEIOPAGTOLOJA_VALE_TICKET)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143881)

    End Select

    Exit Function

End Function

Private Sub Move_Tela_Memoria(objBorderoValeTicket As ClassBorderoValeTicket)

On Error GoTo Erro_Move_Tela_Memoria

    'preenche os atributos de um borderovaleticket
    objBorderoValeTicket.lNumBordero = StrParaLong(Codigo.Text)
    objBorderoValeTicket.dtDataEnvio = StrParaDate(DataEnvio.Text)
    objBorderoValeTicket.iAdmMeioPagto = Codigo_Extrai(AdmMeioPagto.Text)
    objBorderoValeTicket.iParcelamento = Codigo_Extrai(Parcelamento.Text)
    objBorderoValeTicket.dValor = StrParaDbl(ValorEnviar.Text)
    objBorderoValeTicket.dValorN = StrParaDbl(ValorEnviarN.Text)
    objBorderoValeTicket.dtDataBackoffice = DATA_NULA
    objBorderoValeTicket.dtDataImpressao = DATA_NULA
    objBorderoValeTicket.iFilialEmpresa = giFilialEmpresa
    objBorderoValeTicket.sAdmMeioPagto = Nome_Extrai(AdmMeioPagto.Text)
    objBorderoValeTicket.sNomeParcelamento = Nome_Extrai(Parcelamento.Text)

    Exit Sub

Erro_Move_Tela_Memoria:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143882)

    End Select

    Exit Sub

End Sub



Private Function BorderoValeTicket_Codigo_Automatico(lCodigo As Long) As Long

Dim lErro As Long

On Error GoTo Erro_BorderoValeTicket_Codigo_Automatico

    'busca o próximo número de borderô automático
    lErro = CF("Config_ObterAutomatico", "LojaConfig", "COD_PROX_BORDEROVALETICKET", "BorderoValeTicket", "NumBordero", lCodigo)
    If lErro <> SUCESSO Then gError 107406
    
    BorderoValeTicket_Codigo_Automatico = SUCESSO
    
    Exit Function

Erro_BorderoValeTicket_Codigo_Automatico:
    
    BorderoValeTicket_Codigo_Automatico = gErr
    
    Select Case gErr
        
        Case 107406

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143883)

    End Select

    Exit Function

End Function

Private Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Trim(Mid(sTexto, iPosicao + 1))

    Nome_Extrai = sString

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        Call LabelCodigo_Click
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Borderô Vale Ticket"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoValeTicket"

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
