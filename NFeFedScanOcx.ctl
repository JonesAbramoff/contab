VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl NFeFedScanOcx 
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   ScaleHeight     =   2760
   ScaleWidth      =   5655
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3270
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "NFeFedScanOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "NFeFedScanOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "NFeFedScanOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "NFeFedScanOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2370
      Picture         =   "NFeFedScanOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Numeração Automática"
      Top             =   450
      Width           =   300
   End
   Begin MSMask.MaskEdBox Ocorrencia 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   435
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownEntrada 
      Height          =   300
      Left            =   2505
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1035
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataEntrada 
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   1035
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox HoraEntrada 
      Height          =   300
      Left            =   4560
      TabIndex        =   5
      Top             =   1035
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "hh:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownSaida 
      Height          =   300
      Left            =   2505
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1590
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataSaida 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1590
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox HoraSaida 
      Height          =   300
      Left            =   4560
      TabIndex        =   10
      Top             =   1590
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "hh:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Justificativa 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   2175
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Motivo:"
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
      Left            =   750
      TabIndex        =   14
      Top             =   2205
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Saída:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   345
      TabIndex        =   12
      Top             =   1635
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hora Saída:"
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
      Index           =   1
      Left            =   3450
      TabIndex        =   11
      Top             =   1635
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Entrada:"
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
      Index           =   905
      Left            =   195
      TabIndex        =   7
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hora Entrada:"
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
      Left            =   3285
      TabIndex        =   6
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label LabelOcorrencia 
      AutoSize        =   -1  'True
      Caption         =   "Ocorrência:"
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
      Left            =   405
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      Top             =   480
      Width           =   990
   End
End
Attribute VB_Name = "NFeFedScanOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Eventos do browse
Private WithEvents objEventoOcorrencia As AdmEvento
Attribute objEventoOcorrencia.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim objCidades As New ClassCidades
Dim colCidades As New Collection

On Error GoTo Erro_Form_Load
    
    'Inicializa o Browse
    Set objEventoOcorrencia = New AdmEvento

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207478)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objNFeFedScan As ClassNFeFedScan) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de CidadeCadastro

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objNFeFedScan Is Nothing) Then

        lErro = CF("NFeFedScan_Le", objNFeFedScan)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207483

        If lErro = SUCESSO Then

            Call Traz_NFeFedScan_Tela(objNFeFedScan)

        Else
            Ocorrencia.Text = objNFeFedScan.lOcorrencia
        End If

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 207483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207484)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_NFeFedScan_Tela(objNFeFedScan As ClassNFeFedScan) As Long
'Preenche a tela com as informações do banco

On Error GoTo Erro_Traz_NFeFedScan_Tela

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Mostra os dados na tela
    Ocorrencia.Text = objNFeFedScan.lOcorrencia
    Call DateParaMasked(DataEntrada, objNFeFedScan.dtDataEntrada)
    HoraEntrada.Text = Format(CDate(objNFeFedScan.dHoraEntrada), "hh:mm:ss")
    Call DateParaMasked(DataSaida, objNFeFedScan.dtDataSaida)
    If objNFeFedScan.dHoraSaida <> 999 Then
        HoraSaida.Text = Format(CDate(objNFeFedScan.dHoraSaida), "hh:mm:ss")
    End If
    Justificativa.Text = objNFeFedScan.sJustificativa
    
    iAlterado = 0

    Exit Function

Erro_Traz_NFeFedScan_Tela:

    Traz_NFeFedScan_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207485)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objNFeFedScan As ClassNFeFedScan) As Long

    'Move os dados da tela para memória
    objNFeFedScan.lOcorrencia = StrParaLong(Ocorrencia.Text)
    
    If Len(Trim(HoraEntrada.ClipText)) > 0 Then
        objNFeFedScan.dHoraEntrada = CDbl(CDate(HoraEntrada.Text))
    Else
        objNFeFedScan.dHoraEntrada = 0
    End If
    
    If Len(Trim(HoraSaida.ClipText)) > 0 Then
        objNFeFedScan.dHoraSaida = CDbl(CDate(HoraSaida.Text))
    Else
        objNFeFedScan.dHoraSaida = 999
    End If
    
    objNFeFedScan.dtDataEntrada = StrParaDate(DataEntrada.Text)
    objNFeFedScan.dtDataSaida = StrParaDate(DataSaida.Text)
    objNFeFedScan.sJustificativa = Justificativa.Text
    objNFeFedScan.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objNFeFedScan As New ClassNFeFedScan
Dim dtDataSaida As Date
Dim dHoraEntrada As Double
Dim dHoraSaida As Double
Dim dtDataEntrada As Date

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(Ocorrencia.Text)) = 0 Then gError 207486

    If Len(Trim(Justificativa.Text)) = 0 Then gError 207487

    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 207488

    If Len(Trim(HoraEntrada.ClipText)) = 0 Then gError 207489

    dtDataEntrada = CDate(DataEntrada.Text)

    'se a data de saida nao estiver preenchida
    If Len(Trim(DataSaida.ClipText)) > 0 Then
        dtDataSaida = CDate(DataSaida.Text)
    Else
        dtDataSaida = DATA_NULA
    End If

    If dtDataEntrada > dtDataSaida And dtDataSaida <> DATA_NULA Then gError 207500
    
    dHoraEntrada = CDbl(CDate(HoraEntrada.Text))
    
    
    If Len(Trim(HoraSaida.ClipText)) > 0 Then
        dHoraSaida = CDbl(CDate(HoraSaida.Text))
    Else
        dHoraSaida = 999
    End If
    
    If dtDataEntrada = dtDataSaida And dHoraEntrada > dHoraSaida Then gError 207501

    lErro = CF("NFeFedScan_Verifica_Datas", StrParaLong(Ocorrencia.Text), giFilialEmpresa, dtDataEntrada, dtDataSaida, dHoraEntrada, dHoraSaida)
    If lErro <> SUCESSO Then gError 207490

    lErro = Move_Tela_Memoria(objNFeFedScan)
    If lErro <> SUCESSO Then gError 207491
    
    lErro = CF("NFeFedScan_Grava", objNFeFedScan)
    If lErro <> SUCESSO Then gError 207492

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
        
        Case 207486
            Call Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_NAO_PREENCHIDA", gErr)
        
        Case 207487
            Call Rotina_Erro(vbOKOnly, "ERRO_JUSTIFICATIVA_NAO_PREENCHIDA", gErr)

        Case 207488
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_ENTRADA_NAO_PREENCHIDA", gErr)

        Case 207489
            Call Rotina_Erro(vbOKOnly, "ERRO_HORA_ENTRADA_NAO_PREENCHIDA", gErr)
            
        Case 207490 To 207492
            
        Case 207500
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case 207501
            Call Rotina_Erro(vbOKOnly, "ERRO_HORA_ENTRADA_MAIOR_SAIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207493)

    End Select
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objNFeFedScan As New ClassNFeFedScan

On Error GoTo Erro_BotaoExcluir_Click:

    'Verifica se o codigo foi preenchido
    If Len(Trim(Ocorrencia.Text)) = 0 Then gError 207522

    objNFeFedScan.lOcorrencia = StrParaLong(Ocorrencia.Text)

    'Envia aviso perguntando se realmente deseja excluir a Ocorrencia de Scan
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_OCORRENCIA_SCAN", objNFeFedScan.lOcorrencia)

    If vbMsgRes = vbYes Then

        'Exclui a ocorrencia de scan
        lErro = CF("NFeFedScan_Exclui", objNFeFedScan)
        If lErro <> SUCESSO And lErro <> 207514 Then gError 207523

        If lErro <> SUCESSO Then gError 207524

        'Limpa a tela
        Call Limpa_Tela(Me)

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

    End If

    iAlterado = 0

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 207522
            Call Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_NAO_PREENCHIDA", gErr)

        Case 207523

        Case 207524
            Call Rotina_Erro(vbOKOnly, "ERRO_NFEFEDSCAN_NAO_CADASTRADO", gErr, objNFeFedScan.lOcorrencia)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207525)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 207526

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 207526

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207527)

    End Select
    
    iAlterado = 0

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 207528

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 207529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207530)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lOcorrencia As Long

On Error GoTo Erro_BotaoProxNum_Click


    'Obtém o próximo código disponível para NFeFedScan
    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_NFEFEDSCAN", "NFeFedScan", "Ocorrencia", lOcorrencia)
    If lErro <> SUCESSO Then gError 207531
    
    'Coloca o Código obtido na tela
    Ocorrencia.Text = lOcorrencia
        
    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 207531
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207532)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub LabelOcorrencia_Click()

Dim objNFeFedScan As New ClassNFeFedScan
Dim colSelecao As New Collection
Dim objNF As New ClassNFiscal
    
    'Preenche na memória o Código passado
    If Len(Trim(Ocorrencia.ClipText)) > 0 Then objNFeFedScan.lOcorrencia = Ocorrencia.Text

    objNF.iFilialEmpresa = giFilialEmpresa
    Call CF("NFiscal_FilialEmpresa_Customiza", objNF)
    
    colSelecao.Add objNF.iFilialEmpresa
    
    Call Chama_Tela("NFeFedScanLista", colSelecao, objNFeFedScan, objEventoOcorrencia)

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub Ocorrencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataSaida_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub HoraEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub HoraSaida_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Justificativa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO GOTFOCUS DOS CONTROLES - INÍCIO ***
Private Sub Ocorrencia_GotFocus()

    Call MaskEdBox_TrataGotFocus(Ocorrencia, iAlterado)
    
End Sub

Private Sub DataEntrada_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEntrada, iAlterado)
    
End Sub

Private Sub DataSaida_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataSaida, iAlterado)
    
End Sub

Private Sub HoraEntrada_GotFocus()

    Call MaskEdBox_TrataGotFocus(HoraEntrada, iAlterado)
    
End Sub

Private Sub HoraSaida_GotFocus()

    Call MaskEdBox_TrataGotFocus(HoraSaida, iAlterado)
    
End Sub

'*** EVENTO GOTFOCUS DOS CONTROLES - FIM ***

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera as variáveis globais
    Set objEventoOcorrencia = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'*** FUNÇÕES DO SISTEMA DE SETA - INÍCIO ***
Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNFeFedScan As New ClassNFeFedScan

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objNFeFedScan.lOcorrencia = CStr(colCampoValor.Item("Ocorrencia").vValor)

    lErro = CF("NFeFedScan_Le", objNFeFedScan)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207534

    If lErro = SUCESSO Then

        Call Traz_NFeFedScan_Tela(objNFeFedScan)

    End If

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 207534

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207535)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim lErro As Long
Dim objNFeFedScan As New ClassNFeFedScan
Dim objNF As New ClassNFiscal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada a tela
    sTabela = "NFeFedScan"

    lErro = Move_Tela_Memoria(objNFeFedScan)
    If lErro <> SUCESSO Then gError 207536

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Ocorrencia", objNFeFedScan.lOcorrencia, 0, "Ocorrencia"
    
    objNF.iFilialEmpresa = giFilialEmpresa
    lErro = CF("NFiscal_FilialEmpresa_Customiza", objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objNF.iFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 207536

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 207537)

    End Select

    Exit Sub

End Sub
'*** FUNÇÕES DO SISTEMA DE SETA - FIM ***

'*** FUNÇÕES DO BROWSE - INÍCIO

Private Sub objEventoOcorrencia_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFeFedScan As New ClassNFeFedScan
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoOcorrencia_evSelecao
    
    Set objNFeFedScan = obj1

    Call Traz_NFeFedScan_Tela(objNFeFedScan)
    
    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoOcorrencia_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207533)

    End Select
    
    iAlterado = 0

    Exit Sub

End Sub
'*** FUNÇÕES DO BROWSE - FIM ***

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Ocorrências de Período em Contingência de NFe"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFeFedScan"
    
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

Private Sub ProxNum_Click()

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
'**** fim do trecho a ser copiado *****


Private Sub DataSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataSaida_Validate

    'Verifica se a data de emissao está preenchida
    If Len(Trim(DataSaida.ClipText)) > 0 Then

        'Verifica se a data emissao é válida
        lErro = Data_Critica(DataSaida.Text)
        If lErro <> SUCESSO Then gError 207538

    End If

    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 207538

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207539)

    End Select

    Exit Sub

End Sub

Private Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    'Verifica se a data de emissao está preenchida
    If Len(Trim(DataEntrada.ClipText)) > 0 Then

        'Verifica se a data emissao é válida
        lErro = Data_Critica(DataEntrada.Text)
        If lErro <> SUCESSO Then gError 207540

    End If

    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 207540

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207541)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntrada_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_DownClick

    'Diminui a DataEntrada em 1 dia
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 207542

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case gErr

        Case 207542

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207543)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEntrada_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntrada_UpClick

    'Aumenta a DataEntrada em 1 dia
    lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 207544

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case gErr

        Case 207544

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207545)

    End Select

    Exit Sub

End Sub


Private Sub UpDownSaida_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownSaida_DownClick

    'Diminui a DataSaida em 1 dia
    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 207546

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case gErr

        Case 207546

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206547)

    End Select

    Exit Sub

End Sub

Private Sub UpDownSaida_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownSaida_UpClick

    'Aumenta a DataSaida em 1 dia
    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 207548

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case gErr

        Case 207548

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207549)

    End Select

    Exit Sub

End Sub


Public Sub HoraEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraEntrada_Validate

    'Verifica se a hora de entrada foi digitada
    If Len(Trim(HoraEntrada.ClipText)) > 0 Then

        'Critica a hora digitada
        lErro = Hora_Critica(HoraEntrada.Text)
        If lErro <> SUCESSO Then gError 207550

    End If

    Exit Sub

Erro_HoraEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 207550

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207551)

    End Select

    Exit Sub

End Sub


Public Sub HoraSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraSaida_Validate

    'Verifica se a hora de saida foi digitada
    If Len(Trim(HoraSaida.ClipText)) > 0 Then

        'Critica a data digitada
        lErro = Hora_Critica(HoraSaida.Text)
        If lErro <> SUCESSO Then gError 207552


    End If

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 207552

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207553)

    End Select

    Exit Sub

End Sub

