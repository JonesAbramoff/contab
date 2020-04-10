VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpMovCaixaCentral 
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   ScaleHeight     =   1830
   ScaleWidth      =   6960
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMovCaixaCentral.ctx":0000
      Left            =   1200
      List            =   "RelOpMovCaixaCentral.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   2670
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   240
      TabIndex        =   9
      Top             =   885
      Width           =   4035
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   690
         TabIndex        =   2
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   2685
         TabIndex        =   3
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   300
         TabIndex        =   13
         Top             =   300
         Width           =   315
      End
      Begin VB.Label LabelDataAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   2265
         TabIndex        =   12
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4560
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   150
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpMovCaixaCentral.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpMovCaixaCentral.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpMovCaixaCentral.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1680
         Picture         =   "RelOpMovCaixaCentral.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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
      Left            =   4920
      Picture         =   "RelOpMovCaixaCentral.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   975
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMovCaixaCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Movimentação do Caixa Central"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpMovCaixaCentral"

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
    'Parent.UnloadDoFilho

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

'Inicio Tela de Relatório de Movtos Caixa
'Dia 26/11/02
'sergio Ricardo
'Supervisor Shirley

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 113130

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 113131

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 113131

        Case 113130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170016)

    End Select

    Exit Function

End Function

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170017)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 113132

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113132

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170018)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 113133

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 113133
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170019)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 113134

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 113134
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170020)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 113135

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 113135

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170021)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 113136

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 113136
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170022)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 113137

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 113137
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170023)

    End Select

    Exit Sub

End Sub



Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click
    
    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 113142

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 113143

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 113144

    'se a opção de relatório foi gravada em RelatorioOpcoes então adcionar a opção de relatório na comboopções
    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 113142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 113143, 113144

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170024)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreenchgerrelOp
    
    lErro = Data_Critica_Parametros()
    If lErro <> SUCESSO Then gError 113145

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 113146


    If Len(Trim(DataDe.ClipText)) <> 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataDe.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113147
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 113147
    End If
    
    If Len(Trim(DataAte.Text)) <> 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataAte.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113148
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 113148
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, StrParaDate(DataDe.Text), StrParaDate(DataAte.Text))
    If lErro <> SUCESSO Then gError 113149

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreenchgerrelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 113145 To 113149

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170025)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Function Data_Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Data_Critica_Parametros

    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then

         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 113150

    End If

    Data_Critica_Parametros = SUCESSO

    Exit Function

Erro_Data_Critica_Parametros:

    Data_Critica_Parametros = gErr

    Select Case gErr

        Case 113150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170026)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, dtdataDe As Date, dtdataAte As Date) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If dtdataDe <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "DataMovimento >= " & Forprint_ConvData(dtdataDe)

    End If

    If dtdataAte <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataMovimento <= " & Forprint_ConvData(dtdataAte)

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170027)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 113151

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 113151

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170028)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 113153

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 113154

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 113153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 113154

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170029)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objProdutoRankingTela As New ClassProdutosRankingTela
Dim colProdutosRanking As New Collection

'Tirar depois
Dim dtDataIn As Date
Dim dSaldoInicial As Double

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 113155
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 113155

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170030)

    End Select

    Exit Sub

End Sub

'Função para ser Executada no Inicio do Relatório ...

'Function Obtem_MvCaixaCentral_SldInicial(dtDataIn As Date, dSaldoInicial As Double, iCodCaixa As Integer, iFilialEmpresa As Integer) As Long
''Função que Obtem po Saldo a partir da data inicial do Relatorio
'
'Dim lErro As Long
'Dim lTransacao As Long
'Dim sSQLDeb As String
'Dim sSQLCred As String
'Dim alComando(1) As Long
'Dim dtdataAux As Date
'Dim iMesAux As Integer
'Dim iAnoAux As Integer
'Dim iIndice As Integer
'Dim Deb(1 To 12) As Double
'Dim Cred(1 To 12) As Double
'Dim dDebMem As Double
'Dim dCredMem As Double
'Dim iCodCaixaAux As Integer
'Dim iFilialAux As Integer
'
'On Error GoTo Erro_Obtem_MvCaixaCentral_SldInicial
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        alComando(iIndice) = Comando_Abrir
'        If alComando(iIndice) = 0 Then gError 113156
'
'    Next
'
'    'Verifica o Ano para Adquirir o saldo
'    iAnoAux = Year(dtDataIn)
'
'    'Selecionar o saldo aglutinado de Movimentos de Caixa para o ano q passou
'    lErro = Comando_Executar(alComando(0), "SELECT  SaldoInicial FROM CCMov WHERE FilialEmpresa = ? AND Ano = ? AND CodCaixa = ? ", dSaldoInicial, iFilialEmpresa, iAnoAux, iCodCaixa)
'    If lErro <> AD_SQL_SUCESSO Then gError 113163
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113158
'
'    iIndice = 1
'
'    'Cria Uma String Com todos os Meses do ano
'    Do While iIndice <= 12
'
'        sSQLDeb = sSQLDeb & " Deb0" & iIndice & " +"
'        sSQLCred = sSQLCred & " Cred0" & iIndice & " +"
'        iIndice = iIndice + 1
'        If iIndice > 9 Then
'            sSQLDeb = sSQLDeb & " Deb10 + "
'            sSQLCred = sSQLCred & " Cred10 + "
'            sSQLDeb = sSQLDeb & " Deb11 + "
'            sSQLCred = sSQLCred & " Cred11  + "
'            sSQLDeb = sSQLDeb & " Deb12 +"
'            sSQLCred = sSQLCred & " Cred12 +"
'            Exit Do
'        End If
'    Loop
'
'    sSQLDeb = Left(sSQLDeb, Len(sSQLDeb) - 1)
'    sSQLCred = Left(sSQLCred, Len(sSQLCred) - 1)
'
'    'Selecionar o saldo aglutinado de Movimentos de Caixa para o ano q passou
'    lErro = Comando_Executar(alComando(1), "SELECT CodCaixa , FilialEmpresa , SUM(" & sSQLDeb & "), SUM (" & sSQLCred & ") FROM CCMov WHERE FilialEmpresa = ? AND Ano = ? AND CodCaixa = ? GROUP BY (CodCaixa) , (FilialEmpresa), ( SaldoInicial)", iCodCaixaAux, iFilialAux, dDebMem, dCredMem, iFilialEmpresa, iAnoAux, iCodCaixa)
'    If lErro <> AD_SQL_SUCESSO Then gError 113164
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113165
'
'    'Atualiza o Saldo de Movimentos até o mês anterior
'    dSaldoInicial = dSaldoInicial + dDebMem + dCredMem
'
'    'fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        Call Comando_Fechar(alComando(iIndice))
'
'    Next
'
'    Obtem_MvCaixaCentral_SldInicial = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_MvCaixaCentral_SldInicial:
'
'    Obtem_MvCaixaCentral_SldInicial = gErr
'
'    Select Case gErr
'
'        Case 113156
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 113158, 113163, 113164, 113165
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SELECAO_CCMOV", gErr, iCodCaixa, iAnoAux, iFilialEmpresa)
'
'       Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170031)
'
'    End Select
'
'    'fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'
'        Call Comando_Fechar(alComando(iIndice))
'
'    Next
'
'    Exit Function
'
'End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub


Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    'Função que lê no Banco de dados o Codigo do Relatorio e Traz a Coleção de parâmetro carregados
    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 113160

    'DdataDe
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 113161
    Call DateParaMasked(DataDe, StrParaDate(sParam))
    
    'pega a DataFinal e exibe e valida
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 113162
    Call DateParaMasked(DataAte, StrParaDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 113160 To 113162

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170032)

    End Select

    Exit Function

End Function

'Função de Chamada em AdrelVb ...

'Function Obtem_MvCaixaCentral_SldInicialPrimeiro(dtDataIn As Date, dSaldoInicial As Double, iCodCaixa As Integer, iFilialEmpresa As Integer) As Long
'
'Dim lErro As Long
'Dim ObjRelMovCxCentral As ClassRelMovCxCentral
'
'On Error GoTo Erro_Obtem_MvCaixaCentral_SldInicialPrimeiro
'
'    Set ObjSaldoCaixaCentral = ObtemObj("RelMovCxCentral")
'    If ObjSaldoCaixaCentral Is Nothing Then
'
'        Set ObjSaldoCaixaCentral = New ClassRelMovCxCentral
'        lErro = GuardaObj("RelMovCxCentral", ObjRelMovCxCentral)
'        If lErro <> SUCESSO Then gError 113166
'
'    End If
'
'    Obtem_MvCaixaCentral_SldInicialPrimeiro = ObjRelMovCxCentral.Obtem_MvCaixaCentral_SldInicial(dtDataIn, dSaldoInicial, iCodCaixa, iFilialEmpresa)
'
'    Exit Function
'
'Erro_Obtem_MvCaixaCentral_SldInicialPrimeiro:
'
'    Obtem_MvCaixaCentral_SldInicialPrimeiro = gErr
'
'    Select Case gErr
'
'        Case 113166
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170033)
'
'    End Select
'
'End Function

'Fim da Chamada para a Função em AdRelVb ....
