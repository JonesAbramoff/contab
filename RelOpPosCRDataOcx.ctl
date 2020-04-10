VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPosCRDataOcx 
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   ScaleHeight     =   3315
   ScaleWidth      =   6255
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   90
      TabIndex        =   8
      Top             =   1380
      Width           =   3690
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   795
         TabIndex        =   9
         Top             =   285
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   780
         TabIndex        =   10
         Top             =   765
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label LabelClienteAte 
         Caption         =   "Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LabelClienteDe 
         Caption         =   "Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPosCRDataOcx.ctx":0000
      Left            =   870
      List            =   "RelOpPosCRDataOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   255
      Width           =   2925
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
      Left            =   4125
      Picture         =   "RelOpPosCRDataOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   810
      Width           =   1815
   End
   Begin VB.CheckBox CheckPulaPag 
      Caption         =   "Pula p�gina a cada novo cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   2835
      Width           =   3660
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3930
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPosCRDataOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPosCRDataOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPosCRDataOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPosCRDataOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   1995
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   855
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   855
      TabIndex        =   14
      Top             =   855
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "Op��o:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   16
      Top             =   315
      Width           =   660
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data:"
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
      Height          =   255
      Left            =   135
      TabIndex        =   15
      Top             =   870
      Width           =   555
   End
End
Attribute VB_Name = "RelOpPosCRDataOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giFocoInicial As Boolean
Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1

Private Sub ClienteFinal_Validate(Cancel As Boolean)
'1
Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    'Se est� Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou C�digo)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then Error 47734

    End If

    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47734

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171266)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    'se est� Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou C�digo)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then Error 47735

    End If

    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47735

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171267)

    End Select

End Sub

Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click

    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171268)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click

    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171269)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1

    'Preenche o Cliente Final com o Codigo selecionado
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    'Preenche o Cliente Final com Codigo - Descricao
    Call ClienteFinal_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1

    'Preenche o Cliente Inical com o codigo
    ClienteInicial.Text = CStr(objCliente.lCodigo)

    'Preenche o Cliente Inicial com codigo - Descricao
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Function PreencheComboOpcoes(sCodRel As String) As Long
'preenche o Combo de Op��es com as op��es referentes a sCodRel

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_PreencheComboOpcoes

    'le os nomes das opcoes do relat�rio existentes no BD
    lErro = CF("RelOpcoes_Le_Todos",sCodRel, colRelParametros)
    If lErro <> SUCESSO Then Error 23072

    'preenche o ComboBox com os nomes das op��es do relat�rio
    For Each objRelOpcoes In colRelParametros
        ComboOpcoes.AddItem objRelOpcoes.sNome
    Next

    PreencheComboOpcoes = SUCESSO

    Exit Function

Erro_PreencheComboOpcoes:

    PreencheComboOpcoes = Err

    Select Case Err

        Case 23072

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171270)

    End Select

    Exit Function

End Function

Function Critica_Data_RelOpRazao() As Long
'faz a cr�tica da data

Dim lErro As Long

On Error GoTo Erro_Critica_Data_RelOpRazao

    'data n�o pode ser vazia
    If Len(Data.ClipText) = 0 Then Error 23074
   
    Critica_Data_RelOpRazao = SUCESSO

    Exit Function

Erro_Critica_Data_RelOpRazao:

    Critica_Data_RelOpRazao = Err

    Select Case Err

        Case 23074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
            Data.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171271)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 23083

    'pega Cliente Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then Error 23084

    ClienteInicial.Text = CStr(sParam)

    'pega Cliente Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then Error 23085

    ClienteFinal.Text = CStr(sParam)

    'Pega 'Pula p�gina a cada novo conta' e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPULAPAGQBR0", sParam)
    If lErro <> SUCESSO Then Error 23086

    If sParam = "S" Then CheckPulaPag.Value = 1

    'pega data e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro <> SUCESSO Then Error 23088

    Data.PromptInclude = False
    Data.Text = sParam
    Data.PromptInclude = True

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23083, 23084, 23085, 23086, 23088

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171272)

    End Select

    Exit Function

End Function


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usu�rio

Dim lErro As Long
Dim iPer As Integer
Dim iExercicio As Integer
Dim sCheck As String
Dim sDt As String
Dim sCliente_I As String, sCliente_F As String

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Data_RelOpRazao
    If lErro <> SUCESSO Then Error 23089

    'lErro = Obtem_Periodo_Exercicio(iPer, iExercicio, sDt)
    'If lErro <> SUCESSO Then Error 23090

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23091

    'Pegar parametros da tela
    sCliente_I = ClienteInicial.Text
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then Error 23092

    sCliente_F = ClienteFinal.Text
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then Error 23093

    sCliente_I = LCodigo_Extrai(ClienteInicial.Text)
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then Error 23092

    sCliente_F = LCodigo_Extrai(ClienteFinal.Text)
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then Error 23093

    'Pula P�gina a Cada Novo cliente
    If CheckPulaPag.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If


    lErro = objRelOpcoes.IncluirParametro("TPULAPAGQBR0", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 23094

    
    lErro = objRelOpcoes.IncluirParametro("NPERFIM", CStr(iPer))
    If lErro <> AD_BOOL_TRUE Then Error 23096

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(iExercicio))
    If lErro <> AD_BOOL_TRUE Then Error 23097

    lErro = objRelOpcoes.IncluirParametro("DDATA", Data.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23100

    'Se cliente final preenchido
    If Len(Trim(ClienteFinal.Text)) <> 0 Then

        'Verificar se cliente Final � maior que cliente Inicial
        If LCodigo_Extrai(ClienteFinal.Text) < LCodigo_Extrai(ClienteInicial.Text) Then Error 23101

    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sDt)
    If lErro <> SUCESSO Then Error 23102

    '???Call Acha_Nome_TSK(sDtIni_I)

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 23089, 23090, 23091, 23092, 23093, 23094

        Case 23096, 23097, 23100, 23102

        Case 23101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_FINAL_MENOR", Err, Error$)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171273)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sDt As String) As Long
'monta a express�o de sele��o que ser� incluida dinamicamente para a execucao do relatorio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If ClienteInicial.Text <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(LCodigo_Extrai(ClienteInicial.Text))

    If ClienteFinal.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(LCodigo_Extrai(ClienteFinal.Text))
    End If
    
'    'se a data n�o coincide com o per�odo
'    If Data.Text <> sDt Then
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "LancData <= " & Forprint_ConvData(CDate(Data.Text))
'    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171274)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 24976

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 48773

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 48773

        Case 24976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171275)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 23105

    vbMsgRes = Rotina_Aviso(vbYesNo, "EXCLUSAO_RELOPRAZAOCR")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 23106

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as op��es da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 23105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 23106

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171276)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23107

    Me.Enabled = False
    Call gobjRelatorio.Executar_Prossegue

    Unload Me

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23107

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171277)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de op��o"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then Error 23108

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23109

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 23110

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 59496

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 23109, 23110, 59496

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171278)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    CheckPulaPag.Value = 0

    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

Dim lErro As Long

On Error GoTo Erro_ComboOpcoes_Click

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Le",gobjRelOpcoes)
    If (lErro <> SUCESSO) Then Error 23111

    lErro = PreencherParametrosNaTela(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23112

    Exit Sub

Erro_ComboOpcoes_Click:

    Select Case Err

        Case 23111, 23112

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171279)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim sData As String
Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text
        lErro = Data_Critica(sData)
        If lErro <> SUCESSO Then Error 23113

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case Err

        Case 23113

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171280)

    End Select

    Exit Sub

End Sub


Private Sub Form_Load()

Dim colCodigoDescricao As New AdmCollCodigoNome
Dim lErro As Long, iIndice As Integer
Dim objCodigoDescricao As AdmlCodigoNome

On Error GoTo Erro_OpcoesRel_Form_Load

    giFocoInicial = 1

    Set objEventoClienteInic = New AdmEvento

    Set objEventoClienteFim = New AdmEvento

'    'Preenche combo com as op��es de relat�rio
'    lErro = PreencheComboOpcoes(gobjRelatorio.sCodRel)
'    If lErro <> SUCESSO Then Error 23116
'
'    'verifica se o nome da op��o passada est� no ComboBox
'    For iIndice = 0 To ComboOpcoes.ListCount - 1
'
'        If ComboOpcoes.List(iIndice) = gobjRelOpcoes.sNome Then
'
'            ComboOpcoes.Text = ComboOpcoes.List(iIndice)
'            PreencherParametrosNaTela (gobjRelOpcoes)
'
'            Exit For
'
'        End If
'
'    Next
'
'    'Preenche a listbox clientes
'    'Le cada codigo e Nome Reduzido da tabela clientes
'    lErro = CF("LCod_Nomes_Le","clientes", "Codigo", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 23117
'
'    'preenche a listbox clientes com os objetos da colecao colCodigoDescricao
'    For Each objCodigoDescricao In colCodigoDescricao
'
'        ClientesList.AddItem objCodigoDescricao.sNome
'        ClientesList.ItemData(ClientesList.NewIndex) = objCodigoDescricao.lCodigo
'
'    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23116, 23117

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171281)

    End Select

    Unload Me

    Exit Sub

End Sub

'Private Sub ClienteFinal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objCliente As New ClassCliente
'Dim iCodFilial As Integer
'
'On Error GoTo Erro_ClienteFinal_Validate
'
'    giFocoInicial = 0
'
'    lErro = TP_Cliente_Le(ClienteFinal, objCliente, iCodFilial)
'    If lErro Then Error 23078
'
'    Exit Sub
'
'Erro_ClienteFinal_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 23078
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171282)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub ClienteInicial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objCliente As New ClassCliente
'Dim iCodFilial As Integer
'
'On Error GoTo Erro_ClienteInicial_Validate
'
'    giFocoInicial = 1
'
'    lErro = TP_Cliente_Le(ClienteInicial, objCliente, iCodFilial)
'    If lErro <> SUCESSO Then Error 23079
'
'    Exit Sub
'Erro_ClienteInicial_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 23079
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171283)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ClientesList_DblClick()
'
'Dim sListBoxItem As String
'Dim lErro As Long
'
'On Error GoTo Erro_ClientesList_DblClick
'
'    'Se n�o h� Cliente selecionado sai da rotina
'    If ClientesList.ListIndex = -1 Then Exit Sub
'
'    'Pega o nome reduzido do Cliente na ListBox e joga no Cliente que teve o �ltimo foco
'    sListBoxItem = Trim(ClientesList.List(ClientesList.ListIndex))
'
'    'Verifica se o nome reduzido do Cliente est� vazio
'    If Len(sListBoxItem) = 0 Then Error 23076
'
'    If giFocoInicial = 0 Then
'
'        ClienteFinal.Text = sListBoxItem
'        Exit Sub
'
'    End If
'
'    ClienteInicial.Text = sListBoxItem
'
'    Exit Sub
'
'Erro_ClientesList_DblClick:
'
'    Select Case Err
'
'        Case 23076
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_VAZIO", Err, Error$)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171284)
'
'    End Select
'
'    Exit Sub
'
'End Sub



Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23028

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 23028

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171285)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23062

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 23062

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171286)

    End Select

    Exit Sub

End Sub

'Function Obtem_Periodo_Exercicio(iPer As Integer, iExercicio As Integer, sDt As String) As Long
'
'Dim objPer As New ClassPeriodo
'Dim lErro As Long
'
'On Error GoTo Erro_Obtem_Periodo_Exercicio
'
'    'pega o per�odo da Data
'    lErro = CF("Periodo_Le",CDate(Data.Text), objPer)
'    If lErro <> SUCESSO Then Error 23081
'
'    iPer = objPer.iPeriodo
'    iExercicio = objPer.iExercicio
'
'    sDtFim = objPer.dtData
'
'    Obtem_Periodo_Exercicio = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_Periodo_Exercicio:
'
'    Obtem_Periodo_Exercicio = Err
'
'    Select Case Err
'
'        Case 23081
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171287)
'
'    End Select
'
'    Exit Function
'
'End Function


Private Sub Form_Unload(Cancel As Integer)

    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data)

End Sub


Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSFORN
    Set Form_Load_Ocx = Me
    Caption = "Posi��o Cont�bil de Clientes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpRazaoCR"

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


