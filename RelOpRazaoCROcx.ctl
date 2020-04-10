VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRazaoCROcx 
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   LockControls    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   6315
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRazaoCROcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRazaoCROcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRazaoCROcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRazaoCROcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox CheckPulaPag 
      Caption         =   "Pula página a cada novo cliente"
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
      TabIndex        =   5
      Top             =   3255
      Width           =   3660
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
      Left            =   4155
      Picture         =   "RelOpRazaoCROcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   810
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRazaoCROcx.ctx":0A96
      Left            =   900
      List            =   "RelOpRazaoCROcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2925
   End
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
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3690
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   795
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   765
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
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
         TabIndex        =   14
         Top             =   360
         Width           =   615
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
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2475
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   840
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   2475
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1365
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   315
      Left            =   1335
      TabIndex        =   2
      Top             =   1365
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data Inicial :"
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
      Left            =   165
      TabIndex        =   19
      Top             =   885
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data Final :"
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
      Left            =   165
      TabIndex        =   18
      Top             =   1410
      Width           =   1125
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
      Height          =   255
      Left            =   165
      TabIndex        =   17
      Top             =   315
      Width           =   660
   End
End
Attribute VB_Name = "RelOpRazaoCROcx"
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

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate
    
    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then Error 47734

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47734
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 172051)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate
    
    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then Error 47735

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47735
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 172052)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172053)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172054)

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
'preenche o Combo de Opções com as opções referentes a sCodRel

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_PreencheComboOpcoes

    'le os nomes das opcoes do relatório existentes no BD
    lErro = CF("RelOpcoes_Le_Todos",sCodRel, colRelParametros)
    If lErro <> SUCESSO Then Error 23072

    'preenche o ComboBox com os nomes das opções do relatório
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172055)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpRazao() As Long
'faz a crítica da data inicial e da data final

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_RelOpRazao

    'data inicial não pode ser vazia
    If Len(DataInicial.ClipText) = 0 Then Error 23073

    'data final não pode ser vazia
    If Len(DataFinal.ClipText) = 0 Then Error 23074

    'data inicial não pode ser maior que a data final
    If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 23075

    Critica_Datas_RelOpRazao = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpRazao:

    Critica_Datas_RelOpRazao = Err

    Select Case Err

        Case 23073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
            DataInicial.SetFocus

        Case 23074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
            DataFinal.SetFocus

        Case 23075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172056)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros de uma opcao salva anteriormente e exibe na tela

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

    'pega 'Pula página a cada novo conta' e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TPULAPAGQBR0", sParam)
    If lErro <> SUCESSO Then Error 23086

    If sParam = "S" Then CheckPulaPag.Value = 1

    'pega data inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 23087

    DataInicial.PromptInclude = False
    DataInicial.Text = sParam
    DataInicial.PromptInclude = True

    'pega data final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 23088

    DataFinal.PromptInclude = False
    DataFinal.Text = sParam
    DataFinal.PromptInclude = True

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23083, 23084, 23085, 23086, 23087, 23088

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172057)

    End Select

    Exit Function

End Function


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bGeraArqTemp As Boolean = False) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iPer_I As Integer, iPer_F As Integer
Dim iExercicio As Integer, lNumIntRel As Long
Dim sCheck As String, lCliInic As Long, lCliFinal As Long
Dim sDtIni_I As String, sDtFim_F As String
Dim sCliente_I As String, sCliente_F As String

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpRazao
    If lErro <> SUCESSO Then Error 23089

    lErro = Obtem_Periodo_Exercicio(iPer_I, iPer_F, iExercicio, sDtIni_I, sDtFim_F)
    If lErro <> SUCESSO Then Error 23090

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23091

    'Pegar parametros da tela
    sCliente_I = ClienteInicial.Text
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then Error 23092

    sCliente_F = ClienteFinal.Text
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then Error 23093

    lCliInic = LCodigo_Extrai(ClienteInicial.Text)
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", CStr(lCliInic))
    If lErro <> AD_BOOL_TRUE Then Error 23092

    lCliFinal = LCodigo_Extrai(ClienteFinal.Text)
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", CStr(lCliFinal))
    If lErro <> AD_BOOL_TRUE Then Error 23093

    'Pula Página a Cada Novo cliente
    If CheckPulaPag.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If

    lErro = objRelOpcoes.IncluirParametro("TPULAPAGQBR0", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 23094

    lErro = objRelOpcoes.IncluirParametro("NPERINIC", CStr(iPer_I))
    If lErro <> AD_BOOL_TRUE Then Error 23095

    lErro = objRelOpcoes.IncluirParametro("NPERFIM", CStr(iPer_F))
    If lErro <> AD_BOOL_TRUE Then Error 23096

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(iExercicio))
    If lErro <> AD_BOOL_TRUE Then Error 23097

    lErro = objRelOpcoes.IncluirParametro("DINICPERINI", sDtIni_I)
    If lErro <> AD_BOOL_TRUE Then Error 23098

    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23099

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23100

    'Se cliente final preenchido
    If Len(Trim(ClienteFinal.Text)) <> 0 Then

        'Verificar se cliente Final é menor que cliente Inicial
        If lCliFinal < lCliInic Then Error 23101

    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sDtIni_I, sDtFim_F)
    If lErro <> SUCESSO Then Error 23102

    '???Call Acha_Nome_TSK(sDtIni_I)

    If bGeraArqTemp Then
    
        GL_objMDIForm.MousePointer = vbHourglass
        lErro = CF("RelCliSaldo_Prepara",giFilialEmpresa, lNumIntRel, lCliInic, lCliFinal, CDate(sDtIni_I))
        GL_objMDIForm.MousePointer = vbDefault
        If lErro <> SUCESSO Then Error 23102
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then Error 23102

    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23089, 23090, 23091, 23092, 23093, 23094, 23095

        Case 23096, 23097, 23098, 23099, 23100, 23102

        Case 23101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_FINAL_MENOR", Err, Error$)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172058)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sDtIni_I As String, sDtFim_F As String) As Long
'monta a expressão de seleção que será incluida dinamicamente para a execucao do relatorio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If ClienteInicial.Text <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(LCodigo_Extrai(ClienteInicial.Text))

    If ClienteFinal.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(LCodigo_Extrai(ClienteFinal.Text))
    End If

    If sExpressao <> "" Then sExpressao = sExpressao & " E "
    sExpressao = sExpressao & "LancData >= " & Forprint_ConvData(CDate(DataInicial.Text))

    If sExpressao <> "" Then sExpressao = sExpressao & " E "
    sExpressao = sExpressao & "LancData <= " & Forprint_ConvData(CDate(DataFinal.Text))

    If giFilialEmpresa <> EMPRESA_TODA And gobjCTB.giContabCentralizada = 0 Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaLcto = " & Forprint_ConvInt(giFilialEmpresa)
    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = "TipoReg = 1 OU (TipoReg = 2 E " & sExpressao & ")"

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172059)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172060)

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

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172061)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then Error 23107

    Me.Enabled = False
    Call gobjRelatorio.Executar_Prossegue

    Unload Me

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23107

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172062)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de opção"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172063)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172064)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 23113

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case Err

        Case 23113

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172065)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 23115

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 23115

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172066)

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

'    'Preenche combo com as opções de relatório
'    lErro = PreencheComboOpcoes(gobjRelatorio.sCodRel)
'    If lErro <> SUCESSO Then Error 23116
'
'    'verifica se o nome da opção passada está no ComboBox
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172067)

    End Select

    Unload Me

    Exit Sub

End Sub
'
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
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172068)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
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
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172069)
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
'    'Se não há Cliente selecionado sai da rotina
'    If ClientesList.ListIndex = -1 Then Exit Sub
'
'    'Pega o nome reduzido do Cliente na ListBox e joga no Cliente que teve o último foco
'    sListBoxItem = Trim(ClientesList.List(ClientesList.ListIndex))
'
'    'Verifica se o nome reduzido do Cliente está vazio
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
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172070)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23029

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 23029

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172071)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23063

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 23063

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172072)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23028

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 23028

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172073)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23062

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 23062

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172074)

    End Select

    Exit Sub

End Sub

Function Obtem_Periodo_Exercicio(iPer_I As Integer, iPer_F As Integer, iExercicio As Integer, sDtIni_I As String, sDtFim_F As String) As Long
'a partir das datas ( inicial e final ) encontra o período e o exercício
'as datas devem estar no mesmo exercício
'pega também a data inicial do período inicial e a data final do período final

Dim objPer_I As New ClassPeriodo, objPer_F As New ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_Obtem_Periodo_Exercicio

    'pega o período da Data Inicial
    lErro = CF("Periodo_Le",CDate(DataInicial.Text), objPer_I)
    If lErro <> SUCESSO Then Error 23080

    'pega o período da Data Final
    lErro = CF("Periodo_Le",CDate(DataFinal.Text), objPer_F)
    If lErro <> SUCESSO Then Error 23081

    'Data Inicial e Final devem estar num mesmo exercício
    If objPer_I.iExercicio <> objPer_F.iExercicio Then Error 23082

    iPer_I = objPer_I.iPeriodo
    iPer_F = objPer_F.iPeriodo
    iExercicio = objPer_I.iExercicio
    sDtIni_I = objPer_I.dtDataInicio
    sDtFim_F = objPer_I.dtDataFim

    Obtem_Periodo_Exercicio = SUCESSO

    Exit Function

Erro_Obtem_Periodo_Exercicio:

    Obtem_Periodo_Exercicio = Err

    Select Case Err

        Case 23080

        Case 23081

        Case 23082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_COM_EXERCICIOS_DIFERENTES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 172075)

    End Select

    Exit Function

End Function


Private Sub Form_Unload(Cancel As Integer)

    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Private Sub DataFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataInicial)

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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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

    'Parent.HelpContextID = IDH_RELOP_POSFORN
    Set Form_Load_Ocx = Me
    Caption = "Razão Auxiliar de Contas a Receber"
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


Function RelCliSaldo_Prepara(iFilialEmpresa As Integer, lNumIntRel As Long, lCliInic As Long, lCliFinal As Long, dtDataSaldo As Date) As Long
'Insere registros na tabela RelCliSaldo com os saldos anteriores dos clientes que serao necessarios p/execucao de relatorio

Dim lErro As Long, dtData As Date
Dim lTransacao As Long, alComando(0 To 1) As Long, iIndice As Integer, lCliAnterior As Long, dSaldo As Double
Dim iTipoLcto As Integer, lCliente As Long, dValorTotal As Double, dValorIRRF As Double, dValorINSS As Double, iINSSRetido As Integer, dBaixasParcRec_ValorBaixado As Double, iBaixasParcRec_Status As Integer, dtBaixasParcRec_DataCancelamento As Date

On Error GoTo Erro_RelCliSaldo_Prepara

    dtData = dtDataSaldo - 1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81801
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81802

    'obtem numintrel
    lErro = CF("Config_ObterNumInt","CRConfig", "NUM_PROX_REL_CLI_SALDO", lNumIntRel)
    If lErro <> SUCESSO Then gError 81803
    
    If iFilialEmpresa = EMPRESA_TODA Or gobjCTB.giContabCentralizada <> 0 Then
        If lCliInic <> 0 Or lCliFinal <> 0 Then
            lErro = Comando_Executar(alComando(0), "SELECT TipoLcto, Cliente, Valor, ValorIRRF, ValorINSS, INSSRetido, BaixasParcRec_ValorBaixado, BaixasParcRec_Status, BaixasParcRec_DataCancelamento FROM PosCRDataLctos WHERE Cliente >= ? AND Cliente <= ? AND ((TipoLcto IN (1,3) AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?)) OR (TipoLcto IN (2,4) AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?) AND DataContabilBaixa <= ? AND (BaixasParcRec_DataCancelamento = ? OR BaixasParcRec_DataCancelamento > ? ))) ORDER BY Cliente", _
                iTipoLcto, lCliente, dValorTotal, dValorIRRF, dValorINSS, iINSSRetido, dBaixasParcRec_ValorBaixado, iBaixasParcRec_Status, dtBaixasParcRec_DataCancelamento, lCliInic, lCliFinal, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData)
        Else
            lErro = Comando_Executar(alComando(0), "SELECT TipoLcto, Cliente, Valor, ValorIRRF, ValorINSS, INSSRetido, BaixasParcRec_ValorBaixado, BaixasParcRec_Status, BaixasParcRec_DataCancelamento FROM PosCRDataLctos WHERE (TipoLcto IN (1,3) AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?)) OR (TipoLcto IN (2,4) AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?) AND DataContabilBaixa <= ? AND (BaixasParcRec_DataCancelamento = ? OR BaixasParcRec_DataCancelamento > ? )) ORDER BY Cliente", _
                iTipoLcto, lCliente, dValorTotal, dValorIRRF, dValorINSS, iINSSRetido, dBaixasParcRec_ValorBaixado, iBaixasParcRec_Status, dtBaixasParcRec_DataCancelamento, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData)
        End If
    Else
        If lCliInic <> 0 Or lCliFinal <> 0 Then
            lErro = Comando_Executar(alComando(0), "SELECT TipoLcto, Cliente, Valor, ValorIRRF, ValorINSS, INSSRetido, BaixasParcRec_ValorBaixado, BaixasParcRec_Status, BaixasParcRec_DataCancelamento FROM PosCRDataLctos WHERE Cliente >= ? AND Cliente <= ? AND ((TipoLcto IN (1,3) AND DocFilEmp = ? AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?)) OR (TipoLcto IN (2,4) AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?) AND DataContabilBaixa <= ? AND (BaixasParcRec_DataCancelamento = ? OR BaixasParcRec_DataCancelamento > ? ))) ORDER BY Cliente", _
                iTipoLcto, lCliente, dValorTotal, dValorIRRF, dValorINSS, iINSSRetido, dBaixasParcRec_ValorBaixado, iBaixasParcRec_Status, dtBaixasParcRec_DataCancelamento, lCliInic, lCliFinal, iFilialEmpresa, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData)
        Else
            lErro = Comando_Executar(alComando(0), "SELECT TipoLcto, Cliente, Valor, ValorIRRF, ValorINSS, INSSRetido, BaixasParcRec_ValorBaixado, BaixasParcRec_Status, BaixasParcRec_DataCancelamento FROM PosCRDataLctos WHERE (TipoLcto IN (1,3) AND DocFilEmp = ? AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?)) OR (TipoLcto IN (2,4) AND DataEmissao<=? AND (DataBaixa = ? OR DataBaixa > ?) AND DataContabilBaixa <= ? AND (BaixasParcRec_DataCancelamento = ? OR BaixasParcRec_DataCancelamento > ? )) ORDER BY Cliente", _
                iTipoLcto, lCliente, dValorTotal, dValorIRRF, dValorINSS, iINSSRetido, dBaixasParcRec_ValorBaixado, iBaixasParcRec_Status, dtBaixasParcRec_DataCancelamento, iFilialEmpresa, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData, dtData, DATA_NULA, dtData)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 81804
        
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81805
    
    lCliAnterior = -1
    
    Do While lErro = AD_SQL_SUCESSO
    
        If lCliAnterior = -1 Then lCliAnterior = lCliente
        
        'se trocou de Cliente
        If lCliAnterior <> lCliente Then
    
            'insere registro em RelCliSaldo
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelCliSaldo ( NumIntRel, Cliente, Saldo ) VALUES (?,?,?)", lNumIntRel, lCliAnterior, dSaldo)
            If lErro <> AD_SQL_SUCESSO Then gError 81807
            
            lCliAnterior = lCliente
            dSaldo = 0
        
        End If
        
        Select Case iTipoLcto
        
            Case 1, 3 'titulos
                dSaldo = Round(dSaldo + dValorTotal - dValorIRRF - IIf(iINSSRetido <> 0, dValorINSS, 0), 2)
            
            Case Else 'baixas
                dSaldo = Round(dSaldo - dBaixasParcRec_ValorBaixado)
                
        End Select
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81806
    
    Loop
    
    If lCliAnterior <> -1 Then
            
        'insere registro em RelCliSaldo
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelCliSaldo ( NumIntRel, Cliente, Saldo ) VALUES (?,?,?)", lNumIntRel, lCliAnterior, dSaldo)
        If lErro <> AD_SQL_SUCESSO Then gError 81808
            
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81809
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelCliSaldo_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelCliSaldo_Prepara:

    RelCliSaldo_Prepara = gErr
     
    Select Case gErr
          
        Case 81803
        
        Case 81801
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 81802
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 81804, 81805, 81806
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POSCRDATALCTOS", gErr)
        
        Case 81807, 81808
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCLISALDO", gErr)
        
        Case 81809
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172076)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function


