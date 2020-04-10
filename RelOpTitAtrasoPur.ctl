VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpTitAtraso 
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   LockControls    =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   6900
   Begin VB.Frame Frame5 
      Caption         =   "Formas de Pagamento"
      Height          =   1755
      Left            =   60
      TabIndex        =   24
      Top             =   2820
      Width           =   6780
      Begin VB.ListBox FormasPagto 
         Height          =   1410
         Left            =   75
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   240
         Width           =   4290
      End
      Begin VB.CommandButton BotaoMarcar 
         Caption         =   "Marcar Todas"
         Height          =   525
         Left            =   4350
         Picture         =   "RelOpTitAtrasoPur.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   255
         Width           =   1530
      End
      Begin VB.CommandButton BotaoDesmarcar 
         Caption         =   "Desmarcar Todas"
         Height          =   525
         Left            =   4350
         Picture         =   "RelOpTitAtrasoPur.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   900
         Width           =   1530
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vencimento"
      Height          =   705
      Left            =   60
      TabIndex        =   19
      Top             =   15
      Width           =   6780
      Begin MSComCtl2.UpDown UpDownVencimentoDe 
         Height          =   315
         Left            =   1755
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VencimentoDe 
         Height          =   285
         Left            =   600
         TabIndex        =   0
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownVencimentoAte 
         Height          =   315
         Left            =   5085
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VencimentoAte 
         Height          =   285
         Left            =   3930
         TabIndex        =   1
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
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
         Height          =   195
         Left            =   3540
         TabIndex        =   23
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label2 
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
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Layout"
      Height          =   495
      Left            =   60
      TabIndex        =   18
      Top             =   2295
      Width           =   6780
      Begin VB.CheckBox DetalhadoHC 
         Caption         =   "Detalhado com histórico de cobrança"
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
         Left            =   150
         TabIndex        =   7
         Top             =   225
         Width           =   3585
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vendedores"
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   1635
      Width           =   6780
      Begin VB.OptionButton OptVendIndir 
         Caption         =   "Vendas Indiretas"
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
         Left            =   1800
         TabIndex        =   5
         Top             =   180
         Width           =   1800
      End
      Begin VB.OptionButton OptVendDir 
         Caption         =   "Vendas Diretas"
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
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   1800
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   4545
         TabIndex        =   6
         Top             =   210
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   3630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   600
      Left            =   3450
      Picture         =   "RelOpTitAtrasoPur.ctx":21FC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Fechar"
      Top             =   4680
      Width           =   1575
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
      Left            =   1785
      Picture         =   "RelOpTitAtrasoPur.ctx":237A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4695
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   690
      Left            =   60
      TabIndex        =   13
      Top             =   765
      Width           =   6780
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   615
         TabIndex        =   2
         Top             =   255
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3930
         TabIndex        =   3
         Top             =   225
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
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
         Left            =   3495
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   300
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   285
         Width           =   315
      End
   End
End
Attribute VB_Name = "RelOpTitAtraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 90608
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 90608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173376)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 90609

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 90609
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173377)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 90610

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 90610
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173378)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long, iIndice As Integer
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Form_Load
    
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    Set colCodigoDescricao = New AdmColCodigoNome
    lErro = CF("FormasPagamento_Le_CodNome", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 16391
    
    'Preenche ListBox Condições com DescReduzidas de CondicoesPagto
    iIndice = -1
    For Each objCodigoDescricao In colCodigoDescricao
        iIndice = iIndice + 1
        FormasPagto.AddItem objCodigoDescricao.sNome
        FormasPagto.ItemData(FormasPagto.NewIndex) = objCodigoDescricao.iCodigo
        FormasPagto.Selected(iIndice) = True
    Next
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173379)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set objEventoVendedor = Nothing
    
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

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173380)

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

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173381)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteInicial.Text = CStr(objCliente.lCodigo)
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 90611
    
    If DetalhadoHC.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "TITATRHC"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 90611

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173382)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutando As Boolean = False) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sClienteIni As String
Dim sClienteFim As String
Dim iVendedor As Integer
'Dim objVendedor As New ClassVendedor
Dim iTipoVend As Integer
Dim lNumIntRel As Long
Dim iIndice As Integer, iAchou As Integer

On Error GoTo Erro_PreencherRelOp

    If Len(Trim(ClienteInicial.Text)) <> 0 And Len(Trim(ClienteFinal.Text)) <> 0 Then
        If (Codigo_Extrai(ClienteInicial.Text)) > (Codigo_Extrai(ClienteFinal.Text)) Then gError 90612
    End If
       
    sClienteIni = SCodigo_Extrai(ClienteInicial.Text)
    sClienteFim = SCodigo_Extrai(ClienteFinal.Text)
    
    For iIndice = 0 To FormasPagto.ListCount - 1
        If FormasPagto.Selected(iIndice) = True Then
            iAchou = 1
            Exit For
        End If
        
    Next
       
    If iAchou = 0 Then gError 207095
    
    If OptVendDir.Value Then
        iTipoVend = VENDEDOR_DIRETO
    Else
        iTipoVend = VENDEDOR_INDIRETO
    End If
    
'    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
'
'    'Verifica se vendedor existe
'    If objVendedor.sNomeReduzido <> "" Then
'
'        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
'        If lErro <> SUCESSO And lErro <> 25008 Then gError ERRO_SEM_MENSAGEM
'
'        iVendedor = objVendedor.iCodigo
'
'    End If
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 90613
         
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sClienteIni)
    If lErro <> AD_BOOL_TRUE Then gError 90614
         
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90615
    
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sClienteFim)
    If lErro <> AD_BOOL_TRUE Then gError 90616
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90617
    
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(iVendedor))
    If lErro <> AD_BOOL_TRUE Then gError 90617
    
    lErro = objRelOpcoes.IncluirParametro("DDATA_HOJE", CStr(gdtDataAtual))
    If lErro <> AD_BOOL_TRUE Then gError 90617
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", Vendedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", Codigo_Extrai(Vendedor.Text))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NTIPOVEND", CStr(iTipoVend))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If VencimentoDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", VencimentoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    If VencimentoAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", VencimentoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        lErro = CF("RelOpFatProdVend_Prepara", iTipoVend, Codigo_Extrai(Vendedor.Text), lNumIntRel)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sClienteIni, sClienteFim)
    If lErro <> SUCESSO Then gError 90618
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr
    
        Case 90612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)

        Case 90613 To 90618
        
        Case ERRO_SEM_MENSAGEM
        
        Case 207095
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_FORMAPAGTO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173383)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sClienteIni As String, sClienteFim As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim sSub As String, iCount As Integer, iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

   If sClienteIni <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(CLng(sClienteIni))

   If sClienteFim <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sClienteFim))

    End If
        
    If giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> gobjCR.iFilialCentralizadora Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(CInt(giFilialEmpresa))
    End If
    
    sSub = ""
    iCount = 0
    For iIndice = 0 To FormasPagto.ListCount - 1
        If FormasPagto.Selected(iIndice) Then
            iCount = iCount + 1
            If sSub <> "" Then sSub = sSub & " OU "
            sSub = sSub & " FormPagto = " & Forprint_ConvInt(FormasPagto.ItemData(iIndice))
        End If
    Next
    
    'Se selecionou só alguns
    If Len(Trim(sSub)) <> 0 And iCount <> FormasPagto.ListCount Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "(" & sSub & ")"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173384)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CAD_CLI
    Set Form_Load_Ocx = Me
    Caption = "Títulos em Atraso"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitAtraso"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        End If
    
    End If

End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub


Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169098)

    End Select

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
    
    'Preenche com o Vendedor da tela
    objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub


Private Sub VencimentoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VencimentoAte)

End Sub

Private Sub VencimentoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VencimentoDe)

End Sub

Private Sub VencimentoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencimentoAte_Validate

    If Len(VencimentoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(VencimentoAte.Text)
        If lErro <> SUCESSO Then Error 47789

    End If

    Exit Sub

Erro_VencimentoAte_Validate:

    Cancel = True


    Select Case Err

        Case 47789

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173518)

    End Select

    Exit Sub

End Sub

Private Sub VencimentoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencimentoDe_Validate

    If Len(VencimentoDe.ClipText) > 0 Then

        lErro = Data_Critica(VencimentoDe.Text)
        If lErro <> SUCESSO Then Error 47790

    End If

    Exit Sub

Erro_VencimentoDe_Validate:

    Cancel = True


    Select Case Err

        Case 47790

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173519)

    End Select

    Exit Sub

End Sub

    
Private Sub UpDownVencimentoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoDe_DownClick

    lErro = Data_Up_Down_Click(VencimentoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47852

    Exit Sub

Erro_UpDownVencimentoDe_DownClick:

    Select Case Err

        Case 47852
            VencimentoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173523)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimentoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoDe_UpClick

    lErro = Data_Up_Down_Click(VencimentoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47853

    Exit Sub

Erro_UpDownVencimentoDe_UpClick:

    Select Case Err

        Case 47853
            VencimentoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173524)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownVencimentoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoAte_DownClick

    lErro = Data_Up_Down_Click(VencimentoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47854

    Exit Sub

Erro_UpDownVencimentoAte_DownClick:

    Select Case Err

        Case 47854
            VencimentoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173525)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimentoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoAte_UpClick

    lErro = Data_Up_Down_Click(VencimentoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47855

    Exit Sub

Erro_UpDownVencimentoAte_UpClick:

    Select Case Err

        Case 47855
            VencimentoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173526)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcar_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To FormasPagto.ListCount - 1
        FormasPagto.Selected(iIndice) = True
    Next

End Sub

Private Sub BotaoDesmarcar_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To FormasPagto.ListCount - 1
        FormasPagto.Selected(iIndice) = False
    Next

End Sub

Sub Limpa_FormasPagto()

Dim iIndice As Integer

    For iIndice = 0 To FormasPagto.ListCount - 1
        FormasPagto.Selected(iIndice) = False
    Next

End Sub
