VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpDemApInv_LOcx 
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   ScaleHeight     =   2790
   ScaleWidth      =   8265
   Begin VB.ComboBox ComboTotaliza 
      Height          =   315
      ItemData        =   "RelOpDemApInv_LOcx.ctx":0000
      Left            =   1530
      List            =   "RelOpDemApInv_LOcx.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1320
      Width           =   2280
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   1425
      ItemData        =   "RelOpDemApInv_LOcx.ctx":003A
      Left            =   5640
      List            =   "RelOpDemApInv_LOcx.ctx":003C
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Almoxarifados"
      Height          =   840
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   5295
      Begin MSMask.MaskEdBox AlmoxarifadoInicial 
         Height          =   315
         Left            =   705
         TabIndex        =   8
         Top             =   315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AlmoxarifadoFinal 
         Height          =   315
         Left            =   3225
         TabIndex        =   9
         Top             =   315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label labelAlmoxarifadoFinal 
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
         Left            =   2805
         TabIndex        =   14
         Top             =   360
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
         Left            =   330
         TabIndex        =   15
         Top             =   375
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDemApInv_LOcx.ctx":003E
      Left            =   945
      List            =   "RelOpDemApInv_LOcx.ctx":0040
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   270
      Width           =   2790
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
      Left            =   4080
      Picture         =   "RelOpDemApInv_LOcx.ctx":0042
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDemApInv_LOcx.ctx":0144
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDemApInv_LOcx.ctx":029E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDemApInv_LOcx.ctx":0428
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDemApInv_LOcx.ctx":095A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInv 
      Height          =   300
      Left            =   1560
      TabIndex        =   13
      Top             =   840
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
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
      Left            =   240
      TabIndex        =   16
      Top             =   315
      Width           =   615
   End
   Begin VB.Label Data 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   975
      TabIndex        =   17
      Top             =   900
      Width           =   480
   End
   Begin VB.Label LabelAlmoxarifado 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifados"
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
      Left            =   5640
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label9 
      Caption         =   "Ordena por:"
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
      Left            =   450
      TabIndex        =   19
      Top             =   1350
      Width           =   1080
   End
End
Attribute VB_Name = "RelOpDemApInv_LOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giAlmoxInicial As Integer

Private Sub DataInv_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInv)

End Sub

Private Sub DataInv_Validate(Cancel As Boolean)

Dim sDataInv As String
Dim lErro As Long

On Error GoTo Erro_DataInv_Validate

    If Len(DataInv.ClipText) > 0 Then

        sDataInv = DataInv.Text
        
        lErro = Data_Critica(sDataInv)
        If lErro <> SUCESSO Then Error 64817

    End If

    Exit Sub

Erro_DataInv_Validate:

    Cancel = True


    Select Case Err

        Case 64817

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168101)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInv, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 64818

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 64818
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168102)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInv, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 64819

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 64819
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168103)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoInicial_GotFocus()
'Mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoInicial_GotFocus

    giAlmoxInicial = 1

    lErro = Mostra_Lista_Almoxarifado
    If lErro <> SUCESSO Then Error 64820

    Exit Sub

Erro_AlmoxarifadoInicial_GotFocus:

    Select Case Err

        Case 64820

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168104)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoFinal_GotFocus()
'mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoFinal_GotFocus

    giAlmoxInicial = 0

    lErro = Mostra_Lista_Almoxarifado
    If lErro <> SUCESSO Then Error 64821

    Exit Sub

Erro_AlmoxarifadoFinal_GotFocus:

    Select Case Err

        Case 64821

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168105)

    End Select

    Exit Sub

End Sub

Private Function Mostra_Lista_Almoxarifado() As Long
'esconde a treeview de produto e mostra a lista de almoxarifados

Dim lErro As Long

On Error GoTo Erro_Mostra_Lista_Almoxarifado

    'mostra a ListBox de almoxarifados
    Almoxarifados.Visible = True
    LabelAlmoxarifado.Visible = True

    Mostra_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Mostra_Lista_Almoxarifado:

    Mostra_Lista_Almoxarifado = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168106)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
        
    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 64822
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 64822

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168107)

    End Select

    Exit Sub

End Sub

Private Sub Define_Padrao()
'Preenche a tela com as opções padrão

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giAlmoxInicial = 1
    
    ComboTotaliza.ListIndex = 0
        
    Call Mostra_Lista_Almoxarifado
    
    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168108)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iTotaliza As Integer

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 64823

    'Pega status e exibe
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then Error 64824

    'pega data e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINV", sParam)
    If lErro <> SUCESSO Then Error 64825

    Call DateParaMasked(DataInv, CDate(sParam))

   'pega parâmetro Almoxarifado Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXINIC", sParam)
    If lErro Then Error 64826
    
    AlmoxarifadoInicial.Text = sParam
    Call AlmoxarifadoInicial_Validate(bSGECancelDummy)
    
    'pega parâmetro Almoxarifado Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXFIM", sParam)
    If lErro Then Error 64827
    
    AlmoxarifadoFinal.Text = sParam
    Call AlmoxarifadoFinal_Validate(bSGECancelDummy)
       
    'pega parâmetro de totalização
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTOTALIZA", sParam)
    If lErro Then Error 64828

    'seleciona ítem no ComboTotaliza
    iTotaliza = CInt(sParam)
    ComboTotaliza.ListIndex = iTotaliza
              
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 64823 To 64828
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168109)

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

    If Not (gobjRelatorio Is Nothing) Then Error 64830
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 64829
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 64829
        
        Case 64830
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168110)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    AlmoxarifadoInicial.Text = ""
    AlmoxarifadoFinal.Text = ""
    
    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sAlmox_I, sAlmox_F) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
   'critica Almoxarifado Inicial e Final
    If AlmoxarifadoInicial.Text <> "" Then
        sAlmox_I = CStr(Codigo_Extrai(AlmoxarifadoInicial.Text))
        
    Else
        sAlmox_I = ""
        
    End If
        
    If AlmoxarifadoFinal.Text <> "" Then
        sAlmox_F = CStr(Codigo_Extrai(AlmoxarifadoFinal.Text))
    
    Else
        sAlmox_F = ""
        
    End If
       
    If sAlmox_I <> "" And sAlmox_F <> "" Then
          
        If sAlmox_I <> "" And sAlmox_F <> "" Then
        
            If CInt(sAlmox_I) > CInt(sAlmox_F) Then Error 64831
        
        End If
        
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
         
        Case 64831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INICIAL_MAIOR", Err)
            AlmoxarifadoInicial.SetFocus
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168111)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela
    Call Define_Padrao

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sAlmox_I As String
Dim sAlmox_F As String
Dim iIndice As Integer
Dim sTotaliza As String
Dim sStatus As String
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_PreencherRelOp
    
    If Len(DataInv.ClipText) = 0 Then Error 64832
        
    lErro = Formata_E_Critica_Parametros(sAlmox_I, sAlmox_F)
    If lErro <> SUCESSO Then Error 64833

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 64834

    sStatus = "1"

    lErro = objRelOpcoes.IncluirParametro("NSTATUS", sStatus)
    If lErro <> AD_BOOL_TRUE Then Error 64835
    
    If DataInv.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINV", DataInv.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINV", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 64836
                   
    lErro = objRelOpcoes.IncluirParametro("NALMOXINIC", sAlmox_I)
    If lErro <> AD_BOOL_TRUE Then Error 64837
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXINICIAL", AlmoxarifadoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 64838
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXFIM", sAlmox_F)
    If lErro <> AD_BOOL_TRUE Then Error 64839
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXFINAL", AlmoxarifadoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 64840
       
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
   
    'Ler o mês e o ano que está aberto
    lErro = CF("EstoqueMes_Le_Aberto",objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 40673 Then Error 64841

    If lErro = 40673 Then Error 64842
 
    lErro = objRelOpcoes.IncluirParametro("NANO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 64843
 
    lErro = objRelOpcoes.IncluirParametro("NMES", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 64844
    
    'le o ultimo ano/mes apurado
    lErro = CF("EstoqueMes_Le_Apurado",objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then Error 64845
    
    If lErro = 46225 Then
        objEstoqueMes.iAno = 0
        objEstoqueMes.iMes = 0
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NANOAPURADO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 64846
 
    lErro = objRelOpcoes.IncluirParametro("NMESAPURADO", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 64847
            
    sTotaliza = CStr(ComboTotaliza.ListIndex)
    
    lErro = objRelOpcoes.IncluirParametro("NTOTALIZA", sTotaliza)
    If lErro <> AD_BOOL_TRUE Then Error 64848

    If ComboTotaliza.ListIndex = 0 Then
'        If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipendl"
        If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daiprocl"
'        If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitodol"
    ElseIf ComboTotaliza.ListIndex = 1 Then
'        If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipentl"
        If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daiprotl"
'        If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitodtl"
    ElseIf ComboTotaliza.ListIndex = 2 Then
'       If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipeatl"
        If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daipratl"
'       If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitoatl"
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAlmox_I, sAlmox_F)
    If lErro <> SUCESSO Then Error 64849

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 64833 To 64849
                               
        Case 64832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168112)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 64850

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 64851

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 64850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 64851

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168113)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 64852

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 64852

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168114)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 64853

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 64854

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 64855

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 64853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 64854, 64855

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168115)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAlmox_I As String, sAlmox_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

     sExpressao = ""

    If sAlmox_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "Almoxarifado >= " & Forprint_ConvInt(CInt(sAlmox_I))

    End If

    If sAlmox_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado <= " & Forprint_ConvInt(CInt(sAlmox_F))

    End If
            
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168116)

    End Select

    Exit Function

End Function

Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Carrega_Lista_Almoxarifado
    
    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa",giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 64856

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifados.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = Err

    Select Case Err

        Case 64856

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168117)

    End Select

    Exit Function

End Function

Private Sub Almoxarifados_DblClick()
'Preenche Almoxarifado Final ou Inicial com o almoxarifado selecionado

Dim lErro As Long
Dim sListBoxItem As String
Dim objCodigoDescricao As New AdmCodigoNome
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objAlmoxSelecionado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifados_DblClick

    'Guarda a string selecionada na ListBox Almoxarifados
    sListBoxItem = Almoxarifados.List(Almoxarifados.ListIndex)
 
    If giAlmoxInicial = 1 Then
    
        AlmoxarifadoInicial.Text = sListBoxItem
        
    Else
        AlmoxarifadoFinal.Text = sListBoxItem

    End If

    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case Err

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168118)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoInicial_Validate

    If Len(Trim(AlmoxarifadoInicial.Text)) > 0 Then
   
        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoInicial, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 64857

    End If
    
    Exit Sub

Erro_AlmoxarifadoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 64857

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168119)

    End Select

End Sub

Private Sub AlmoxarifadoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoFinal_Validate

    If Len(Trim(AlmoxarifadoFinal.Text)) > 0 Then

        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoFinal, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 64858

    End If
 
    Exit Sub

Erro_AlmoxarifadoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 64858

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168120)

    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DEM_APURACAO_INVENTARIO_L
    Set Form_Load_Ocx = Me
    Caption = "Demonstrativo de Apuração de Inventário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDemApInv_L"
    
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


Private Sub labelAlmoxarifadoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelAlmoxarifadoFinal, Source, X, Y)
End Sub

Private Sub labelAlmoxarifadoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelAlmoxarifadoFinal, Button, Shift, X, Y)
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

Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

