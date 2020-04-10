VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRegEntradaOcx 
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   ScaleHeight     =   3015
   ScaleWidth      =   7050
   Begin VB.Frame Frame3 
      Caption         =   "Emitente"
      Height          =   1005
      Left            =   2100
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   1365
      Begin VB.OptionButton Emitente 
         Caption         =   "Código"
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
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton Emitente 
         Caption         =   "Nome"
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
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modelo"
      Height          =   1005
      Left            =   240
      TabIndex        =   25
      Top             =   1530
      Width           =   2205
      Begin VB.OptionButton Modelo1 
         Caption         =   "Modelo 1"
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
         Index           =   0
         Left            =   450
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton Modelo1 
         Caption         =   "Modelo 1A"
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
         Index           =   1
         Left            =   450
         TabIndex        =   6
         Top             =   660
         Width           =   1245
      End
   End
   Begin VB.CommandButton BotaoLivroAberto 
      Caption         =   "Traz Livro Aberto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4950
      TabIndex        =   11
      Top             =   2100
      Width           =   1755
   End
   Begin VB.CheckBox CheckFechado 
      Caption         =   "Livro já fechado"
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
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   2670
      Width           =   1830
   End
   Begin VB.CommandButton BotaoLivroCadastrados 
      Caption         =   "Livros Fechados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4950
      TabIndex        =   10
      Top             =   1470
      Width           =   1755
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
      Height          =   480
      Left            =   4950
      Picture         =   "RelOpRegEntrada.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   1755
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   720
      Left            =   255
      TabIndex        =   19
      Top             =   720
      Width           =   4410
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1665
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   645
         TabIndex        =   1
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   3750
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   2730
         TabIndex        =   2
         Top             =   255
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIni 
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   285
         TabIndex        =   23
         Top             =   315
         Width           =   345
      End
      Begin VB.Label dFim 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2325
         TabIndex        =   22
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressão"
      Height          =   1005
      Left            =   2640
      TabIndex        =   18
      Top             =   1530
      Width           =   2055
      Begin VB.OptionButton OptionTipo 
         Caption         =   "Teste"
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
         Index           =   1
         Left            =   420
         TabIndex        =   4
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton OptionTipo 
         Caption         =   "Definitiva"
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
         Index           =   0
         Left            =   420
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRegEntrada.ctx":0102
      Left            =   1590
      List            =   "RelOpRegEntrada.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2916
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4755
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRegEntrada.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRegEntrada.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRegEntrada.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRegEntrada.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
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
      Left            =   915
      TabIndex        =   17
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRegEntradaOcx"
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

'Eventos dos Browses
Private WithEvents objEventoLivrosFechados As AdmEvento
Attribute objEventoLivrosFechados.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
            
    Set objEventoLivrosFechados = New AdmEvento
    
    'Traz dados do último Livro Fiscal para a tela
    lErro = Traz_LivroFiscal_Tela()
    If lErro <> SUCESSO Then gError 75381
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 75381
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172210)

    End Select

    Exit Sub

End Sub

Function Traz_LivroFiscal_Tela() As Long

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial
Dim objLivroFechado As New ClassLivrosFechados

On Error GoTo Erro_Traz_LivroFiscal_Tela
    
    'Verifica se o Livro de Registro de Entrada está aberto
    objLivrosFilial.iCodLivro = LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    lErro = CF("LivrosFilial_Le", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67992 Then gError 70554
               
    'Se não encontrou o Livro de Registro de Entrada Aberto
    If lErro = 67992 Then
        
        'Lê o último livro de Registro de Entrada Fechado
        objLivroFechado.iCodLivro = LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
        objLivroFechado.iFilialEmpresa = giFilialEmpresa
        lErro = CF("LivrosFechados_Le_UltimaData", objLivroFechado)
        If lErro <> SUCESSO And lErro <> 70231 Then gError 70555
                                
        If lErro = SUCESSO Then
        
            'Coloca as datas do último Livro de Registro de Entrada Fechado na tela
            Call DateParaMasked(DataInicial, objLivroFechado.dtDataInicial)
            Call DateParaMasked(DataFinal, objLivroFechado.dtDataFinal)
            
            CheckFechado.Value = vbChecked
                
        End If
        
    'Se encontro o Livro de Registro de Entrada aberto
    Else
    
        'Coloca as datas do Livro de Registro de Entrada Aberto na tela
        Call DateParaMasked(DataInicial, objLivrosFilial.dtDataInicial)
        Call DateParaMasked(DataFinal, objLivrosFilial.dtDataFinal)
        
        CheckFechado.Value = vbUnchecked
    
    End If


    Traz_LivroFiscal_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_LivroFiscal_Tela:

    Traz_LivroFiscal_Tela = gErr
        
    Select Case gErr
        
        Case 70554, 70555
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172211)
    
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    'Limpa a tela
    Call Limpar_Tela

    'Carrega Opções de Relatório
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 70557
    
    'pega Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro Then gError 70558
    
    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 70559

    Call DateParaMasked(DataFinal, CDate(sParam))
        
    'Definitiva
    lErro = objRelOpcoes.ObterParametro("NTESTEDEF", sParam)
    If lErro Then gError 70560
    
    OptionTipo(CInt(sParam)).Value = True
          
    'Emitente
    lErro = objRelOpcoes.ObterParametro("NEMITENTE", sParam)
    If lErro Then gError 75370
    
    Emitente(CInt(sParam)).Value = True
    
    'Modelo
    lErro = objRelOpcoes.ObterParametro("NMODELO", sParam)
    If lErro Then gError 75372
    
    Modelo1(CInt(sParam)).Value = True
    
    'Verifica se as datas Inicial e Final estão dentro do Intervalo de um Livro Aberto ou fechado
    lErro = LivroFiscal_Data_Critica()
    If lErro <> SUCESSO Then gError 70568
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 70557 To 70560, 70568, 75372
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172212)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoLivrosFechados = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 70581
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    Caption = gobjRelatorio.sCodRel

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 70561

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 70561
        
        Case 70581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172213)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    CheckFechado.Value = vbUnchecked
    
    ComboOpcoes.SetFocus
    OptionTipo(0).Value = True
    Emitente(0).Value = True
    Modelo1(0).Value = True
    
End Sub

Private Function Formata_E_Critica_Parametros(sFolha As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'Se a data Inicial não está preenchida, erro
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 70587
    
    'Se a data Final não está preenchida, erro
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 70588
    
    'Data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 70562
    
        'Se foi selecionada a Impressão definitiva
        If OptionTipo(0).Value = True Then
            
            'Verifica se as datas Fazem parte de um Livro Fiscal de Registro de inventário aberto ou Fechado
            lErro = LivroFiscal_Data_Critica(sFolha)
            If lErro <> SUCESSO Then gError 70582
    
        End If
        
    End If
               
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 70562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
                       
        Case 70582
        
        Case 70587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
        
        Case 70588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172214)

    End Select

    Exit Function

End Function

Function LivroFiscal_Data_Critica(Optional sFolha As String) As Long

Dim lErro As Long
Dim objLivroFechado As New ClassLivrosFechados
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_LivroFiscal_Data_Critica
    
    'Verifica se o Livro de Registro de Entrada está aberto
    objLivrosFilial.iCodLivro = LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    objLivrosFilial.dtDataInicial = StrParaDate(DataInicial.Text)
    objLivrosFilial.dtDataFinal = StrParaDate(DataFinal.Text)

    'Lê o Livro Fiscal Aberto que possui a data inicial e final dentro do intervalo passado
    lErro = CF("LivrosFilial_Le_IntervaloData", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 70599 Then gError 70583
               
    'Se não encontrou o Livro de Registro de Entrada Aberto com as datas no intervalo passado
    If lErro = 70599 Then
        
        'Lê o Livro Fiscal Fechado que possui a data inicial e final dentro do intervalo passado
        objLivroFechado.iCodLivro = LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
        objLivroFechado.iFilialEmpresa = giFilialEmpresa
        objLivroFechado.dtDataInicial = StrParaDate(DataInicial.Text)
        objLivroFechado.dtDataFinal = StrParaDate(DataFinal.Text)
        lErro = CF("LivrosFechados_Le_IntervaloData", objLivroFechado)
        If lErro <> SUCESSO And lErro <> 70595 Then gError 70584
        
        'Se não encontrou o Livro de Registro de Entrada Fechado, erro
        If lErro = 70595 Then gError 70585
                
        CheckFechado.Value = vbChecked
            
        sFolha = CStr(objLivroFechado.iFolhaInicial)
        
    'Se encontrou o Livro Fiscal passado
    Else
        CheckFechado.Value = vbUnchecked
        sFolha = CStr(objLivrosFilial.iNumeroProxFolha)
    End If
    
    LivroFiscal_Data_Critica = SUCESSO
    
    Exit Function

Erro_LivroFiscal_Data_Critica:
    
    LivroFiscal_Data_Critica = gErr
    
    Select Case gErr
    
        Case 70583, 70584
                    
        Case 70585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTERVALO_DATA_DIFERENTE_LIVROFISCAL", gErr, DataInicial.Text, DataFinal.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172215)
        
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub BotaoLivroCadastrados_Click()

Dim colSelecao As New Collection
Dim objLivrosFechados As ClassLivrosFechados
        
    colSelecao.Add LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
    
    Call Chama_Tela("LivrosFechadosLista", colSelecao, objLivrosFechados, objEventoLivrosFechados)

End Sub

Private Sub objEventoLivrosFechados_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objLivrosFechados As ClassLivrosFechados

On Error GoTo Erro_objEventoLivrosFechados_evSelecao

    Set objLivrosFechados = obj1
    
    Call DateParaMasked(DataInicial, objLivrosFechados.dtDataInicial)
    Call DateParaMasked(DataFinal, objLivrosFechados.dtDataFinal)
    
    CheckFechado.Value = vbChecked
    
    Me.Show

    Exit Sub

Erro_objEventoLivrosFechados_evSelecao:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172216)

    End Select

    Exit Sub

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
Dim iIndice As Integer
Dim sTipo As String
Dim sFolha As String
Dim sEmitente As String
Dim sModelo As String

On Error GoTo Erro_PreencherRelOp

    'Critica as datas
    lErro = Formata_E_Critica_Parametros(sFolha)
    If lErro <> SUCESSO Then gError 70563
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 70564
      
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 70565

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 70566
        
    For iIndice = 0 To 1
        If OptionTipo(iIndice).Value = True Then sTipo = CStr(iIndice)
    Next
    
    lErro = objRelOpcoes.IncluirParametro("NTESTEDEF", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError 70567
        
    For iIndice = 0 To 1
        If Emitente(iIndice).Value = True Then sEmitente = CStr(iIndice)
    Next
    
    lErro = objRelOpcoes.IncluirParametro("NEMITENTE", sEmitente)
    If lErro <> AD_BOOL_TRUE Then gError 75369
    
    For iIndice = 0 To 1
        If Modelo1(iIndice).Value = True Then sModelo = CStr(iIndice)
    Next
    
    lErro = objRelOpcoes.IncluirParametro("NMODELO", sModelo)
    If lErro <> AD_BOOL_TRUE Then gError 75371
    
    lErro = objRelOpcoes.IncluirParametro("NFOLHA", sFolha)
    If lErro <> AD_BOOL_TRUE Then gError 70556
                
    If Modelo1(0).Value = True Then
        gobjRelatorio.sNomeTsk = "LREMod1"
    Else
        gobjRelatorio.sNomeTsk = "LREMod1a"
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 70556, 70563 To 70567, 75369, 75371

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172217)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 70569

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 70570

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 70569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 70570

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172218)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 70571
        
    lErro = gobjRelatorio.Executar_Prossegue2(Me)
    If lErro <> SUCESSO And lErro <> 7072 Then gError 70890
    
    'Se cancelou o Relatório
    If lErro = 7072 Then gError 70891
    
    'Se foi selecionada a Impressão definitiva
    If OptionTipo(0).Value = True Then
        
        objLivrosFilial.iFilialEmpresa = giFilialEmpresa
        objLivrosFilial.iCodLivro = LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
        objLivrosFilial.dtDataInicial = StrParaDate(DataInicial.Text)
        objLivrosFilial.dtDataFinal = StrParaDate(DataFinal.Text)
        
        'Atualiza data de Impressão
        lErro = CF("LivrosFilial_Atualiza_DataImpresao", objLivrosFilial)
        If lErro <> SUCESSO Then gError 70792
    
    End If
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 70571, 70792, 70890
        
        Case 70891
            Unload Me
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172219)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 70572

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 70573

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 70574

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 70572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 70573, 70574

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172220)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLivroAberto_Click()

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_BotaoLivroAberto_Click
    
    'Verifica se o Livro de Registro de Entrada está aberto
    objLivrosFilial.iCodLivro = LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    lErro = CF("LivrosFilial_Le", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67992 Then gError 70589
               
    'Se não encontrou o Livro de Registro de Entrada Aberto
    If lErro = 67992 Then gError 70590
    
    Call DateParaMasked(DataInicial, objLivrosFilial.dtDataInicial)
    Call DateParaMasked(DataFinal, objLivrosFilial.dtDataFinal)
    
    CheckFechado.Value = vbUnchecked
    
    Exit Sub
    
Erro_BotaoLivroAberto_Click:
    
    Select Case gErr
    
        Case 70589
        
        Case 70590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVRO_FISCAL_ABERTO_INEXISTENTE", gErr, LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172221)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 70575

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 70575

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172222)

    End Select

    Exit Sub

End Sub


Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 70576

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 70576

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172223)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 70577

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 70577
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172224)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 70578

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 70578
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172225)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 70579

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 70579
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172226)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 70580

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 70580
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172227)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Livro de Reg. de Entradas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRegEntrada"
    
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


Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'se impressao definitiva
    'se confirmar execucao atualizar LivroFilial (data de impressao)

'Como default colocar a Data de e data Ate com os dados de Livro Filial
'Se não existir colocar o ultimo fechado

'Sabe-se qual é o livro Através da Data de e Data Até e Código do livro que é 1
'Para imprimir um relátorio as datas de Até tem que está dentro de um
'livro Configurado em livros filial ou ser um livro fechado em LivrosFechados
'dar erro se nao estiver

'Parametros que o Relatório espera

'Parametros que serão lidos da Tabela de LivrosFilial se for um livro
'aberto e configurado ou da Tabela de Livros Fechados se for um livro
'já fechado
'@NFolha

'Dados vindos da Tela
'@DDataDe
'@DDataAte
'@NTesteDef -> 0 Teste; 1 -> Definitiva

'O Botão Livros Cadastrados --> Lista os Livros já Fechados e o Livro Filial Filial aberto

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

