VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRegInventarioOcx 
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ScaleHeight     =   3975
   ScaleWidth      =   6435
   Begin VB.CheckBox FiltroNatureza 
      Caption         =   $"RelOpRegInventario.ctx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   255
      TabIndex        =   23
      ToolTipText     =   "O inventário deverá ser apresentado no arquivo da EFD-ICMS/IPI, no segundo mês subsequente ao evento."
      Top             =   3225
      Value           =   1  'Checked
      Width           =   5460
   End
   Begin VB.Frame Frame2 
      Caption         =   "Quebra por"
      Height          =   1365
      Left            =   2130
      TabIndex        =   16
      Top             =   1350
      Width           =   2205
      Begin VB.OptionButton Quebra 
         Caption         =   "Natureza"
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
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1050
         Width           =   1725
      End
      Begin VB.OptionButton Quebra 
         Caption         =   "Class. Fiscal"
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
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton Quebra 
         Caption         =   "Almoxarifado"
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
         Left            =   240
         TabIndex        =   18
         Top             =   540
         Width           =   1425
      End
      Begin VB.OptionButton Quebra 
         Caption         =   "Conta Contábil"
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
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   790
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressão"
      Height          =   1365
      Left            =   180
      TabIndex        =   13
      Top             =   1350
      Width           =   1755
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
         Left            =   240
         TabIndex        =   15
         Top             =   450
         Value           =   -1  'True
         Width           =   1155
      End
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
         Left            =   240
         TabIndex        =   14
         Top             =   870
         Width           =   1215
      End
   End
   Begin VB.CommandButton BotaoRegCadastrado 
      Caption         =   "Registro de Inventário Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4560
      TabIndex        =   3
      Top             =   1740
      Width           =   1755
   End
   Begin VB.CheckBox CheckFechado 
      Caption         =   "Já fechado"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   900
      Width           =   1320
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
      Height          =   675
      Left            =   4575
      Picture         =   "RelOpRegInventario.ctx":00A6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   870
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4185
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "RelOpRegInventario.ctx":01A8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRegInventario.ctx":0302
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRegInventario.ctx":048C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRegInventario.ctx":09BE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRegInventario.ctx":0B3C
      Left            =   1080
      List            =   "RelOpRegInventario.ctx":0B3E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2916
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2070
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   825
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   300
      Left            =   1050
      TabIndex        =   1
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
   Begin MSMask.MaskEdBox Folha 
      Height          =   300
      Left            =   1815
      TabIndex        =   21
      Top             =   2880
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "A partir da Folha:"
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
      Left            =   240
      TabIndex        =   22
      Top             =   2925
      Width           =   1485
   End
   Begin VB.Label dIni 
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
      Left            =   510
      TabIndex        =   12
      Top             =   870
      Width           =   480
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
      Left            =   435
      TabIndex        =   9
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRegInventarioOcx"
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
Private WithEvents objEventoBotaoInv As AdmEvento
Attribute objEventoBotaoInv.VB_VarHelpID = -1

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
            
    Set objEventoBotaoInv = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172243)

    End Select

    Exit Sub

End Sub

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
    If lErro Then gError 70961
    
    'pega Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro Then gError 70962
    
    Call DateParaMasked(Data, CDate(sParam))
            
    'Definitiva
    lErro = objRelOpcoes.ObterParametro("NTESTEDEF", sParam)
    If lErro Then gError 70963
    
    OptionTipo(CInt(sParam)).Value = True
          
    'pega a folha e exibe
    lErro = objRelOpcoes.ObterParametro("NFOLHA", sParam)
    If lErro <> SUCESSO Then gError 70963

    If Len(Trim(sParam)) > 0 Then Folha.Text = CInt(sParam)

    lErro = objRelOpcoes.ObterParametro("NFILTRONATUREZA", sParam)
    If lErro <> SUCESSO Then gError 70963
    
    If StrParaInt(sParam) = 2 Then
        FiltroNatureza.Value = vbUnchecked
    Else
        FiltroNatureza.Value = vbChecked
    End If

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 70961 To 70963
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172244)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoBotaoInv = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 70965
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 70964

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 70964
        
        Case 70965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172245)

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
    
End Sub

Private Function Formata_E_Critica_Parametros(sFolha As String) As Long

Dim lErro As Long
Dim dtData As Date
Dim objRegInventario As New ClassRegInventario
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'Se a data não está preenchida, erro
    If Len(Trim(Data.ClipText)) = 0 Then gError 70966
                           
    'Se a folha não foi preenchida ---> Erro
    If Len(Trim(Folha.ClipText)) = 0 Then gError 84702 '78089

    dtData = StrParaDate(Data.Text)
    
    'Verifica se a Data está entre a DataInicial e Final de algum LivroAberto ou LivroFechado
    lErro = CF("LivrosFiscais_Valida_Data", dtData)
    If lErro <> SUCESSO And lErro <> 76323 Then gError 76324
    
    If lErro = 76323 Then gError 76325
    
    objRegInventario.dtData = StrParaDate(Data.Text)
    objRegInventario.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se existem Registro de Inventários gerados para a data
    lErro = CF("RegInventario_Le_Data", objRegInventario)
    If lErro <> SUCESSO And lErro <> 70237 Then gError 70968
    
    'Se não encontrou o Registro de Inventário, erro
    If lErro = 70237 Then gError 70969
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 84702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FOLHA_NAO_PREENCHIDA", gErr)
            Folha.SetFocus

        Case 70562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            Data.SetFocus
        
        Case 70966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr, dtData)
        
        Case 70968
        
        Case 70969
            'Pergunta se deseja criar Registro de inventário
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGINVENTARIO", Data.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela EdicaoRegInventario
                Call Chama_Tela("EdicaoRegInventario", objRegInventario)
            End If
                        
        Case 76324
        
        Case 76325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINVENTARIO_FORA_PERIODO", gErr, dtData)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172246)

    End Select

    Exit Function

End Function

Function LivroFiscal_Data_Critica(Optional sFolha As String) As Long

Dim lErro As Long
Dim objLivroFechado As New ClassLivrosFechados
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_LivroFiscal_Data_Critica
    
    'Verifica se o Livro de Registro de Entrada está aberto
    objLivrosFilial.iCodLivro = LIVRO_REG_INVENTARIO_CODIGO
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    objLivrosFilial.dtDataInicial = StrParaDate(Data.Text)
    objLivrosFilial.dtDataFinal = StrParaDate(Data.Text)

    'Lê o Livro Fiscal Aberto que possui a data inicial e final dentro do intervalo passado
    lErro = CF("LivrosFilial_Le_IntervaloData", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 70599 Then gError 70970
               
    'Se não encontrou o Livro de Registro de Entrada Aberto com as datas no intervalo passado
    If lErro = 70599 Then
        
        'Lê o Livro Fiscal Fechado que possui a data inicial e final dentro do intervalo passado
        objLivroFechado.iCodLivro = LIVRO_REG_INVENTARIO_CODIGO
        objLivroFechado.iFilialEmpresa = giFilialEmpresa
        objLivroFechado.dtDataInicial = StrParaDate(Data.Text)
        objLivroFechado.dtDataFinal = StrParaDate(Data.Text)
        lErro = CF("LivrosFechados_Le_IntervaloData", objLivroFechado)
        If lErro <> SUCESSO And lErro <> 70595 Then gError 70971
        
        'Se não encontrou o Livro de Registro de Entrada Fechado, erro
        If lErro = 70595 Then gError 70972
                
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
    
        Case 70970, 70971
                    
        Case 70972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTERVALO_DATA_DIFERENTE_LIVROFISCAL", gErr, Data.Text, Data.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172247)
        
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub BotaoRegCadastrado_Click()

Dim colSelecao As New Collection
Dim objRegInventario As New ClassRegInventario
    
    Call Chama_Tela("RegInventarioLista", colSelecao, objRegInventario, objEventoBotaoInv)

End Sub

Private Sub objEventoBotaoInv_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRegInventario As ClassRegInventario

On Error GoTo Erro_objEventoBotaoInv_evSelecao

    Set objRegInventario = obj1
    
    Call DateParaMasked(Data, objRegInventario.dtData)
        
    'Critica data do Livro Fiscal
    lErro = CF("LivrosFiscais_Valida_Data", objRegInventario.dtData)
    If lErro <> SUCESSO And lErro <> 76323 Then gError 70973
    If lErro = 76323 Then gError 76326
    
    Me.Show

    Exit Sub

Erro_objEventoBotaoInv_evSelecao:

    Select Case gErr
        
        Case 70973
        
        Case 76326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINVENTARIO_FORA_PERIODO", gErr, objRegInventario.dtData)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172248)

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
'preenche o arquivo com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sTipo As String
Dim sFolha As String
Dim iFiltroNatureza As Integer

On Error GoTo Erro_PreencherRelOp

    'Critica as datas
    lErro = Formata_E_Critica_Parametros(sFolha)
    If lErro <> SUCESSO Then gError 70974
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 70975
      
    If Data.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATA", Data.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATA", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 70976
        
    For iIndice = 0 To 1
        If OptionTipo(iIndice).Value = True Then sTipo = CStr(iIndice)
    Next
    
    lErro = objRelOpcoes.IncluirParametro("NTESTEDEF", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError 70977
        
    lErro = objRelOpcoes.IncluirParametro("NFOLHA", CInt(Folha.Text))
    If lErro <> AD_BOOL_TRUE Then gError 70978
    
    If FiltroNatureza.Value = vbChecked Then
        iFiltroNatureza = MARCADO
    Else
        iFiltroNatureza = 2 '0(zero) vai ser considerado marcado para trazer gravações antigas para tela com a marcação
    End If
                        
    lErro = objRelOpcoes.IncluirParametro("NFILTRONATUREZA", CInt(iFiltroNatureza))
    If lErro <> AD_BOOL_TRUE Then gError 70978
    
    If Quebra(0).Value = True Then
        gobjRelatorio.sNomeTsk = "RegInv"
    ElseIf Quebra(1).Value = True Then
        gobjRelatorio.sNomeTsk = "RegInvAl"
    ElseIf Quebra(2).Value = True Then
        gobjRelatorio.sNomeTsk = "RegInvCC"
    ElseIf Quebra(3).Value = True Then
        gobjRelatorio.sNomeTsk = "RegInvNa"
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 70974 To 70978

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172249)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 70979

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 70980

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 70979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 70980

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172250)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 70981
        
    lErro = gobjRelatorio.Executar_Prossegue2(Me)
    If lErro <> SUCESSO And lErro <> 7072 Then gError 70982
    
    'Se cancelou o Relatório
    If lErro = 7072 Then gError 70983
    
    'Se foi selecionada a Impressão definitiva
    If OptionTipo(0).Value = True Then
        
        objLivrosFilial.iFilialEmpresa = giFilialEmpresa
        objLivrosFilial.iCodLivro = LIVRO_REG_INVENTARIO_CODIGO
        
        'Atualiza data de Impressão
        lErro = CF("LivrosFilial_Atualiza_DataImpresao", objLivrosFilial)
        If lErro <> SUCESSO Then gError 70984
    
    End If
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 70981, 70984, 70982
        
        Case 70983
            Unload Me
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172251)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 70985

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 70986

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 70987

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 70985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 70986, 70987

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172252)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Data.ClipText) > 0 Then

        sDataInic = Data.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 70988

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 70988

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172253)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 70989

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 70989
            Data.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172254)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 70990

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 70990
            Data.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172255)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Data Then
            Call BotaoRegCadastrado_Click
        End If
    End If

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Livro de Reg. de Inventário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRegInventario"
    
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

