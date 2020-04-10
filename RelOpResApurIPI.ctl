VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpResApurIPIOcx 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ScaleHeight     =   2610
   ScaleWidth      =   7095
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
      Left            =   255
      TabIndex        =   20
      Top             =   2025
      Width           =   1320
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   720
      Left            =   120
      TabIndex        =   15
      Top             =   900
      Width           =   4950
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1860
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   840
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
         Height          =   300
         Left            =   4350
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3330
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
         Left            =   2925
         TabIndex        =   19
         Top             =   315
         Width           =   360
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
         Left            =   450
         TabIndex        =   18
         Top             =   285
         Width           =   345
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
      Height          =   480
      Left            =   5190
      Picture         =   "RelOpResApurIPI.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   930
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressão"
      Height          =   645
      Left            =   1800
      TabIndex        =   14
      Top             =   1770
      Width           =   3285
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
         TabIndex        =   3
         Top             =   300
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
         Left            =   1770
         TabIndex        =   4
         Top             =   270
         Width           =   1215
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
      Left            =   5190
      TabIndex        =   7
      Top             =   2070
      Width           =   1755
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
      Left            =   5190
      TabIndex        =   6
      Top             =   1500
      Width           =   1755
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpResApurIPI.ctx":0102
      Left            =   930
      List            =   "RelOpResApurIPI.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2916
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4845
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpResApurIPI.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpResApurIPI.ctx":0260
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpResApurIPI.ctx":03EA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpResApurIPI.ctx":091C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
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
      Left            =   225
      TabIndex        =   13
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpResApurIPIOcx"
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
    If lErro <> SUCESSO Then gError 75384
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
                    
        Case 75384
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172986)

    End Select

    Exit Sub

End Sub

Function Traz_LivroFiscal_Tela() As Long

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial
Dim objLivroFechado As New ClassLivrosFechados

On Error GoTo Erro_Traz_LivroFiscal_Tela

    'Verifica se o Livro de Registro de Entrada está aberto
    objLivrosFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    lErro = CF("LivrosFilial_Le", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67992 Then gError 70845
               
    'Se não encontrou o Livro de Registro de Entrada Aberto
    If lErro = 67992 Then
        
        'Lê o último livro de Registro de Entrada Fechado
        objLivroFechado.iCodLivro = LIVRO_APURACAO_IPI_CODIGO
        objLivroFechado.iFilialEmpresa = giFilialEmpresa
        lErro = CF("LivrosFechados_Le_UltimaData", objLivroFechado)
        If lErro <> SUCESSO And lErro <> 70231 Then gError 70846
                                
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
        
        Case 70845, 70846
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172987)
    
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
    If lErro Then gError 70847
    
    'pega Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro Then gError 70848
    
    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega Data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 70849

    Call DateParaMasked(DataFinal, CDate(sParam))
        
    'Definitiva
    lErro = objRelOpcoes.ObterParametro("NTESTEDEF", sParam)
    If lErro Then gError 70850
    
    OptionTipo(CInt(sParam)).Value = True
      
    'Verifica se as datas Inicial e Final estão dentro do Intervalo de um Livro Aberto ou fechado
    lErro = LivroFiscal_Data_Critica()
    If lErro <> SUCESSO Then gError 70851
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 70847 To 70851
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172988)

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

    If Not (gobjRelatorio Is Nothing) Then gError 70853
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 70852

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 70852
        
        Case 70853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172989)

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

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'Se a data Inicial não está preenchida, erro
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 70854
    
    'Se a data Final não está preenchida, erro
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 70855
    
    'Data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 70857
    
        'Se foi selecionada a Impressão definitiva
        If OptionTipo(0).Value = True Then
            
            'Verifica se as datas Fazem parte de um Livro Fiscal de Registro de inventário aberto ou Fechado
            lErro = LivroFiscal_Data_Critica(sFolha)
            If lErro <> SUCESSO Then gError 70856
        
        End If
            
    End If
               
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                                       
        Case 70854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
        
        Case 70855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
        
        Case 70856
        
        Case 70857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172990)

    End Select

    Exit Function

End Function

Function LivroFiscal_Data_Critica(Optional sFolha As String) As Long

Dim lErro As Long
Dim objLivroFechado As New ClassLivrosFechados
Dim objLivrosFilial As New ClassLivrosFilial
Dim objApuracao As New ClassRegApuracao

On Error GoTo Erro_LivroFiscal_Data_Critica
    
    'Verifica se o Livro de Registro de Entrada está aberto
    objLivrosFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    objLivrosFilial.dtDataInicial = StrParaDate(DataInicial.Text)
    objLivrosFilial.dtDataFinal = StrParaDate(DataFinal.Text)

    'Lê o Livro Fiscal Aberto que possui a data inicial e final dentro do intervalo passado
    lErro = CF("LivrosFilial_Le_IntervaloData", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 70599 Then gError 70858
               
    'Se não encontrou o Livro de Registro de Entrada Aberto com as datas no intervalo passado
    If lErro = 70599 Then
        
        'Lê o Livro Fiscal Fechado que possui a data inicial e final dentro do intervalo passado
        objLivroFechado.iCodLivro = LIVRO_APURACAO_IPI_CODIGO
        objLivroFechado.iFilialEmpresa = giFilialEmpresa
        objLivroFechado.dtDataInicial = StrParaDate(DataInicial.Text)
        objLivroFechado.dtDataFinal = StrParaDate(DataFinal.Text)
        lErro = CF("LivrosFechados_Le_IntervaloData", objLivroFechado)
        If lErro <> SUCESSO And lErro <> 70595 Then gError 70859
        
        'Se não encontrou o Livro de Registro de Entrada Fechado, erro
        If lErro = 70595 Then gError 70861
                
        CheckFechado.Value = vbChecked
            
        sFolha = CStr(objLivroFechado.iFolhaInicial)
        
    'Se encontrou o Livro Fiscal passado
    Else
        CheckFechado.Value = vbUnchecked
        sFolha = CStr(objLivrosFilial.iNumeroProxFolha)
    End If
        
'    'Verifica se existem ApuraçõesIPI com o intervalo de data passados
'    objApuracao.iFilialEmpresa = giFilialEmpresa
'    objApuracao.dtDataInicial = StrParaDate(DataInicial.Text)
'    objApuracao.dtDataFinal = StrParaDate(DataFinal.Text)
'    lErro = CF("ApuracaoIPI_Le_IntervaloData",objApuracao)
'    If lErro <> SUCESSO And lErro <> 70717 Then gError 70860
'
'    'Se não encontrou Apuração IPI, erro
'    If lErro = 70717 Then gError 70862
    
    LivroFiscal_Data_Critica = SUCESSO
    
    Exit Function

Erro_LivroFiscal_Data_Critica:
    
    LivroFiscal_Data_Critica = gErr
    
    Select Case gErr
    
        Case 70858, 70859, 70860
                    
        Case 70861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTERVALO_DATA_DIFERENTE_LIVROFISCAL", gErr, DataInicial.Text, DataFinal.Text)
        
        Case 70862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPI_NAO_CADASTRADA", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172991)
        
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
        
    colSelecao.Add LIVRO_APURACAO_IPI_CODIGO
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172992)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sTipo As String
Dim sFolha As String
Dim lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp

    'Critica as datas
    lErro = Formata_E_Critica_Parametros(sFolha)
    If lErro <> SUCESSO Then gError 70863
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 70864
      
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 70865

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 70866
        
    For iIndice = 0 To 1
        If OptionTipo(iIndice).Value = True Then sTipo = CStr(iIndice)
    Next
    
    lErro = objRelOpcoes.IncluirParametro("NTESTEDEF", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError 70867
        
    lErro = objRelOpcoes.IncluirParametro("NFOLHA", sFolha)
    If lErro <> AD_BOOL_TRUE Then gError 70868
    
    If bExecutando Then
    
        lErro = CF("RelRegApurIPI_Prepara", giFilialEmpresa, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), StrParaInt(sFolha), lNumIntRel)
        If lErro <> SUCESSO Then gError 70868
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 70868
    
    End If
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 70863 To 70868

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172993)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 70869

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 70870

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call Limpar_Tela
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 70869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 70870

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172994)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 70871

    lErro = gobjRelatorio.Executar_Prossegue2(Me)
    If lErro <> SUCESSO And lErro <> 7072 Then gError 70896
    
    'Se cancelou o Relatório
    If lErro = 7072 Then gError 70897
    
    'Se foi selecionada a Impressão definitiva
    If OptionTipo(0).Value = True Then
        
        objLivrosFilial.iFilialEmpresa = giFilialEmpresa
        objLivrosFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO
        objLivrosFilial.dtDataInicial = StrParaDate(DataInicial.Text)
        objLivrosFilial.dtDataFinal = StrParaDate(DataFinal.Text)
        
        'Atualiza data de Impressão
        lErro = CF("LivrosFilial_Atualiza_DataImpresao", objLivrosFilial)
        If lErro <> SUCESSO Then gError 70872
    
    End If
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 70871, 70872, 70896
        
        Case 70897
            Unload Me

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172995)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 70873

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 70874

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 70875

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 70873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 70874, 70875

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172996)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLivroAberto_Click()

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial

On Error GoTo Erro_BotaoLivroAberto_Click
    
    'Verifica se o Livro de Registro de Entrada está aberto
    objLivrosFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa
    lErro = CF("LivrosFilial_Le", objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67992 Then gError 70876
               
    'Se não encontrou o Livro de Registro de Entrada Aberto
    If lErro = 67992 Then gError 70877
    
    Call DateParaMasked(DataInicial, objLivrosFilial.dtDataInicial)
    Call DateParaMasked(DataFinal, objLivrosFilial.dtDataFinal)
    
    CheckFechado.Value = vbUnchecked
    
    Exit Sub
    
Erro_BotaoLivroAberto_Click:
    
    Select Case gErr
    
        Case 70876
        
        Case 70877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVRO_FISCAL_ABERTO_INEXISTENTE", gErr, LIVRO_APURACAO_IPI_CODIGO)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172997)
    
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
        If lErro <> SUCESSO Then gError 70878

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 70878

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172998)

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
        If lErro <> SUCESSO Then gError 70879

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 70879

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172999)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 70880

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 70880
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173000)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 70881

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 70881
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173001)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 70882

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 70882
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173002)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 70883

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 70883
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173003)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Resumo da Apuração IPI"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelResOpApurIPI"
    
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

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

