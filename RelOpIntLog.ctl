VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpIntLogOcx 
   Appearance      =   0  'Flat
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   6615
   Begin VB.Frame Frame4 
      Caption         =   "Integração"
      Height          =   2070
      Left            =   4500
      TabIndex        =   22
      Top             =   1770
      Width           =   1980
      Begin VB.OptionButton Nomal 
         Caption         =   "Só executar o relatório"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   75
         TabIndex        =   25
         Top             =   1410
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton Exportar 
         Caption         =   "Exportar e Gerar arquivo antes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         TabIndex        =   24
         Top             =   870
         Width           =   1770
      End
      Begin VB.OptionButton Importar 
         Caption         =   "Importar e Atualizar dados antes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         TabIndex        =   23
         Top             =   270
         Width           =   1770
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
      Height          =   2865
      Left            =   105
      TabIndex        =   8
      Top             =   975
      Width           =   4350
      Begin VB.CheckBox SoErros 
         Caption         =   "Exibe somente os erros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   255
         TabIndex        =   21
         Top             =   2100
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tabelas"
         Height          =   825
         Left            =   105
         TabIndex        =   16
         Top             =   330
         Width           =   4140
         Begin VB.ComboBox Tabelas 
            Height          =   315
            ItemData        =   "RelOpIntLog.ctx":0000
            Left            =   1740
            List            =   "RelOpIntLog.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   255
            Width           =   2310
         End
         Begin VB.CheckBox Todas 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   30
         End
         Begin VB.Label Label4 
            Caption         =   "Arquivos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   19
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data"
         Height          =   780
         Left            =   105
         TabIndex        =   9
         Top             =   1275
         Width           =   4140
         Begin MSComCtl2.UpDown UpDownDtIni 
            Height          =   315
            Left            =   1545
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataInicial 
            Height          =   300
            Left            =   570
            TabIndex        =   11
            Top             =   315
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDtFim 
            Height          =   315
            Left            =   3450
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataFinal 
            Height          =   300
            Left            =   2490
            TabIndex        =   13
            Top             =   315
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2115
            TabIndex        =   15
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   14
            Top             =   330
            Width           =   345
         End
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpIntLog.ctx":0004
      Left            =   930
      List            =   "RelOpIntLog.ctx":0006
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   2730
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
      Left            =   4530
      Picture         =   "RelOpIntLog.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4350
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpIntLog.ctx":010A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpIntLog.ctx":0288
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpIntLog.ctx":07BA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpIntLog.ctx":0944
         Style           =   1  'Graphical
         TabIndex        =   3
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
      Left            =   240
      TabIndex        =   7
      Top             =   420
      Width           =   615
   End
End
Attribute VB_Name = "RelOpIntLogOcx"
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

Dim gcolArquivo As Collection

Public Sub Form_Load()

Dim lErro As Long
Dim objTipoArq As ClassTipoArqIntegracao
Dim colArquivos As New Collection

On Error GoTo Erro_Form_Load

    Todas.Value = vbChecked
    Tabelas.Enabled = False
    
    Tabelas.Clear
    
    lErro = CF("TipoArqIntegracao_Le_Todos", colArquivos)
    If lErro <> SUCESSO Then gError 189967
    
    For Each objTipoArq In colArquivos
        Tabelas.AddItem objTipoArq.sDescricao
        Tabelas.ItemData(Tabelas.NewIndex) = objTipoArq.iCodigo
    Next
    
    Set gcolArquivo = colArquivos

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 189967

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189968)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 189969
    
    Importar.Value = False
    Exportar.Value = False
    Nomal.Value = True
      
    lErro = objRelOpcoes.ObterParametro("NTIPOARQ", sParam)
    If lErro Then gError 189970
    
    If StrParaInt(sParam) = 0 Then
        Todas.Value = vbChecked
        Tabelas.ListIndex = -1
    Else
        Todas.Value = vbUnchecked
        Call Combo_Seleciona_ItemData(Tabelas, StrParaInt(sParam))
    End If
         
    lErro = objRelOpcoes.ObterParametro("NIMPORTAR", sParam)
    If lErro Then gError 189971
    
    If StrParaInt(sParam) = MARCADO Then
        Importar.Value = True
    Else
        Importar.Value = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NEXPORTAR", sParam)
    If lErro Then gError 189972
    
    If StrParaInt(sParam) = MARCADO Then
        Exportar.Value = True
    Else
        Exportar.Value = False
    End If
    
    lErro = objRelOpcoes.ObterParametro("NERROS", sParam)
    If lErro Then gError 189973
    
    If StrParaInt(sParam) = MARCADO Then
        SoErros.Value = vbChecked
    Else
        SoErros.Value = vbUnchecked
    End If
           
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 189974

    Call DateParaMasked(DataInicial, StrParaDate(sParam))
 
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 189975

    Call DateParaMasked(DataFinal, StrParaDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 189969 To 189975

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189976)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
  
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set gcolArquivo = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 189977
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 189978
    
    If objRelatorio.sOpcao = "?*?" Then
        Importar.Value = True
    End If
  
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
                
        Case 189977
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 189978
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189979)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros(iTabela As Integer, iImportar As Integer, iExportar As Integer, iSoErros As Integer, sAviso As String) As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
             
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 189980
    
    End If
    
    If Todas.Value = vbChecked Then
        iTabela = 0
    Else
        iTabela = Tabelas.ItemData(Tabelas.ListIndex)
    End If
    
    If Importar.Value Then
        iImportar = MARCADO
    Else
        iImportar = DESMARCADO
    End If
    
    If Exportar.Value Then
        iExportar = MARCADO
    Else
        iExportar = DESMARCADO
    End If
    
    If SoErros.Value = vbChecked Then
        iSoErros = MARCADO
        sAviso = "Só erros"
    Else
        iSoErros = DESMARCADO
        sAviso = "Erros e Avisos"
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 189980
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189981)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 189982
    
    Todas.Value = vbChecked
    Tabelas.Enabled = False
    Importar.Value = False
    Exportar.Value = False
    Nomal.Value = True
    SoErros.Value = vbUnchecked
    
    If ComboOpcoes.Visible Then
        ComboOpcoes.Text = ""
        ComboOpcoes.SetFocus
    End If
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 189982
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189983)

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
Dim iTabela As Integer
Dim iImportar As Integer
Dim iExportar As Integer
Dim iSoErros As Integer
Dim sNomeArqParam As String
Dim sAviso As String
Dim objArqExp As New ClassArqExportacaoAux
Dim sDiretorio As String
Dim lRetorno As Long
Dim colArquivos As New Collection
Dim colArquivosAux As New Collection
Dim colCampos As New Collection
Dim objFolder As Folder
Dim objFile As File
Dim objTipoArq As ClassTipoArqIntegracao
Dim sFileName As String
Dim iIndice As Integer
Dim objArqImpArq As ClassArqImportacaoArq
Dim objArqImp As New ClassArqImportacaoAux
Dim objFSO As New FileSystemObject
Dim bCopiar As Boolean
Dim colTipoArqImp As New Collection
Dim colTipoArqExp As New Collection

On Error GoTo Erro_PreencherRelOp

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Formata_E_Critica_Parametros(iTabela, iImportar, iExportar, iSoErros, sAviso)
    If lErro <> SUCESSO Then gError 189984
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 189985
    
    lErro = objRelOpcoes.IncluirParametro("NTIPOARQ", CStr(iTabela))
    If lErro <> AD_BOOL_TRUE Then gError 189986

    lErro = objRelOpcoes.IncluirParametro("NIMPORTAR", CStr(iImportar))
    If lErro <> AD_BOOL_TRUE Then gError 189987
   
    lErro = objRelOpcoes.IncluirParametro("NEXPORTAR", CStr(iExportar))
    If lErro <> AD_BOOL_TRUE Then gError 189988
   
    lErro = objRelOpcoes.IncluirParametro("NERROS", CStr(iSoErros))
    If lErro <> AD_BOOL_TRUE Then gError 189989
    
    lErro = objRelOpcoes.IncluirParametro("TAVISO", sAviso)
    If lErro <> AD_BOOL_TRUE Then gError 189990
   
    lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(StrParaDate(DataInicial.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 189991

    lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(StrParaDate(DataFinal.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 189992
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 189993
    
    If bExecutando Then
    
        If iImportar = MARCADO Or iExportar = MARCADO Then
        
            lErro = Sistema_Preparar_Batch(sNomeArqParam)
            If lErro <> SUCESSO Then gError 189994
        
            If iImportar = MARCADO Then
            
                lErro = CF("TipoArqIntegracao_Le_Todos", colTipoArqImp, TIPO_INTEGRACAO_IMPORTACAO)
                If lErro <> SUCESSO Then gError 190144
            
                'Obtém o diretório onde estão os arquivos
                sDiretorio = String(512, 0)
                lRetorno = GetPrivateProfileString("Geral", "dirArqImport", "c:\", sDiretorio, 512, "ADM100.INI")
                sDiretorio = left(sDiretorio, lRetorno)
            
                'Pega todos os aquivos da pasta
                Set objFolder = objFSO.GetFolder(sDiretorio)
            
                'Para cada arquivo
                For Each objFile In objFolder.Files
                    'Pega o Nome sem a data
                    sFileName = objFile.Name
                    For iIndice = 1 To Len(objFile.Name)
                        If IsNumeric(Mid(objFile.Name, iIndice, 1)) Then
                            sFileName = left(objFile.Name, iIndice - 1)
                            Exit For
                        End If
                    Next
                    bCopiar = False
                    For Each objTipoArq In colTipoArqImp
                        If objTipoArq.sSiglaArq = sFileName Then
                            bCopiar = True
                            Exit For
                        End If
                    Next
                    
                    If bCopiar Then
                    
                        Set objArqImpArq = New ClassArqImportacaoArq
                        Set objArqImpArq.objFile = objFile
                        Set objArqImpArq.objTipoArq = objTipoArq
                        
                        objArqImpArq.dtData = StrParaDate(Mid(objFile.Name, iIndice, 4) & SEPARADOR & Mid(objFile.Name, iIndice + 4, 2) & SEPARADOR & Mid(objFile.Name, iIndice + 6, 2))
                        objArqImpArq.sNomeArquivo = objFile.Name
                        
                        colArquivosAux.Add objArqImpArq
                    
                    End If
                           
                Next
                
                colCampos.Add "sNomeArquivo"
            
                Call Ordena_Colecao(colArquivosAux, colArquivos, colCampos)

                Set objArqImp.colArquivos = colArquivos
                objArqImp.sDiretorio = sDiretorio
                
                If right(objArqImp.sDiretorio, 1) <> "\" Or right(objArqImp.sDiretorio, 1) <> "/" Then
                    objArqImp.sDiretorio = objArqImp.sDiretorio & "\"
                End If
            
                lErro = CF("Rotina_Importa_Dados", sNomeArqParam, objArqImp)
            Else
            
                lErro = CF("TipoArqIntegracao_Le_Todos", colTipoArqExp, TIPO_INTEGRACAO_EXPORTACAO)
                If lErro <> SUCESSO Then gError 190145
            
                'Obtém o diretório onde estão os arquivos
                sDiretorio = String(512, 0)
                lRetorno = GetPrivateProfileString("Geral", "dirArqExport", "c:\", sDiretorio, 512, "ADM100.INI")
                sDiretorio = left(sDiretorio, lRetorno)
            
                objArqExp.iExportar = EXPORTAR_DADOS_TODOS_NAO_EXPORTADOS
                Set objArqExp.colTiposArq = colTipoArqExp
                
                objArqExp.sDiretorio = sDiretorio
                
                If right(objArqExp.sDiretorio, 1) <> "\" Or right(objArqExp.sDiretorio, 1) <> "/" Then
                    objArqExp.sDiretorio = objArqExp.sDiretorio & "\"
                End If
                
                lErro = CF("Rotina_Exporta_Dados", sNomeArqParam, objArqExp)
            End If
            If lErro <> SUCESSO Then gError 189995
        
        End If
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    GL_objMDIForm.MousePointer = vbDefault

    PreencherRelOp = gErr

    Select Case gErr

        Case 189984 To 189995, 190144, 190145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 189996)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 189997

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 189998

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 189999
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 189997
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 189998, 189999

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190000)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 190001

    If Not Importar.Value And Not Exportar.Value Then
        Call gobjRelatorio.Executar_Prossegue2(Me)
    End If

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 190001

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190002)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 190003

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 190004

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 190005

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 190006
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 190003
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 190004 To 190006

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190007)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

     If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

    End If
    
    If Todas.Value = vbUnchecked Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoArq <= " & Forprint_ConvInt(Tabelas.ItemData(Tabelas.ListIndex))
    
    End If
    
    If SoErros.Value = vbChecked Then
    
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoAviso <= " & Forprint_ConvInt(DESMARCADO)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190008)

    End Select

    Exit Function

End Function

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 190009

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 190009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190010)

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
        If lErro <> SUCESSO Then gError 190011

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 190011

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190012)

    End Select

    Exit Sub

End Sub

Private Sub Todas_Click()

    If Todas.Value = vbChecked Then
        Tabelas.ListIndex = -1
        Tabelas.Enabled = False
    Else
        Tabelas.Enabled = True
    End If

End Sub

Private Sub Todas_Change()

    If Todas.Value = vbChecked Then
        Tabelas.ListIndex = -1
        Tabelas.Enabled = False
    Else
        Tabelas.Enabled = True
    End If

End Sub

Private Sub UpDownDtIni_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtIni_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190013

    Exit Sub

Erro_UpDownDtIni_DownClick:

    Select Case gErr

        Case 190013
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190014)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDtIni_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtIni_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190015

    Exit Sub

Erro_UpDownDtIni_UpClick:

    Select Case gErr

        Case 190015
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190016)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDtFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtFim_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190017

    Exit Sub

Erro_UpDownDtFim_DownClick:

    Select Case gErr

        Case 190017
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190018)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDtFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDtFim_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190019

    Exit Sub

Erro_UpDownDtFim_UpClick:

    Select Case gErr

        Case 190019
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190020)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_NOTAS_REC
    Set Form_Load_Ocx = Me
    Caption = "Log de Atualização de Dados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpIntLog"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
    
    End If

End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
