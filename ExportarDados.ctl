VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ExportarDadosOcx 
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   LockControls    =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8625
   Begin VB.TextBox NomeDiretorio 
      Height          =   285
      Left            =   1050
      TabIndex        =   0
      Top             =   255
      Width           =   4590
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      TabIndex        =   1
      Top             =   210
      Width           =   555
   End
   Begin VB.Frame Frame 
      Caption         =   "Exportar"
      Height          =   2265
      Left            =   225
      TabIndex        =   20
      Top             =   735
      Width           =   8130
      Begin VB.Frame Frame1 
         Caption         =   "Período da modificação"
         Height          =   780
         Left            =   1755
         TabIndex        =   24
         Top             =   1320
         Width           =   5250
         Begin MSComCtl2.UpDown UpDownDataInicio 
            Height          =   300
            Left            =   1560
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   300
            Left            =   540
            TabIndex        =   5
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
         Begin MSComCtl2.UpDown UpDownDataFim 
            Height          =   300
            Left            =   4140
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   300
            Left            =   3135
            TabIndex        =   7
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
         Begin VB.Label LabelFim 
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
            Left            =   2700
            TabIndex        =   26
            Top             =   315
            Width           =   360
         End
         Begin VB.Label LabelInicio 
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
            Left            =   150
            TabIndex        =   25
            Top             =   300
            Width           =   315
         End
      End
      Begin VB.OptionButton OptPeriodo 
         Caption         =   "Apenas registros modificados no período abaixo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   810
         TabIndex        =   4
         Top             =   975
         Width           =   4950
      End
      Begin VB.OptionButton OptTodosNaoExp 
         Caption         =   "Todos registros não exportados anteriormente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   810
         TabIndex        =   3
         Top             =   615
         Width           =   4290
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos os registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   810
         TabIndex        =   2
         Top             =   255
         Width           =   1965
      End
      Begin VB.CheckBox ImprimeEtiquetas 
         Caption         =   "Imprimir Etiquetas:"
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
         Left            =   -20000
         TabIndex        =   22
         Top             =   1860
         Width           =   1935
      End
      Begin VB.TextBox QuantEtiquetas 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -20000
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame FramePrincipal 
      Caption         =   "Seleção de Arquivos"
      Height          =   3210
      Left            =   225
      TabIndex        =   19
      Top             =   3060
      Width           =   8115
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   555
         Left            =   6570
         Picture         =   "ExportarDados.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1005
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   555
         Left            =   6555
         Picture         =   "ExportarDados.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   255
         Width           =   1425
      End
      Begin VB.ComboBox FilialEmpresa 
         Height          =   315
         Left            =   -23270
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1560
         Width           =   1725
      End
      Begin VB.TextBox Item 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1665
         TabIndex        =   16
         Top             =   1215
         Width           =   4710
      End
      Begin VB.CheckBox Selecionado 
         Height          =   255
         Left            =   990
         TabIndex        =   15
         Top             =   1230
         Width           =   650
      End
      Begin MSFlexGridLib.MSFlexGrid GridArquivos 
         Height          =   2325
         Left            =   225
         TabIndex        =   9
         Top             =   255
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   4101
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox NumEtiqueta 
         Height          =   300
         Left            =   -10000
         TabIndex        =   23
         Top             =   1560
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6675
      ScaleHeight     =   495
      ScaleWidth      =   1590
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   1650
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   90
         Picture         =   "ExportarDados.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exporta os arquivos selecionados"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "ExportarDados.ctx":263E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "ExportarDados.ctx":2B70
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Diretório:"
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
      Left            =   225
      TabIndex        =   27
      Top             =   285
      Width           =   795
   End
End
Attribute VB_Name = "ExportarDadosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais
Dim iAlterado As Integer

'Variável utilizada para manuseio do grid
Dim objGridArquivos As AdmGrid

Dim gcolArquivos As Collection

'Variáveis das colunas do grid
Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Item_Col As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long
Dim objTipoArq As ClassTipoArqIntegracao
Dim colArquivos As New Collection
Dim iLinha As Integer
Dim sDiretorio As String
Dim lRetorno As Long

On Error GoTo Erro_Form_Load

    'instancia as variáveis globais
    Set objGridArquivos = New AdmGrid
    
    OptTodosNaoExp.Value = True
    
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "dirArqExport", "c:\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = Left(sDiretorio, lRetorno)
    
    NomeDiretorio.Text = sDiretorio
    Call NomeDiretorio_Validate(bSGECancelDummy)
    
    'Inicializa o Grid
    lErro = Inicializa_Grid_Arquivos(objGridArquivos)
    If lErro <> SUCESSO Then gError 189777
    
    lErro = CF("TipoArqIntegracao_Le_Todos", colArquivos, TIPO_INTEGRACAO_EXPORTACAO)
    If lErro <> SUCESSO Then gError 189778
    
    Set gcolArquivos = colArquivos
    
    iLinha = 0
    For Each objTipoArq In colArquivos
        iLinha = iLinha + 1
        GridArquivos.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(MARCADO)
        GridArquivos.TextMatrix(iLinha, iGrid_Item_Col) = objTipoArq.sDescricao
    Next
    
    objGridArquivos.iLinhasExistentes = colArquivos.Count
    
    Call Grid_Refresh_Checkbox(objGridArquivos)
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 189777, 189778

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189779)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long
'A tela não espera recebimento de parâmetros, portanto, essa função sempre retorna sucesso
    Trata_Parametros = SUCESSO
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Libera os objetos e coleções globais
    Set objGridArquivos = Nothing
    Set gcolArquivos = Nothing

End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoGerar_Click()
'Dispara a geração dos arquivos e relatórios selecionados

Dim lErro As Long
Dim sNomeArqParam As String
Dim objArqExp As New ClassArqExportacaoAux

On Error GoTo Erro_BotaoGerar_Click

    lErro = Move_Tela_Memoria(objArqExp)
    If lErro <> SUCESSO Then gError 189780
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 189865
    If objArqExp.colTiposArq.Count = 0 Then gError 189866
    
    If objArqExp.dtExpDataAte <> DATA_NULA And objArqExp.dtExpDataDe <> DATA_NULA Then
        If objArqExp.dtExpDataDe > objArqExp.dtExpDataAte Then gError 189867
    End If
    
    'prepara o sistema para trabalhar com rotina batch
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 189781
    
    'inicia o Batch
    lErro = CF("Rotina_Exporta_Dados", sNomeArqParam, objArqExp)
    If lErro <> SUCESSO Then gError 189782
    
    Exit Sub

Erro_BotaoGerar_Click:

    Select Case gErr
    
        Case 189780 To 189782
               
        Case 189865
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeDiretorio.SetFocus
            
        Case 189866
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ARQUIVO_SELECIONADO", gErr)
    
        Case 189867
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            DataDe.SetFocus
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189783)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se algum campo foi alterado e confirma se o usuário deseja
    'salvar antes de limpar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 189784

    'limpa a tela
    Call Limpa_Tela_ExportarDados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 189784
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189785)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Se a data está preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 189786

    End If

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 189786

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189787)
    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Se a data está preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 189788

    End If

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 189788

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189789)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Se a data está preenchida
    If Len(Trim(DataAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 189790

    End If

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 189790

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189791)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Se a data está preenchida
    If Len(Trim(DataAte.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 189792

    End If

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 189792

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189793)

    End Select

    Exit Sub

End Sub

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub DataDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub DataAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO ***

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a data digitada é válida
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 189794

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 189794
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189795)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a data digitada é válida
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 189796

Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 189796
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189797)
            
    End Select
    
    Exit Sub

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** FUNCIONAMENTO DO GridArquivos - INÍCIO ***

'***** EVENTOS DO GRID - INÍCIO *******
Private Sub GridArquivos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridArquivos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridArquivos, iAlterado)
    End If

End Sub

Private Sub GridArquivos_EnterCell()
    Call Grid_Entrada_Celula(objGridArquivos, iAlterado)
End Sub

Private Sub GridArquivos_GotFocus()
    Call Grid_Recebe_Foco(objGridArquivos)
End Sub

Private Sub GridArquivos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridArquivos)
End Sub

Private Sub GridArquivos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridArquivos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridArquivos, iAlterado)
    End If

End Sub

Private Sub GridArquivos_LeaveCell()
    Call Saida_Celula(objGridArquivos)
End Sub

Private Sub GridArquivos_RowColChange()
    Call Grid_RowColChange(objGridArquivos)
End Sub

Private Sub GridArquivos_Scroll()
    Call Grid_Scroll(objGridArquivos)
End Sub

Private Sub GridArquivos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridArquivos)
End Sub
'***** EVENTOS DO GRID - FIM *******

'**** EVENTOS DOS CONTROLES DO GRID - INÍCIO *********
Private Sub Selecionado_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Selecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArquivos)
End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArquivos)
End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArquivos.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridArquivos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Item_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Item_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridArquivos)
End Sub

Private Sub Item_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridArquivos)
End Sub

Private Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridArquivos.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGridArquivos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'**** EVENTOS DOS CONTROLES DO GRID - FIM *********

'**** SAÍDA DE CÉLULA DO GRID E DOS CONTROLES - INÍCIO ******
Public Function Saida_Celula(objGridArquivos As AdmGrid) As Long
'faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridArquivos)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridArquivos.objGrid.Col

            Case iGrid_Selecionado_Col
                lErro = Saida_Celula_Selecionado(objGridArquivos)
                If lErro <> SUCESSO Then gError 189798


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridArquivos)
        If lErro <> SUCESSO Then gError 189799

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 189798 To 189799

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189800)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Selecionado(objGridArquivos As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Selecionado

    Set objGridArquivos.objControle = Selecionado

    lErro = Grid_Abandona_Celula(objGridArquivos)
    If lErro <> SUCESSO Then gError 189801

    Saida_Celula_Selecionado = SUCESSO

    Exit Function

Erro_Saida_Celula_Selecionado:

    Saida_Celula_Selecionado = gErr

    Select Case gErr

        Case 189801

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189802)

    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(ByVal objArqExp As ClassArqExportacaoAux) As Long
'Transfere os dados tela para objIN86Modelo

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Move_Tela_Memoria

    If OptTodos.Value Then
        objArqExp.iExportar = EXPORTAR_DADOS_TODOS
    ElseIf OptTodosNaoExp.Value Then
        objArqExp.iExportar = EXPORTAR_DADOS_TODOS_NAO_EXPORTADOS
    ElseIf OptPeriodo.Value Then
        objArqExp.iExportar = EXPORTAR_DADOS_POR_PERIODO
    End If
    
    objArqExp.dtExpDataDe = StrParaDate(DataDe.Text)
    objArqExp.dtExpDataAte = StrParaDate(DataAte.Text)
    
    For iLinha = 1 To objGridArquivos.iLinhasExistentes
        If StrParaInt(GridArquivos.TextMatrix(iLinha, iGrid_Selecionado_Col)) = MARCADO Then
            objArqExp.colTiposArq.Add gcolArquivos.Item(iLinha)
        End If
    Next
    
    objArqExp.sDiretorio = NomeDiretorio.Text
    
'    If Right(objArqExp.sDiretorio, 1) <> "\" Or Right(objArqExp.sDiretorio, 1) <> "/" Then
'        objArqExp.sDiretorio = objArqExp.sDiretorio & "\"
'    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189803)

    End Select

End Function

Private Sub Limpa_Tela_ExportarDados()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ExportarDados

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    OptTodosNaoExp.Value = True

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_ExportarDados:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189804)

    End Select
    
    Exit Sub
    
End Sub

Private Function Inicializa_Grid_Arquivos(objGridArquivos As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Arquivos

    Set objGridArquivos.objForm = Me

    'entitula as colunas
    objGridArquivos.colColuna.Add ""
    objGridArquivos.colColuna.Add "Gerar"
    objGridArquivos.colColuna.Add "Arquivo"

    'guarda os nomes dos campos
    objGridArquivos.colCampo.Add Selecionado.Name
    objGridArquivos.colCampo.Add Item.Name

    'inicializa os índices das colunas
    iGrid_Selecionado_Col = 1
    iGrid_Item_Col = 2

    'configura os atributos
    GridArquivos.ColWidth(0) = 300
    GridArquivos.Rows = 20

    'vincula o grid da tela propriamente dito ao controlador de grid
    objGridArquivos.objGrid = GridArquivos

    'configura sua visualização
    objGridArquivos.iLinhasVisiveis = 8
    objGridArquivos.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridArquivos.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridArquivos.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'inicializa o grid
    Call Grid_Inicializa(objGridArquivos)

    Inicializa_Grid_Arquivos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Arquivos:

    Inicializa_Grid_Arquivos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189805)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração dos arquivos de exportação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ExportarDados"

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property
'**** fim do trecho a ser copiado *****

'*** TRATAMENTO PARA MODO DE EDIÇÃO - INÍCIO ***
Private Sub LabelInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelInicio, Button, Shift, X, Y)
End Sub

Private Sub LabelInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelInicio, Source, X, Y)
End Sub

Private Sub LabelFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFim, Button, Shift, X, Y)
End Sub

Private Sub LabelFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFim, Source, X, Y)
End Sub

'*** TRATAMENTO PARA MODO DE EDIÇÃO - FIM ***

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPOS As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If Right(NomeDiretorio.Text, 1) <> "\" And Right(NomeDiretorio.Text, 1) <> "/" Then
        iPOS = InStr(1, NomeDiretorio.Text, "/")
        If iPOS = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 189870

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 189870, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189871)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "This is the title"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189881)

    End Select

    Exit Sub
  
End Sub

Private Sub Marca_Desmarca(ByVal bMarca As Boolean, ByVal objGridInt As AdmGrid, ByVal iColuna As Integer)

Dim iLinha As Integer

    For iLinha = 1 To objGridInt.iLinhasExistentes
        If bMarca Then
            objGridInt.objGrid.TextMatrix(iLinha, iColuna) = CStr(MARCADO)
        Else
            objGridInt.objGrid.TextMatrix(iLinha, iColuna) = CStr(DESMARCADO)
        End If
    Next
    
    Call Grid_Refresh_Checkbox(objGridInt)

End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Marca_Desmarca(True, objGridArquivos, iGrid_Selecionado_Col)
End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call Marca_Desmarca(False, objGridArquivos, iGrid_Selecionado_Col)
End Sub
