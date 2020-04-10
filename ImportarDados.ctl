VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ImportarDadosOcx 
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   LockControls    =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8625
   Begin VB.CheckBox optValidacaoManual 
      Caption         =   "Perguntar se deseja manter ou atualizar registros já existentes"
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
      Left            =   915
      TabIndex        =   17
      Top             =   540
      Width           =   5745
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
      Left            =   5625
      TabIndex        =   1
      Top             =   105
      Width           =   555
   End
   Begin VB.TextBox NomeDiretorio 
      Height          =   285
      Left            =   915
      TabIndex        =   0
      Top             =   150
      Width           =   4590
   End
   Begin VB.Frame FramePrincipal 
      Caption         =   "Seleção de Arquivos"
      Height          =   5475
      Left            =   225
      TabIndex        =   12
      Top             =   885
      Width           =   8190
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   555
         Left            =   1845
         Picture         =   "ImportarDados.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4830
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   555
         Left            =   195
         Picture         =   "ImportarDados.ctx":11E2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4830
         Width           =   1425
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   225
         Left            =   3375
         TabIndex        =   14
         Top             =   1935
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.TextBox TotReg 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   15
         Top             =   1215
         Width           =   1230
      End
      Begin VB.ComboBox FilialEmpresa 
         Height          =   315
         Left            =   -23270
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   1725
      End
      Begin VB.TextBox Item 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1665
         TabIndex        =   9
         Top             =   1215
         Width           =   3360
      End
      Begin VB.CheckBox Selecionado 
         Height          =   255
         Left            =   510
         TabIndex        =   8
         Top             =   1230
         Width           =   1110
      End
      Begin MSFlexGridLib.MSFlexGrid GridArquivos 
         Height          =   2325
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   4101
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox NumEtiqueta 
         Height          =   300
         Left            =   -10000
         TabIndex        =   13
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
      Left            =   6735
      ScaleHeight     =   495
      ScaleWidth      =   1590
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   30
      Width           =   1650
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   90
         Picture         =   "ImportarDados.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Importa os arquivos selecionados"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "ImportarDados.ctx":263E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "ImportarDados.ctx":2B70
         Style           =   1  'Graphical
         TabIndex        =   7
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
      Left            =   90
      TabIndex        =   16
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "ImportarDadosOcx"
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

Dim sNomeArqAnt As String

'Variáveis das colunas do grid
Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_TotalRegistros_Col As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long
Dim sDiretorio As String
Dim lRetorno As Long

On Error GoTo Erro_Form_Load

    'instancia as variáveis globais
    Set objGridArquivos = New AdmGrid
    
    'Inicializa o Grid
    lErro = Inicializa_Grid_Arquivos(objGridArquivos)
    If lErro <> SUCESSO Then gError 189882
    
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "dirArqImport", "c:\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = left(sDiretorio, lRetorno)
    
    NomeDiretorio.Text = sDiretorio
    Call NomeDiretorio_Validate(bSGECancelDummy)
       
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 189882

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189883)

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

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub BotaoGerar_Click()
'Dispara a geração dos arquivos e relatórios selecionados

Dim lErro As Long
Dim sNomeArqParam As String
Dim objArqImp As New ClassArqImportacaoAux

On Error GoTo Erro_BotaoGerar_Click

    lErro = Move_Tela_Memoria(objArqImp)
    If lErro <> SUCESSO Then gError 189884
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 189885
    'If objArqImp.colArquivos.Count = 0 Then gError 189886
    
    'prepara o sistema para trabalhar com rotina batch
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 189887
    
    'inicia o Batch
    lErro = CF("Rotina_Importa_Dados", sNomeArqParam, objArqImp)
    'Força a releitura dos arquivos pois parte já pode ter sido importada
    'e outra parte já ter sido transferida para uma pasta de arquivos temporários (para não tratar arquivos remotos)
    sNomeArqAnt = ""
    Call NomeDiretorio_Validate(bSGECancelDummy)
    If lErro <> SUCESSO Then gError 189888
    
    Exit Sub

Erro_BotaoGerar_Click:

    Select Case gErr
    
        Case 189884, 189887, 189888
        
        Case 189885
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeDiretorio.SetFocus
            
        Case 189886
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ARQUIVO_SELECIONADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189889)
            
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
    If lErro <> SUCESSO Then gError 189890

    'limpa a tela
    Call Limpa_Tela_ImportarDados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 189890
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189891)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

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
                If lErro <> SUCESSO Then gError 189892


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridArquivos)
        If lErro <> SUCESSO Then gError 189893

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 189892 To 189893

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189894)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Selecionado(objGridArquivos As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Selecionado

    Set objGridArquivos.objControle = Selecionado

    lErro = Grid_Abandona_Celula(objGridArquivos)
    If lErro <> SUCESSO Then gError 189895

    Saida_Celula_Selecionado = SUCESSO

    Exit Function

Erro_Saida_Celula_Selecionado:

    Saida_Celula_Selecionado = gErr

    Select Case gErr

        Case 189895

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189896)

    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(ByVal objArqImp As ClassArqImportacaoAux) As Long
'Transfere os dados tela para objIN86Modelo

Dim lErro As Long
Dim iLinha As Integer
Dim objArqImpArq As ClassArqImportacaoArq

On Error GoTo Erro_Move_Tela_Memoria
    
    For iLinha = 1 To objGridArquivos.iLinhasExistentes
        If StrParaInt(GridArquivos.TextMatrix(iLinha, iGrid_Selecionado_Col)) = MARCADO Then
            objArqImp.colArquivos.Add gcolArquivos.Item(iLinha)
        End If
    Next
    
    If optValidacaoManual.Value = vbChecked Then
        objArqImp.iValidacaoManual = MARCADO
    Else
        objArqImp.iValidacaoManual = DESMARCADO
    End If
    
    objArqImp.sDiretorio = NomeDiretorio.Text
    
'    If Right(objArqImp.sDiretorio, 1) <> "\" Or Right(objArqImp.sDiretorio, 1) <> "/" Then
'        objArqImp.sDiretorio = objArqImp.sDiretorio & "\"
'    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189897)

    End Select

End Function

Private Sub Limpa_Tela_ImportarDados()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ImportarDados

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridArquivos)
    
    optValidacaoManual.Value = vbUnchecked
    
    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_ImportarDados:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189898)

    End Select
    
    Exit Sub
    
End Sub

Private Function Inicializa_Grid_Arquivos(objGridArquivos As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Arquivos

    Set objGridArquivos.objForm = Me

    'entitula as colunas
    objGridArquivos.colColuna.Add ""
    objGridArquivos.colColuna.Add "Importar"
    objGridArquivos.colColuna.Add "Arquivo"
    objGridArquivos.colColuna.Add "Data"
    objGridArquivos.colColuna.Add "Total Registros"

    'guarda os nomes dos campos
    objGridArquivos.colCampo.Add Selecionado.Name
    objGridArquivos.colCampo.Add Item.Name
    objGridArquivos.colCampo.Add Data.Name
    objGridArquivos.colCampo.Add TotReg.Name

    'inicializa os índices das colunas
    iGrid_Selecionado_Col = 1
    iGrid_Item_Col = 2
    iGrid_Data_Col = 3
    iGrid_TotalRegistros_Col = 4

    'configura os atributos
    GridArquivos.ColWidth(0) = 300
    
    GridArquivos.Rows = 100 + 1

    'vincula o grid da tela propriamente dito ao controlador de grid
    objGridArquivos.objGrid = GridArquivos

    'configura sua visualização
    objGridArquivos.iLinhasVisiveis = 13
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189899)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Importação de arquivos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ImportarDados"

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

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colArquivos As New Collection
Dim colArquivosAux As New Collection
Dim colCampos As New Collection
Dim objFolder As Folder
Dim objFile As File
Dim colTiposArq As New Collection
Dim objTipoArq As ClassTipoArqIntegracao
Dim sFileName As String
Dim bCopiar As Boolean
Dim objArqImpArq As ClassArqImportacaoArq
Dim objTS As TextStream
Dim objFSO As New FileSystemObject
Dim iIndice As Integer
Dim sRegistro As String
Dim iLinha As Integer
Dim bArqJaImp As Boolean
Dim iPos As Integer
Dim sNomeArqID As String

On Error GoTo Erro_NomeDiretorio_Validate

    If UCase(sNomeArqAnt) = UCase(NomeDiretorio.Text) Then Exit Sub
    
    Call Grid_Limpa(objGridArquivos)
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub

    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorio.Text, "/")
        If iPos = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If
    
    sNomeArqAnt = NomeDiretorio.Text
    
    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 189900
    
    lErro = CF("TipoArqIntegracao_Le_Todos", colTiposArq, TIPO_INTEGRACAO_IMPORTACAO)
    If lErro <> SUCESSO Then gError 189901
    
    'Pega todos os aquivos da pasta
    Set objFolder = objFSO.GetFolder(NomeDiretorio.Text)

    'Para cada arquivo
    For Each objFile In objFolder.Files
    
        'Pega o Nome sem a data
        sNomeArqID = objFile.Name
        bArqJaImp = False
        
        If UCase(right(objFile.Name, 4)) = ".TXT" Then
        
            Set objTS = objFile.OpenAsTextStream(ForReading)
            
            sRegistro = objTS.ReadLine
                    
            lErro = CF("Integracao_Obtem_NomeID_Cust", sRegistro, sNomeArqID)
            If lErro <> SUCESSO Then gError 189901
            
            iPos = InStr(1, sNomeArqID, ".")
            If iPos <> 0 Then
                sFileName = left(sNomeArqID, iPos - 1)
            Else
                sFileName = sNomeArqID
            End If
            
    '        For iIndice = 1 To Len(objFile.Name)
    '            If IsNumeric(Mid(objFile.Name, iIndice, 14)) Then
    '                sFileName = Left(objFile.Name, iIndice - 1)
    '                Exit For
    '            End If
    '        Next
    
            If Len(sFileName) > 14 Then
                If IsNumeric(right(sFileName, 14)) Then
                    iIndice = Len(sFileName) - 14 + 1
                    sFileName = left(sFileName, iIndice - 1)
                End If
            End If
            
            bCopiar = False
            For Each objTipoArq In colTiposArq
                If objTipoArq.sSiglaArq = sFileName Then
                    bCopiar = True
                    Exit For
                End If
            Next
            
            If bCopiar Then
                
                lErro = CF("ArqImport_Verifica_JaImportado", sNomeArqID, bArqJaImp)
                If lErro <> SUCESSO Then gError 189902
                
            End If
            
            If bArqJaImp Then bCopiar = False
            
            If bCopiar Then
            
                Set objArqImpArq = New ClassArqImportacaoArq
                Set objArqImpArq.objFile = objFile
                Set objArqImpArq.objTipoArq = objTipoArq
                
                objArqImpArq.dtData = StrParaDate(Mid(sNomeArqID, iIndice, 4) & SEPARADOR & Mid(sNomeArqID, iIndice + 4, 2) & SEPARADOR & Mid(sNomeArqID, iIndice + 6, 2))
    
                Set objTS = objFile.OpenAsTextStream(ForReading)
                
    '            Do While Not objTS.AtEndOfLine
    '                sRegistro = objTS.ReadLine
    '            Loop
                
                Do While Not objTS.AtEndOfLine
                    sRegistro = objTS.ReadLine
                    Exit Do
                Loop
                
                'objArqImpArq.lTotalRegistros = StrParaLong(Mid(sRegistro, 3, 6))
                objArqImpArq.lTotalRegistros = objFile.Size / Len(sRegistro)
                objArqImpArq.sNomeArquivo = sNomeArqID
                
                colArquivosAux.Add objArqImpArq
            
            End If
            
        ElseIf UCase(right(objFile.Name, 4)) = ".XLS" Then
        
            bCopiar = False
            For Each objTipoArq In colTiposArq
                If objTipoArq.sSiglaArq = "Cli" And InStr(1, sNomeArqID, "Cliente") <> 0 Then
                    bCopiar = True
                    Exit For
                End If
            Next
            If Not bCopiar Then
                For Each objTipoArq In colTiposArq
                    If objTipoArq.sSiglaArq = "PedVend" And InStr(1, sNomeArqID, "Pedido") <> 0 Then
                        bCopiar = True
                        Exit For
                    End If
                Next
            End If
            
            If bCopiar Then
                
                lErro = CF("ArqImport_Verifica_JaImportado", sNomeArqID, bArqJaImp)
                If lErro <> SUCESSO Then gError 189902
                
            End If
            
            If bArqJaImp Then bCopiar = False
        
            If bCopiar Then
                Set objArqImpArq = New ClassArqImportacaoArq
                Set objArqImpArq.objFile = objFile
                Set objArqImpArq.objTipoArq = objTipoArq
                objArqImpArq.dtData = objFile.DateCreated
                objArqImpArq.lTotalRegistros = 0
                objArqImpArq.sNomeArquivo = sNomeArqID
                
                colArquivosAux.Add objArqImpArq
            End If
        
        End If
               
    Next
    
    colCampos.Add "sNomeArquivo"

    Call Ordena_Colecao(colArquivosAux, colArquivos, colCampos)
    
    iLinha = 0
    For Each objArqImpArq In colArquivos
        iLinha = iLinha + 1
        GridArquivos.TextMatrix(iLinha, iGrid_Selecionado_Col) = CStr(MARCADO)
        GridArquivos.TextMatrix(iLinha, iGrid_Item_Col) = objArqImpArq.sNomeArquivo
        GridArquivos.TextMatrix(iLinha, iGrid_Data_Col) = Format(objArqImpArq.dtData, "DD/MM/YYYY")
        GridArquivos.TextMatrix(iLinha, iGrid_TotalRegistros_Col) = CStr(objArqImpArq.lTotalRegistros)
    Next
    
    objGridArquivos.iLinhasExistentes = colArquivos.Count
    
    Call Grid_Refresh_Checkbox(objGridArquivos)
    
    Set gcolArquivos = colArquivos
        
    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 189900, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case 189901, 189902

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189903)

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
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189904)

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
