VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ConfigOutrosOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8175
      ScaleHeight     =   495
      ScaleWidth      =   1155
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   1215
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ConfigOutrosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ConfigOutrosOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Modulo 
      Height          =   315
      ItemData        =   "ConfigOutrosOcx.ctx":02D8
      Left            =   1275
      List            =   "ConfigOutrosOcx.ctx":02DA
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   180
      Width           =   6405
   End
   Begin VB.Frame Frame 
      Caption         =   "Configurações"
      Height          =   5415
      Left            =   30
      TabIndex        =   1
      Top             =   555
      Width           =   9360
      Begin VB.CommandButton BotaoProcuraDir 
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
         Left            =   7950
         TabIndex        =   10
         Top             =   -15
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton BotaoProcuraFile 
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
         Left            =   8730
         TabIndex        =   9
         Top             =   -15
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox Descricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   555
         MaxLength       =   250
         TabIndex        =   8
         Top             =   555
         Width           =   5075
      End
      Begin VB.ComboBox Valor 
         Height          =   315
         Left            =   5475
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   525
         Width           =   3160
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   4935
         Left            =   75
         TabIndex        =   0
         Top             =   345
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7665
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Módulos:"
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
      Left            =   390
      TabIndex        =   2
      Top             =   225
      Width           =   765
   End
End
Attribute VB_Name = "ConfigOutrosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

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

Dim gcolModulos As New Collection
Dim gcolConfigs As New Collection

Public objGridItens As AdmGrid
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Valor_Col As Integer

Public iAlterado As Integer
Public iAlteradoAux As Integer
Public sModuloAnt As String
Dim iLinhaAnterior As Integer

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Outras Configurações"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ConfigOutros"

End Function

Public Sub Show()
'    Me.Show
'    Parent.SetFocus
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
    End If

End Sub

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

Public Sub Form_Load()

Dim lErro As Long
Dim objModulo As AdmModulo

On Error GoTo Erro_Form_Load

    Set objGridItens = New AdmGrid
    
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Modulos_Le_Todos", gcolModulos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Modulo.Clear
    For Each objModulo In gcolModulos
        Modulo.AddItem objModulo.sNome
    Next
    Modulo.ListIndex = -1

    iAlterado = 0
    iAlteradoAux = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213377)

    End Select

    iAlterado = 0
    iAlteradoAux = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal sModulo As String) As Long

Dim lErro As Long
Dim iIndice As Integer, bAchou As Boolean
Dim objModulo As AdmModulo

On Error GoTo Erro_Trata_Parametros

    If sModulo <> "" Then
    
        iIndice = -1
        bAchou = False
        For Each objModulo In gcolModulos
            iIndice = iIndice + 1
            If objModulo.sSigla = sModulo Then
                bAchou = True
                Exit For
            End If
        Next
        If bAchou Then
            Modulo.ListIndex = iIndice
        End If

    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213378)
    
    End Select
    
    Exit Function
        
End Function

Function Saida_Celula(objGridItens As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridItens)

    If lErro = SUCESSO Then

        Select Case GridItens.Col

            Case iGrid_Valor_Col

                lErro = Saida_Celula_Valor(objGridItens)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridItens)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213379)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
        
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213380)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_Configs_Memoria()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("ConfigOutros_Grava", gcolConfigs)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    iAlteradoAux = 0
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213381)

    End Select

    Exit Function

End Function

Private Sub Valor_Change()
    iAlteradoAux = REGISTRO_ALTERADO
End Sub

Private Sub Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    Set gcolModulos = Nothing
    Set gcolConfigs = Nothing
    
End Sub

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long

Dim lErro As Long, bAchou As Boolean
Dim objConfig As ClassConfigOutros
Dim objConfigVV As ClassConfigValoresValidos
Dim dValorAux As Double, iPos As Integer

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor
    
    If UCase(Trim(Valor.Text)) <> UCase(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col))) Then
        iAlterado = REGISTRO_ALTERADO
        
        Set objConfig = gcolConfigs.Item(GridItens.Row)
        
        If Len(Trim(Valor.Text)) > 0 Then
        
            Select Case UCase(Trim(objConfig.sTipoControle))
            
                Case UCase("ComboBox")
                    bAchou = False
                    For Each objConfigVV In objConfig.colValores
                        If Codigo_Extrai(Valor.Text) = objConfigVV.iSeq Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then gError 213382 'Não existe opções válidas
            
                Case UCase("Integer")
                
                    If Trim(Valor.Text) <> "0" Then lErro = Inteiro_Critica(Valor.Text)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                Case UCase("Long")
            
                    If Trim(Valor.Text) <> "0" Then lErro = Long_Critica(Valor.Text)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                Case UCase("Float")
            
                    lErro = Valor_NaoNegativo_Critica(Valor.Text)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    Valor.Text = Format(Valor.Text, "STANDARD")
            
                Case UCase("Percent")
                        
                    'Critica a porcentagem
                    lErro = Porcentagem_Critica(Valor.Text)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                    dValorAux = CDbl(Valor.Text)
            
                    Valor.Text = Format(dValorAux, "Fixed")
                    
                Case UCase("String")
                
                    If Len(Trim(Valor.Text)) > STRING_CONFIG_CONTEUDO Then gError 213383
            
                Case UCase("Date")
                
                    lErro = Data_Critica(Valor.Text)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                Case UCase("Dir")
                
                    If Len(Trim(Valor.Text)) > STRING_CONFIG_CONTEUDO Then gError 213383
                
                    If right(Valor.Text, 1) <> "\" And right(Valor.Text, 1) <> "/" Then
                        iPos = InStr(1, Valor.Text, "/")
                        If iPos = 0 Then
                            Valor.Text = Valor.Text & "\"
                        Else
                            Valor.Text = Valor.Text & "/"
                        End If
                    End If
                
                    If Len(Trim(Dir(Valor.Text, vbDirectory))) = 0 Then gError 213384
            
                Case UCase("File")
            
                    If Len(Trim(Valor.Text)) > STRING_CONFIG_CONTEUDO Then gError 213383
            
            End Select
            
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182568

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr
    
        Case 213382
            Call Rotina_Erro(vbOKOnly, "ERRO_COMBO_VALOR_INVALIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 213383
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_COM_TAM_MAIOR_PERM", gErr, "Opção", Valor.Text, Len(Trim(Valor.Text)), STRING_CONFIG_CONTEUDO)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 213384, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, Valor.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213385)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Opção")
    

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Descricao.Name)
    objGridInt.colCampo.Add (Valor.Name)

    iGrid_Descricao_Col = 1
    iGrid_Valor_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 13

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlteradoAux)
    End If

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlteradoAux)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlteradoAux)
    End If

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)
    
    Call Trata_Controles
    
End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer

    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Private Function Traz_Configs_Tela(ByVal sModulo As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objConfig As ClassConfigOutros
Dim objConfigVV As ClassConfigValoresValidos
Dim sValor As String

On Error GoTo Erro_Traz_Configs_Tela

    Set gcolConfigs = New Collection
    
    Call Grid_Limpa(objGridItens)
    
    lErro = CF("ConfigOutros_Le", sModulo, gcolConfigs)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objConfig In gcolConfigs
            
        iIndice = iIndice + 1
        
        GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = objConfig.sDescricaoGrid
        
        sValor = objConfig.sConteudo
        For Each objConfigVV In objConfig.colValores
            If UCase(Trim(sValor)) = UCase(Trim(objConfigVV.sValor)) Then
                sValor = CStr(objConfigVV.iSeq) & SEPARADOR & objConfigVV.sDescricao
            End If
        Next

        GridItens.TextMatrix(iIndice, iGrid_Valor_Col) = sValor

    Next
    
    objGridItens.iLinhasExistentes = gcolConfigs.Count

    Traz_Configs_Tela = SUCESSO

    Exit Function

Erro_Traz_Configs_Tela:

    Traz_Configs_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213386)

    End Select

    Exit Function

End Function

Private Function Move_Configs_Memoria() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objConfig As ClassConfigOutros
Dim objConfigVV As ClassConfigValoresValidos
Dim sValor As String

On Error GoTo Erro_Move_Configs_Memoria
        
    For Each objConfig In gcolConfigs
            
        iIndice = iIndice + 1
        
        If UCase(Trim(objConfig.sTipoControle)) = UCase("ComboBox") Then
            For Each objConfigVV In objConfig.colValores
                If Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Valor_Col)) = objConfigVV.iSeq Then
                    sValor = objConfigVV.sValor
                    Exit For
                End If
            Next
        Else
            sValor = GridItens.TextMatrix(iIndice, iGrid_Valor_Col)
        End If
        
        objConfig.sConteudoNovo = sValor

    Next

    Move_Configs_Memoria = SUCESSO

    Exit Function

Erro_Move_Configs_Memoria:

    Move_Configs_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213387)

    End Select

    Exit Function

End Function


Public Sub Rotina_Grid_Enable(ByVal iLinha As Integer, ByVal objControl As Object, ByVal iLocalChamada As Integer)

Dim lErro As Long
Dim iCodValor As Integer, iIndiceAux As Integer
Dim iIndice As Integer
Dim objConfig As ClassConfigOutros
Dim objConfigVV As ClassConfigValoresValidos
Dim sValor As String

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name
        
        Case Valor.Name
            
            objControl.Enabled = True
        
            'Guardo o valor da Unidade de Medida da Linha
            iCodValor = Codigo_Extrai(Valor.Text)
            sValor = Valor.Text
            
            Valor.Clear
            
            If iLinha > 0 And iLinha <= gcolConfigs.Count Then
            
                Set objConfig = gcolConfigs.Item(iLinha)
                
                If UCase(Trim(objConfig.sTipoControle)) = UCase("ComboBox") Then
                    iIndice = -1
                    iIndiceAux = -1
                    For Each objConfigVV In objConfig.colValores
                        iIndiceAux = iIndiceAux + 1
                        If iCodValor = objConfigVV.iSeq Then
                            iIndice = iIndiceAux
                        End If
                        Valor.AddItem CStr(objConfigVV.iSeq) & SEPARADOR & objConfigVV.sDescricao
                    Next
                    Valor.ListIndex = iIndice
                Else
                    Valor.Text = sValor
                End If
            
            End If
            
        Case Else
            
            objControl.Enabled = False
        
    End Select
    
    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213388)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Modulo_Change()
    Call Trata_Modulo
End Sub

Private Sub Modulo_Click()
    Call Trata_Modulo
End Sub

Private Sub Trata_Modulo()

Dim lErro As Long
Dim objModulo As AdmModulo

On Error GoTo Erro_Trata_Modulo

    If sModuloAnt <> Modulo.Text Then
    
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        For Each objModulo In gcolModulos
            If objModulo.sNome = Modulo.Text Then
                Exit For
            End If
        Next

        lErro = Traz_Configs_Tela(objModulo.sSigla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Exit Sub
    
Erro_Trata_Modulo:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213389)
            
    End Select
    
    Exit Sub
End Sub

Private Sub BotaoProcuraFile_Click()

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then

        ' Set CancelError is True
        CD.CancelError = True
        
        On Error GoTo Erro_BotaoProcuraFile_Click
        ' Set flags
        CD.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
        ' Set filters
        CD.Filter = "All Files (*.*)|*.*"
        ' Specify default filter
        CD.FilterIndex = 2
        ' Display the Open dialog box
        CD.ShowOpen
        ' Display name of selected file
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = CD.FileName
        
    End If
    
    Exit Sub

Erro_BotaoProcuraFile_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub BotaoProcuraDir_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
Dim iPos As Integer, sDir As String

On Error GoTo Erro_BotaoProcuraDir_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        szTitle = "Localização do diretório"
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
            
            sDir = sBuffer
            
            If right(sDir, 1) <> "\" And right(sDir, 1) <> "/" Then
                iPos = InStr(1, sDir, "/")
                If iPos = 0 Then
                    sDir = sDir & "\"
                Else
                    sDir = sDir & "/"
                End If
            End If
           
            GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = sDir
      
        End If
        
    End If
  
    Exit Sub

Erro_BotaoProcuraDir_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213390)

    End Select

    Exit Sub
  
End Sub

Private Sub Trata_Controles()

Const POS_ORIGINAL_LEFT = -100
Const POS_ORIGINAL_TOP = -100

Const POS_LEFT = 8330

Dim objConfig As ClassConfigOutros

On Error GoTo Erro_Trata_Controles

    GridItens.ToolTipText = ""
    If GridItens.Row > 0 And GridItens.Col > 0 Then
        GridItens.ToolTipText = GridItens.TextMatrix(GridItens.Row, GridItens.Col)
    End If
    
    BotaoProcuraDir.top = POS_ORIGINAL_TOP
    BotaoProcuraDir.left = POS_ORIGINAL_LEFT
    BotaoProcuraDir.Visible = False
    
    BotaoProcuraFile.top = POS_ORIGINAL_TOP
    BotaoProcuraFile.left = POS_ORIGINAL_LEFT
    BotaoProcuraFile.Visible = False
    
    If iLinhaAnterior <> GridItens.Row Then
    
        If GridItens.Row > 0 And GridItens.Row <= gcolConfigs.Count Then
        
            Set objConfig = gcolConfigs.Item(GridItens.Row)
            
            If UCase(Trim(objConfig.sTipoControle)) = UCase("Dir") Then
                BotaoProcuraDir.top = GridItens.CellTop + GridItens.CellHeight
                BotaoProcuraDir.left = POS_LEFT
                BotaoProcuraDir.Visible = True
            End If
            
            If UCase(Trim(objConfig.sTipoControle)) = UCase("File") Then
                BotaoProcuraFile.top = POS_ORIGINAL_TOP
                BotaoProcuraFile.left = POS_LEFT
                BotaoProcuraFile.Visible = True
            End If
        
        End If
        
        iLinhaAnterior = GridItens.Row
    
    End If
  
    Exit Sub

Erro_Trata_Controles:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213391)

    End Select

    Exit Sub
End Sub

