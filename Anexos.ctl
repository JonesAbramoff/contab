VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl AnexosOcx 
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoAbrirArq 
      Caption         =   "Abrir arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1110
      TabIndex        =   5
      Top             =   3810
      Width           =   1020
   End
   Begin VB.CommandButton BotaoArquivo 
      Caption         =   "Procurar arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   45
      TabIndex        =   4
      Top             =   3810
      Width           =   1020
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3900
      Picture         =   "Anexos.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3810
      Width           =   885
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4890
      Picture         =   "Anexos.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3810
      Width           =   885
   End
   Begin VB.Frame FrameAnexos 
      Caption         =   "Anexos"
      Height          =   3735
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   9330
      Begin VB.TextBox ItemDescricao 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3045
         TabIndex        =   3
         Top             =   870
         Width           =   5265
      End
      Begin VB.TextBox ItemArquivo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1410
         TabIndex        =   2
         Top             =   600
         Width           =   5265
      End
      Begin MSFlexGridLib.MSFlexGrid GridAnexos 
         Height          =   2325
         Left            =   105
         TabIndex        =   1
         Top             =   285
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   4075
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin MSComDlg.CommonDialog CDProcurar 
      Left            =   2310
      Top             =   3855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "AnexosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer

Dim gobjAnexos As ClassAnexos

Dim objGridAnexos As AdmGrid

Dim iGrid_ItemArquivo_Col As Integer
Dim iGrid_ItemDescricao_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Anexos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Anexos"

End Function

Public Sub Show()
'    Parent.Show
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
        
        If Me.ActiveControl Is ItemArquivo Then
            Call BotaoArquivo_Click
        End If
    
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridAnexos = Nothing
    Set gobjAnexos = Nothing

'    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201516)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGridAnexos = New AdmGrid

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201517)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(ByVal objAnexos As ClassAnexos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Set gobjAnexos = objAnexos

    lErro = Inicializa_Anexos(objGridAnexos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Traz_Anexos_Tela(gobjAnexos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201518)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objAnexos As ClassAnexos) As Long

Dim lErro As Long, iLinha As Integer, objAnexoArq As ClassAnexosArq

On Error GoTo Erro_Move_Tela_Memoria

    Set objAnexos.colArq = New Collection
    
    For iLinha = 1 To objGridAnexos.iLinhasExistentes
    
        If Len(Trim(GridAnexos.TextMatrix(iLinha, iGrid_ItemArquivo_Col))) = 0 Then Exit For
        
        Set objAnexoArq = New ClassAnexosArq
        
        objAnexoArq.iSeq = iLinha
        objAnexoArq.sArquivo = GridAnexos.TextMatrix(iLinha, iGrid_ItemArquivo_Col)
        objAnexoArq.sDescricao = GridAnexos.TextMatrix(iLinha, iGrid_ItemDescricao_Col)
        
        objAnexos.colArq.Add objAnexoArq
    
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201519)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long
'
End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long
'
End Function

Function Traz_Anexos_Tela(objAnexos As ClassAnexos) As Long

Dim lErro As Long, objAnexoArq As ClassAnexosArq, iLinha As Integer

On Error GoTo Erro_Traz_Anexos_Tela

    iLinha = 0
    
    For Each objAnexoArq In objAnexos.colArq
    
        iLinha = iLinha + 1
            
        GridAnexos.TextMatrix(iLinha, iGrid_ItemArquivo_Col) = objAnexoArq.sArquivo
        GridAnexos.TextMatrix(iLinha, iGrid_ItemDescricao_Col) = objAnexoArq.sDescricao
    
    Next

    objGridAnexos.iLinhasExistentes = iLinha
    
    iAlterado = 0

    Traz_Anexos_Tela = SUCESSO

    Exit Function

Erro_Traz_Anexos_Tela:

    Traz_Anexos_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201530)

    End Select

    Exit Function

End Function

Private Function Inicializa_Anexos(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add (" ")
    objGrid.colColuna.Add ("Arquivo")
    objGrid.colColuna.Add ("Descrição")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ItemArquivo.Name)
    objGrid.colCampo.Add (ItemDescricao.Name)

    'Colunas do Grid
    iGrid_ItemArquivo_Col = 1
    iGrid_ItemDescricao_Col = 2

    objGrid.objGrid = GridAnexos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridAnexos.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    Inicializa_Anexos = SUCESSO

End Function

Private Sub GridAnexos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAnexos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAnexos, iAlterado)
    End If

End Sub

Private Sub GridAnexos_GotFocus()
    Call Grid_Recebe_Foco(objGridAnexos)
End Sub

Private Sub GridAnexos_EnterCell()
    Call Grid_Entrada_Celula(objGridAnexos, iAlterado)
End Sub

Private Sub GridAnexos_LeaveCell()
    Call Saida_Celula(objGridAnexos)
End Sub

Private Sub GridAnexos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAnexos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAnexos, iAlterado)
    End If

End Sub

Private Sub GridAnexos_RowColChange()
    Call Grid_RowColChange(objGridAnexos)
End Sub

Private Sub GridAnexos_Scroll()
    Call Grid_Scroll(objGridAnexos)
End Sub

Private Sub GridAnexos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridAnexos)
End Sub

Private Sub GridAnexos_LostFocus()
    Call Grid_Libera_Foco(objGridAnexos)
End Sub

Private Sub ItemArquivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemArquivo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnexos)
End Sub

Private Sub ItemArquivo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnexos)
End Sub

Private Sub ItemArquivo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAnexos.objControle = ItemArquivo
    lErro = Grid_Campo_Libera_Foco(objGridAnexos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ItemDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAnexos)
End Sub

Private Sub ItemDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAnexos)
End Sub

Private Sub ItemDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAnexos.objControle = ItemDescricao
    lErro = Grid_Campo_Libera_Foco(objGridAnexos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ItemArquivo(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemArquivo

    Set objGridInt.objControle = ItemArquivo

    If Len(Trim(ItemArquivo.Text)) <> 0 Then
    
        If (GridAnexos.Row - GridAnexos.FixedRows) = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ItemArquivo = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemArquivo:

    Saida_Celula_ItemArquivo = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201520)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ItemDescricao(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemDescricao

    Set objGridInt.objControle = ItemDescricao


    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ItemDescricao = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemDescricao:

    Saida_Celula_ItemDescricao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201521)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col

            Case iGrid_ItemArquivo_Col

                lErro = Saida_Celula_ItemArquivo(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case iGrid_ItemDescricao_Col

                lErro = Saida_Celula_ItemDescricao(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End Select
                    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 201522

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 201522
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201523)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name

        Case ItemArquivo.Name
            objControl.Enabled = True
        
        Case ItemDescricao.Name
            If Len(Trim(GridAnexos.TextMatrix(iLinha, iGrid_ItemArquivo_Col))) <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201524)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click
  
    lErro = Move_Tela_Memoria(gobjAnexos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
      
    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207597)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoArquivo_Click()

On Error GoTo Erro_BotaoArquivo_Click

    If GridAnexos.Row <> 0 Then

        ' Set CancelError is True
        CDProcurar.CancelError = True
                
        ' Set flags
        CDProcurar.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
        ' Set filters
        CDProcurar.Filter = "All Files (*.*)|*.*|"
        
        ' Specify default filter
        CDProcurar.FilterIndex = 2
        ' Display the Open dialog box
        CDProcurar.ShowOpen
        ' Display name of selected file
    
        If ActiveControl.Name = ItemArquivo.Name Then
            ItemArquivo.Text = CDProcurar.FileName
        Else
            GridAnexos.TextMatrix(GridAnexos.Row, iGrid_ItemArquivo_Col) = CDProcurar.FileName
            
            If (GridAnexos.Row - GridAnexos.FixedRows) = objGridAnexos.iLinhasExistentes Then
                objGridAnexos.iLinhasExistentes = objGridAnexos.iLinhasExistentes + 1
            End If
        End If
    Else
        Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    End If
    
    Exit Sub

Erro_BotaoArquivo_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub BotaoAbrirArq_Click()

Dim sArq As String

On Error GoTo Erro_BotaoAbrirArq_Click

    If GridAnexos.Row = 0 Then gError 207597
        sArq = GridAnexos.TextMatrix(GridAnexos.Row, iGrid_ItemArquivo_Col)
        If Len(Trim(sArq)) > 0 Then Call ShellExecute(hWnd, "open", sArq, vbNullString, vbNullString, 1)
    Exit Sub

    Exit Sub

Erro_BotaoAbrirArq_Click:

    Select Case gErr

        Case 207597
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207597)

    End Select

    Exit Sub
    
End Sub
