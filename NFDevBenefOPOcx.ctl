VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl NFDevBenefOPOcx 
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   5310
   Begin VB.Frame Frame1 
      Caption         =   "Calcular a devolução com base nas OPs"
      Height          =   2820
      Left            =   75
      TabIndex        =   4
      Top             =   750
      Width           =   5130
      Begin VB.TextBox OP 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1260
         MaxLength       =   250
         TabIndex        =   5
         Top             =   735
         Width           =   3420
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2385
         Left            =   75
         TabIndex        =   6
         Top             =   210
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   4207
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.CommandButton BotaoOP 
      Caption         =   "Ordens de Produção para Beneficiamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   3615
      Width           =   5145
   End
   Begin VB.CommandButton BotaoCalc 
      Caption         =   "Calcular quantidade a devolver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   90
      TabIndex        =   2
      Top             =   4350
      Width           =   2295
   End
   Begin VB.CommandButton BotaoCancelar 
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
      Height          =   510
      Left            =   2910
      TabIndex        =   1
      Top             =   4365
      Width           =   2295
   End
   Begin VB.CommandButton BotaoSugerir 
      Caption         =   "Sugestão automática das OPs a serem consideradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   3990
      Width           =   5145
   End
   Begin VB.Label Cliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1155
      TabIndex        =   11
      Top             =   60
      Width           =   4050
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   420
      TabIndex        =   12
      Top             =   105
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "Filial:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   615
      TabIndex        =   10
      Top             =   450
      Width           =   465
   End
   Begin VB.Label FilialCliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1155
      TabIndex        =   9
      Top             =   420
      Width           =   1515
   End
   Begin VB.Label DataUltDev 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4020
      TabIndex        =   8
      Top             =   435
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Data Últ. Dev.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2700
      TabIndex        =   7
      Top             =   480
      Width           =   1320
   End
End
Attribute VB_Name = "NFDevBenefOPOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjNFDevBenef As ClassNFDevBenef

Dim objGridItens As AdmGrid
Dim iGrid_OP_Col As Integer

Private WithEvents objEventoOP As AdmEvento
Attribute objEventoOP.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Ordens de Produção - Dev. Simb. por Benef."
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFDevBenefOPOcx"
    
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
'**** fim do trecho a ser copiado *****

Public Function Trata_Parametros(objNFDevBenef As ClassNFDevBenef) As Long

Dim lErro As Long
Dim objCli As New ClassCliente
Dim objFilCli As New ClassFilialCliente

On Error GoTo Erro_Trata_Parametros

    Set gobjNFDevBenef = objNFDevBenef
    Set objEventoOP = New AdmEvento
    Set objGridItens = New AdmGrid
    
    gobjNFDevBenef.iRetorno = vbCancel
    
    'inicializacao do grid
    Call Inicializa_Grid_Itens(objGridItens)
    
    objNFDevBenef.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("NFDevBenef_Obter_UltimaDev", objNFDevBenef)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If gobjFAT.iNFDevSimbSugOP = MARCADO And objNFDevBenef.colOP.Count = 0 Then
    
        Call BotaoSugerir_Click
        
    ElseIf objNFDevBenef.colOP.Count <> 0 Then
        
        lErro = Traz_OPs_Tela(objNFDevBenef)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objNFDevBenef.dtDataUltDev <> DATA_NULA Then
        DataUltDev.Caption = Format(objNFDevBenef.dtDataUltDev, "dd/mm/yyyy")
    End If
    
    objCli.lCodigo = objNFDevBenef.lCliente

    lErro = CF("Cliente_Le", objCli)
    If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
    
    Cliente.Caption = CStr(objCli.lCodigo) & SEPARADOR & objCli.sNomeReduzido

    objFilCli.lCodCliente = objNFDevBenef.lCliente
    objFilCli.iCodFilial = objNFDevBenef.iFilialCliente
    
    lErro = CF("FilialCliente_Le", objFilCli)
    If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM

    FilialCliente.Caption = CStr(objFilCli.iCodFilial) & SEPARADOR & objFilCli.sNome
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209968)

    End Select

    Exit Function
    
End Function

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209969)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Itens

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Ordem de Produção")

    'campos de edição do grid
    objGridInt.colCampo.Add (OP.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_OP_Col = 1

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 750

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Itens:

    Inicializa_Grid_Itens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209970)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set objEventoOP = Nothing
    Set gobjNFDevBenef = Nothing
    Set objGridItens = Nothing
End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridItens.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_OP_Col

                    lErro = Saida_Celula_OP(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 209971

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'erros tratatos nas rotinas chamadas

        Case 209971
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209972)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case Else
            objControl.Enabled = True

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209973)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_OP(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemOP As New ClassItemOP
Dim iProdutoOPPreenchido As Integer
Dim sProdutoOPFormatado As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_Saida_Celula_OP

    Set objGridInt.objControle = OP

    If Len(Trim(OP.Text)) > 0 Then

        objOrdemProducao.iFilialEmpresa = giFilialEmpresa
        objOrdemProducao.sCodigo = OP.Text

        lErro = CF("OrdemProducao_Le1", objOrdemProducao, True)
        If lErro <> SUCESSO And lErro <> 94578 And lErro <> 94579 Then gError ERRO_SEM_MENSAGEM

        'Se a OP não estiver cadastrada
        If lErro = 94579 Then gError 209974
        
        'Se não for uma OP desse cliente\filial
        If objOrdemProducao.lCodTerc <> gobjNFDevBenef.lCliente Or objOrdemProducao.iFilialTerc <> gobjNFDevBenef.iFilialCliente Then gError 209975

        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_OP = SUCESSO

    Exit Function

Erro_Saida_Celula_OP:

    Saida_Celula_OP = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 209974
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_CADASTRADA", gErr, objOrdemProducao.sCodigo)

        Case 209975
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_TERC_DIF", gErr)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209976)

    End Select

    Exit Function

End Function

Private Sub OP_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OP_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub OP_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub OP_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = OP
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoSugerir_Click()

Dim lErro As Long
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoSugerir_Click

    If gobjNFDevBenef.colOP.Count <> 0 Then
    
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_LIMPAR_GRID1")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    lErro = CF("NFDevBenef_Sugerir_OPs", gobjNFDevBenef)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_OPs_Tela(gobjNFDevBenef)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_BotaoSugerir_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209977)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCalc_Click()

Dim lErro As Long
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoCalc_Click

    If gobjNFDevBenef.colItens.Count <> 0 Then
    
        'Avisa que os itens já foram calculados, se seguir vai alterar os itens na NF
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_NFBENEF_ITENS_JA_CALC")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
           
    End If
    
    lErro = Move_OPs_Memoria(gobjNFDevBenef)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("NFDevBenef_Calcula", gobjNFDevBenef)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    gobjNFDevBenef.iRetorno = vbOK

    Unload Me

    Exit Sub

Erro_BotaoCalc_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209978)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCancelar_Click()
    Unload Me
End Sub

Private Function Traz_OPs_Tela(ByVal objNFDevBenef As ClassNFDevBenef) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vValor As Variant

On Error GoTo Erro_Traz_OPs_Tela

    Call Grid_Limpa(objGridItens)
    
    iIndice = 0
    For Each vValor In objNFDevBenef.colOP
        iIndice = iIndice + 1
        GridItens.TextMatrix(iIndice, iGrid_OP_Col) = vValor
    Next
    objGridItens.iLinhasExistentes = iIndice

    Traz_OPs_Tela = SUCESSO

    Exit Function

Erro_Traz_OPs_Tela:

    Traz_OPs_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209979)

    End Select

    Exit Function

End Function

Private Function Move_OPs_Memoria(ByVal objNFDevBenef As ClassNFDevBenef) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_OPs_Memoria

    Set objNFDevBenef.colOP = New Collection
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        objNFDevBenef.colOP.Add GridItens.TextMatrix(iIndice, iGrid_OP_Col)
    Next

    Move_OPs_Memoria = SUCESSO

    Exit Function

Erro_Move_OPs_Memoria:

    Move_OPs_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209980)

    End Select

    Exit Function

End Function

Private Sub objEventoOP_evSelecao(obj1 As Object)

Dim objOrdemProducao As New ClassOrdemDeProducao

    Set objOrdemProducao = obj1

    OP.Text = objOrdemProducao.sCodigo
    
    If Not (Me.ActiveControl Is OP) Then
        'Coloca o produto, a Descrição e a Unidade de Medida da tela
        GridItens.TextMatrix(GridItens.Row, iGrid_OP_Col) = OP.Text
    End If

    Me.Show

End Sub

Private Sub BotaoOP_Click()

Dim objOrdemProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection
Dim sOP As String, sFiltro As String

On Error GoTo Erro_BotaoOP_Click

    If Me.ActiveControl Is OP Then
        sOP = OP.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 209981
        sOP = GridItens.TextMatrix(GridItens.Row, iGrid_OP_Col)
    End If
        
    objOrdemProducao.sCodigo = sOP
    
    colSelecao.Add LCodigo_Extrai(Cliente.Caption)
    colSelecao.Add Codigo_Extrai(FilialCliente.Caption)
    
    sFiltro = "TipoTerc = 1 AND CodTerc = ? AND FilialTerc = ?"

    Call Chama_Tela_Modal("OrdProdBaixadasListaModal", colSelecao, objOrdemProducao, objEventoOP, sFiltro)

    Exit Sub
    
Erro_BotaoOP_Click:
    
    Select Case gErr
    
        Case 209981
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209982)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is OP Then
            Call BotaoOP_Click
        End If
    
    End If
End Sub
