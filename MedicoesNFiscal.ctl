VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl MedicoesNFiscalOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8700
   Begin VB.CommandButton BotaoMedicao 
      Caption         =   "Medições"
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
      TabIndex        =   18
      Top             =   4845
      Width           =   1845
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   4230
      Picture         =   "MedicoesNFiscal.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1005
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2490
      Picture         =   "MedicoesNFiscal.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5265
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produto "
      Height          =   1065
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   8550
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
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
         Left            =   7125
         TabIndex        =   13
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Item 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7680
         TabIndex        =   12
         Top             =   585
         Width           =   525
      End
      Begin VB.Label DescContrato 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2220
         TabIndex        =   11
         Top             =   225
         Width           =   5985
      End
      Begin VB.Label Contrato 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1050
         TabIndex        =   8
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato:"
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
         Index           =   8
         Left            =   165
         TabIndex        =   7
         Top             =   270
         Width           =   795
      End
      Begin VB.Label DescricaoProduto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2235
         TabIndex        =   5
         Top             =   615
         Width           =   4695
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   225
         TabIndex        =   4
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1035
         TabIndex        =   3
         Top             =   615
         Width           =   1185
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Medições"
      Height          =   3555
      Left            =   60
      TabIndex        =   6
      Top             =   1215
      Width           =   8535
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4905
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   930
         Width           =   855
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Left            =   4740
         TabIndex        =   15
         Top             =   2085
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   270
         Left            =   4350
         TabIndex        =   14
         Top             =   1710
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Medicao 
         Height          =   270
         Left            =   2955
         TabIndex        =   0
         Top             =   1245
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2460
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   4339
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label QuantidadeTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3465
         TabIndex        =   21
         Top             =   3135
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Total:"
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
         Left            =   1830
         TabIndex        =   20
         Top             =   3210
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total:"
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
         Left            =   5460
         TabIndex        =   17
         Top             =   3210
         Width           =   1005
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6540
         TabIndex        =   16
         Top             =   3135
         Width           =   1740
      End
   End
End
Attribute VB_Name = "MedicoesNFiscalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolMedicoes As Collection
Dim gobjItemContrato As ClassItensDeContrato
Dim giTipoContrato As Integer

Public objGridItens As AdmGrid

Private WithEvents objEventoMedicao As AdmEvento
Attribute objEventoMedicao.VB_VarHelpID = -1

Dim iGrid_Valor_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Medicao_Col As Integer
Dim iGrid_UM_Col As Integer

Public iAlterado As Integer

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "NFiscal - Medições"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MedicoesNFiscal"

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
        If Me.ActiveControl Is Medicao Then
            Call BotaoMedicao_Click
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

    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    Set objGridItens = New AdmGrid
    
    'Seta as Variáveis das Telas de browse
    Set objEventoMedicao = New AdmEvento
   
    Call Inicializa_Grid_Itens(objGridItens)
   
    'Sinaliza que o Form_Loas ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Function Trata_Parametros(colMedicoes As Collection, objItemContrato As ClassItensDeContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Faz a variável global a tela apontar para a variável passada
    Set gcolMedicoes = colMedicoes
    Set gobjItemContrato = objItemContrato
        
    lErro = Traz_Medicoes_Tela(colMedicoes, objItemContrato)
    If lErro <> SUCESSO Then gError 136202
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 136202
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162696)
    
    End Select
    
    Exit Function
        
End Function

Function Saida_Celula(objGridItens As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridItens)

    If lErro = SUCESSO Then

        lErro = Saida_Celula_Medicao(objGridItens)
        If lErro <> SUCESSO Then gError 123221

        lErro = Grid_Finaliza_Saida_Celula(objGridItens)
        If lErro <> SUCESSO Then gError 123222

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 123221, 123222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162697)

    End Select

    Exit Function

End Function

Private Sub BotaoCancela_Click()
    
    'Nao mexer no obj da tela
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 126559
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 126559
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162698)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_Medicoes_Memoria()
    If lErro <> SUCESSO Then gError 136212
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 136212
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162699)

    End Select

    Exit Function

End Function

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Valor()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Medicao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Medicao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Medicao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Medicao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Medicao()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    
    Set objEventoMedicao = Nothing

End Sub

Private Function Saida_Celula_Medicao(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objItemMedicao As New ClassItensMedCtr
Dim bAchou As Boolean
Dim objNF As New ClassNFiscal
Dim objItemNF As New ClassItemNF
Dim iItem As Integer
Dim sContrato As String
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Medicao

    Set objGridInt.objControle = Medicao

    'Se quantidade estiver preenchida
    If Len(Trim(Medicao.Text)) > 0 Then
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If iIndice <> GridItens.Row And StrParaLong(GridItens.TextMatrix(iIndice, iGrid_Medicao_Col)) = StrParaLong(Medicao.Text) Then gError 136219
        Next
        
        iItem = StrParaInt(Item.Caption)
        sContrato = Contrato.Caption
        
        If Len(Trim(sContrato)) = 0 Then gError 136156
        If iItem = 0 Then gError 136157
    
         'Critica o valor
        lErro = Valor_Long_Critica(Medicao.Text)
        If lErro <> SUCESSO Then gError 132963
   
        objItemMedicao.lMedicao = StrParaLong(Medicao.Text)
        objItemMedicao.iItem = iItem
        
        'Le a medição
        lErro = CF("ItensDeMedicaoContrato_Le2", objItemMedicao)
        If lErro <> SUCESSO And lErro <> 136173 Then gError 132960
        
        If lErro = 136173 Then gError 132964
        
        If objItemMedicao.sContrato <> sContrato Then gError 136182
        
        lErro = CF("ItensDeContrato_Le_DadosFatura", objNF, objItemNF)
        If lErro <> SUCESSO And lErro <> 129904 And lErro <> 129907 And lErro <> 129908 Then gError 129950
        If lErro = SUCESSO Then gError 132965 'Se essa medição já foi faturada => Erro

        GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = Format(objItemMedicao.dVlrCobrar, "STANDARD")
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(objItemMedicao.dQuantidade)
        GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = objItemMedicao.objItensDeContrato.sUM

        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    Call Calcula_Totais

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132961

    Saida_Celula_Medicao = SUCESSO

    Exit Function

Erro_Saida_Celula_Medicao:

    Saida_Celula_Medicao = gErr

    Select Case gErr

        Case 132960, 132961, 132963
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
         Case 132964
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMMEDICAO_NAO_CADASTRADO", gErr, objItemMedicao.lMedicao, sContrato, objItemMedicao.iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 132965
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMMEDICAO_FATURADO", gErr, objItemMedicao.lMedicao, objItemMedicao.sContrato, objItemMedicao.iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 136156
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 136157
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CODIGO_CONTRATO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 136182
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_CONTRATO_DIF", gErr, objItemMedicao.sContrato, sContrato)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 136219
            Call Rotina_Erro(vbOKOnly, "ERRO_MEDICAO_REPETIDA", gErr, StrParaLong(Medicao.Text), iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162700)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub objEventoMedicao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objItemMedicao As New ClassItensMedCtr
Dim bCancel As Boolean
Dim iIndice As Integer

On Error GoTo Erro_objEventoMedicao_evSelecao

    Set objItemMedicao = obj1

    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row And StrParaLong(GridItens.TextMatrix(iIndice, iGrid_Medicao_Col)) = objItemMedicao.lMedicao Then gError 136218
    Next

    GridItens.TextMatrix(GridItens.Row, iGrid_Medicao_Col) = objItemMedicao.lMedicao

    lErro = CF("ItensDeMedicaoContrato_Le2", objItemMedicao)
    If lErro <> SUCESSO And lErro <> 136173 Then gError 136165

    If lErro = 136173 Then gError 136179

    GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col) = Format(objItemMedicao.dVlrCobrar, "STANDARD")
    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(objItemMedicao.dQuantidade)
    GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = objItemMedicao.objItensDeContrato.sUM

    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    End If
    
    Call Calcula_Totais

    Me.Show
    
    Exit Sub
    
Erro_objEventoMedicao_evSelecao:

    Select Case gErr
    
        Case 136165
        
        Case 136179
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_NAO_CADASTRADO", gErr, objItemMedicao.lMedicao, objItemMedicao.iItem)

        Case 136218
            Call Rotina_Erro(vbOKOnly, "ERRO_MEDICAO_REPETIDA", gErr, objItemMedicao.lMedicao, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162701)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoMedicao_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objItemMedicao As New ClassItensMedCtr
Dim sContrato As String
Dim iSeq As Integer
Dim lMedicao As Long

On Error GoTo Erro_BotaoMedicao_Click

    If GridItens.Row = 0 Then gError 132937

    sContrato = Contrato.Caption
    lMedicao = StrParaLong(GridItens.TextMatrix(GridItens.Row, iGrid_Medicao_Col))
    iSeq = StrParaInt(Item.Caption)

    If Len(Trim(sContrato)) = 0 Then gError 132938
    If iSeq = 0 Then gError 132939
    
    colSelecao.Add sContrato
    colSelecao.Add iSeq
    
    objItemMedicao.lMedicao = lMedicao
    
    If giTipoContrato = CONTRATOS_RECEBER Then
        Call Chama_Tela_Modal("MedicaoCliItensAFaturarLista", colSelecao, objItemMedicao, objEventoMedicao)
    Else
        Call Chama_Tela_Modal("MedicaoFornItensAPagarLista", colSelecao, objItemMedicao, objEventoMedicao)
    End If

    Exit Sub
    
Erro_BotaoMedicao_Click:
    
    Select Case gErr

        Case 132937
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 132938
             Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case 132939
             Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162702)

    End Select
    
    Exit Sub

End Sub

Function Traz_Medicoes_Tela(colMedicoes As Collection, objItemContrato As ClassItensDeContrato) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objContrato As New ClassContrato
Dim sProduto As String
Dim objItemMedicao As ClassItensMedCtr
Dim iIndice As Integer

On Error GoTo Erro_Traz_Medicoes_Tela

    objProduto.sCodigo = objItemContrato.sProduto
    
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 35943
    
    Produto.Caption = sProduto
    DescricaoProduto.Caption = objItemContrato.sDescProd
    
    'Le o Contrato
    objContrato.lNumIntDoc = objItemContrato.lNumIntContrato

    lErro = CF("Contrato_Le2", objContrato)
    If lErro <> 129261 And lErro <> SUCESSO Then gError 136203
    
    Contrato.Caption = objContrato.sCodigo
    DescContrato.Caption = objContrato.sDescricao
    
    giTipoContrato = objContrato.iTipo
    
    Item.Caption = objItemContrato.iSeq
      
    If colMedicoes.Count <> 0 Then
    
        For Each objItemMedicao In colMedicoes
        
            iIndice = iIndice + 1
             
            GridItens.TextMatrix(iIndice, iGrid_Medicao_Col) = objItemMedicao.lMedicao
            GridItens.TextMatrix(iIndice, iGrid_UM_Col) = objItemMedicao.objItensDeContrato.sUM
            GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemMedicao.dQuantidade)
            GridItens.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objItemMedicao.dVlrCobrar, gobjFAT.sFormatoPrecoUnitario)
            
        Next
    
    End If
    
    objGridItens.iLinhasExistentes = iIndice
    
    Call Calcula_Totais
           
    Traz_Medicoes_Tela = SUCESSO

    Exit Function

Erro_Traz_Medicoes_Tela:

    Traz_Medicoes_Tela = gErr
    
    Select Case gErr
    
        Case 136203
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162703)
    
    End Select
    
    Exit Function
    
End Function

Function Move_Medicoes_Memoria() As Long

Dim lErro As Long
Dim objItemMedicao As ClassItensMedCtr
Dim iIndice As Integer

On Error GoTo Erro_Move_Medicoes_Memoria

    For iIndice = gcolMedicoes.Count To 1 Step -1
        gcolMedicoes.Remove iIndice
    Next

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objItemMedicao = New ClassItensMedCtr
    
        objItemMedicao.lMedicao = StrParaLong(GridItens.TextMatrix(iIndice, iGrid_Medicao_Col))
        objItemMedicao.iItem = StrParaInt(Item.Caption)
        
        'Le a medição
        lErro = CF("ItensDeMedicaoContrato_Le2", objItemMedicao)
        If lErro <> SUCESSO And lErro <> 136173 Then gError 136204
        
        If lErro = 136173 Then gError 136205
        
        gcolMedicoes.Add objItemMedicao
        
    Next
      
    Move_Medicoes_Memoria = SUCESSO

    Exit Function

Erro_Move_Medicoes_Memoria:

    Move_Medicoes_Memoria = gErr
    
    Select Case gErr
    
        Case 136204
        
        Case 136205
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMMEDICAO_NAO_CADASTRADO", gErr, objItemMedicao.lMedicao, Contrato.Caption, objItemMedicao.iItem)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162704)
    
    End Select
    
    Exit Function
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Medicao")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Valor Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Medicao.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Valor.Name)

    iGrid_Medicao_Col = 1
    iGrid_UM_Col = 2
    iGrid_Quantidade_Col = 3
    iGrid_Valor_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function


Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
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

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Private Sub Calcula_Totais()

Dim iIndice As Integer
Dim dQuantidade As Double
Dim dValor As Double

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        dQuantidade = dQuantidade + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        dValor = dValor + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Valor_Col))
    
    Next
    
    ValorTotal.Caption = Format(dValor, "STANDARD")
    QuantidadeTotal.Caption = Formata_Estoque(dQuantidade)

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case Medicao.Name
            
            'Se o produto estiver preenchido, habilita o controle
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Medicao_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162705)

    End Select

    Exit Sub

End Sub
