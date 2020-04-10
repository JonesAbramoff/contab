VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl RastreamentoSerie 
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   9000
   Begin VB.CommandButton BotaoSerie 
      Caption         =   "Séries"
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
      Left            =   7020
      TabIndex        =   13
      Top             =   3930
      Width           =   1845
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   4725
      Picture         =   "RastreamentoSerie.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4260
      Width           =   1005
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2970
      Picture         =   "RastreamentoSerie.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4245
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produto "
      Height          =   3690
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5190
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtde Séries:"
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
         Left            =   210
         TabIndex        =   23
         Top             =   3150
         Width           =   1065
      End
      Begin VB.Label QuantidadeTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1350
         TabIndex        =   22
         Top             =   3120
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movto:"
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
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   2664
         Width           =   1035
      End
      Begin VB.Label TipoMovto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1365
         TabIndex        =   20
         Top             =   2640
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Escaninho:"
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
         Left            =   315
         TabIndex        =   19
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Escaninho 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1350
         TabIndex        =   18
         Top             =   2160
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UM:"
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
         Index           =   4
         Left            =   915
         TabIndex        =   17
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label QtdAlm 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3255
         TabIndex        =   16
         Top             =   1695
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtd. Almox:"
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
         Left            =   2205
         TabIndex        =   15
         Top             =   1755
         Width           =   990
      End
      Begin VB.Label UM 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1350
         TabIndex        =   14
         Top             =   1695
         Width           =   645
      End
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
         Left            =   840
         TabIndex        =   12
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Item 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1350
         TabIndex        =   11
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Almoxarifado 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1350
         TabIndex        =   8
         Top             =   1215
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label DescricaoProduto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2820
         TabIndex        =   5
         Top             =   330
         Width           =   2250
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
         Left            =   540
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1350
         TabIndex        =   3
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Rastreamento por Número de Série"
      Height          =   3705
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   3375
      Begin MSMask.MaskEdBox Serie 
         Height          =   270
         Left            =   360
         TabIndex        =   0
         Top             =   1800
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3240
         Left            =   120
         TabIndex        =   1
         Top             =   315
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   5715
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
End
Attribute VB_Name = "RastreamentoSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolRastreamentoMovto As Collection
Dim gobjItemMovEstoque As ClassItemMovEstoque

Dim gsTelaOrigem As String
Dim gbPodeAlterarQtd As Boolean
Dim bCarregando As Boolean

Public objGridItens As AdmGrid

Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1

Dim iGrid_Serie_Col As Integer

Public iAlterado As Integer

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object
    Set Form_Load_Ocx = Me
    Caption = "Rastreamento por Número de Série"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "RastreamentoSerie"
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
        If Me.ActiveControl Is Serie Then
            Call BotaoSerie_Click
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
    bCarregando = True
    
    giRetornoTela = vbAbort
    
    Set objGridItens = New AdmGrid
    
    'Seta as Variáveis das Telas de browse
    Set objEventoSerie = New AdmEvento
   
    Call Inicializa_Grid_Itens(objGridItens)
   
    bCarregando = False
   
    'Sinaliza que o Form_Loas ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Function Trata_Parametros(ByVal colRastreamentoMovto As Collection, ByVal objItemMovEstoque As ClassItemMovEstoque, Optional sTelaOrigem As String = "", Optional bPodeAlterarQtd As Boolean = True) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    bCarregando = True

    'Faz a variável global a tela apontar para a variável passada
    Set gcolRastreamentoMovto = colRastreamentoMovto
    Set gobjItemMovEstoque = objItemMovEstoque
    
    gsTelaOrigem = sTelaOrigem
    gbPodeAlterarQtd = bPodeAlterarQtd
        
    lErro = Traz_RastreamentoSerie_Tela(colRastreamentoMovto, objItemMovEstoque)
    If lErro <> SUCESSO Then gError 141815
    
    bCarregando = False
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 141815
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141832)
    
    End Select
    
    Exit Function
        
End Function

Function Saida_Celula(objGridItens As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridItens)
    If lErro = SUCESSO Then

        If objGridItens.objGrid.Name = GridItens.Name Then

           Select Case GridItens.Col
    
                Case iGrid_Serie_Col

                    lErro = Saida_Celula_Serie(objGridItens)
                    If lErro <> SUCESSO Then gError 141816
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridItens)
        If lErro <> SUCESSO Then gError 141818

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 141816, 141818

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141831)

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
    If lErro <> SUCESSO Then gError 141819
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 141819
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141830)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_RastreamentoSerie_Memoria()
    If lErro <> SUCESSO Then gError 141820
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 141820
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141829)

    End Select

    Exit Function

End Function

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Serie_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Serie()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objGridItens = Nothing
    
    Set objEventoSerie = Nothing

End Sub

Private Function Saida_Celula_Serie(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Serie

    Set objGridInt.objControle = Serie

    lErro = Lote_Saida_Celula(Serie.Text, GridItens.Row)
    If lErro <> SUCESSO Then gError 141821

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141822

    Saida_Celula_Serie = SUCESSO

    Exit Function

Erro_Saida_Celula_Serie:

    Saida_Celula_Serie = gErr

    Select Case gErr
    
        Case 141821, 141822
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141828)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroSaldo As ClassRastroLoteSaldo

On Error GoTo Erro_objEventoSerie_evSelecao

    Set objRastroSaldo = obj1
   
    Serie.Text = objRastroSaldo.sLote
   
    If Not (Me.ActiveControl Is Serie) Then
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Serie_Col) = objRastroSaldo.sLote
        
        lErro = Lote_Saida_Celula(objRastroSaldo.sLote, GridItens.Row)
        If lErro <> SUCESSO Then gError 141827
    
    End If

    Me.Show
    
    Exit Sub
    
Erro_objEventoSerie_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Serie_Col) = ""

    Select Case gErr
    
        Case 141827
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141832)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoSerie_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoSerie_Click

    If GridItens.Row = 0 Then gError 141833

    lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141834

    colSelecao.Add sProdutoFormatado
    colSelecao.Add gobjItemMovEstoque.iAlmoxarifado

    Call Chama_Tela_Modal("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoSerie, "Produto = ? AND Almoxarifado = ?")

    Exit Sub
    
Erro_BotaoSerie_Click:
    
    Select Case gErr

        Case 141833
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 141834
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141835)

    End Select
    
    Exit Sub

End Sub

Function Traz_RastreamentoSerie_Tela(ByVal colRastreamentoMovto As Collection, ByVal objItemMovEstoque As ClassItemMovEstoque) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoLote As ClassRastreamentoLote
Dim objProduto As ClassProduto
Dim sProdutoMascarado As String
Dim objTipoMovEst As New ClassTipoMovEst
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim bSaida As Boolean
Dim iEscaninho As Integer
Dim objEscaninho As New ClassEscaninho
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dFator As Double

On Error GoTo Erro_Traz_RastreamentoSerie_Tela

    Set objProduto = New ClassProduto

    QtdAlm.Caption = Formata_Estoque(objItemMovEstoque.dQuantidade)
    UM.Caption = objItemMovEstoque.sSiglaUM

    lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 141837

    Produto.Caption = sProdutoMascarado
    
    objProduto.sCodigo = objItemMovEstoque.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141838

    DescricaoProduto.Caption = objProduto.sDescricao

    Almoxarifado.Caption = objItemMovEstoque.sAlmoxarifadoNomeRed
    
    objAlmoxarifado.sNomeReduzido = objItemMovEstoque.sAlmoxarifadoNomeRed

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 177300
    
    objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
    
    Item.Caption = objItemMovEstoque.iItemNF
    
    lErro = CF("RastreamentoSerie_Obtem_Escaninho", objItemMovEstoque, objTipoMovEst, objEstoqueProduto, iEscaninho, bSaida)
    If lErro <> SUCESSO Then gError 177293
    
    objEscaninho.iCodigo = iEscaninho
    
    lErro = Escaninho_Obtem(objEscaninho)
    If lErro <> SUCESSO Then gError 177294
    
    TipoMovto.Caption = objTipoMovEst.iCodigo & SEPARADOR & objTipoMovEst.sDescricao
    
    Escaninho.Caption = objEscaninho.sNome
    
    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemMovEstoque.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 177301
    
    objItemMovEstoque.dQuantidadeEst = objItemMovEstoque.dQuantidade * dFator
    objItemMovEstoque.sSiglaUMEst = objProduto.sSiglaUMEstoque
        
    For Each objRastreamentoMovto In colRastreamentoMovto
            
        iIndice = iIndice + 1
        
        GridItens.TextMatrix(iIndice, iGrid_Serie_Col) = objRastreamentoMovto.sLote
    
'        lErro = Lote_Saida_Celula(objRastreamentoMovto.sLote, iIndice)
'        If lErro <> SUCESSO Then gError 141862
        
        If objRastreamentoMovto.dQuantidade = 0 Then
            objRastreamentoMovto.dQuantidade = objRastreamentoMovto.dQuantidadeEst / dFator
        Else
            objRastreamentoMovto.dQuantidadeEst = objRastreamentoMovto.dQuantidade * dFator
        End If
        
    Next
    
    objGridItens.iLinhasExistentes = iIndice
    
    Call Calcula_Totais
           
    Traz_RastreamentoSerie_Tela = SUCESSO

    Exit Function

Erro_Traz_RastreamentoSerie_Tela:

    Traz_RastreamentoSerie_Tela = gErr
    
    Select Case gErr
    
        Case 141837 To 141839, 141862, 177293, 177294, 177300, 177301
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141840)
    
    End Select
    
    Exit Function
    
End Function

Function Move_RastreamentoSerie_Memoria() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objRastreamentoLote As ClassRastreamentoLote
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoMovtoAux As ClassRastreamentoMovto
Dim dFator As Double

On Error GoTo Erro_Move_RastreamentoSerie_Memoria

    For iIndice = gcolRastreamentoMovto.Count To 1 Step -1
        Set objRastreamentoMovtoAux = gcolRastreamentoMovto.Item(iIndice)
        gcolRastreamentoMovto.Remove iIndice
    Next

    dFator = gobjItemMovEstoque.dQuantidadeEst / gobjItemMovEstoque.dQuantidade

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Serie_Col))) = 0 Then gError 177305
    
        Set objRastreamentoMovto = New ClassRastreamentoMovto
        Set objRastreamentoLote = New ClassRastreamentoLote
    
        objRastreamentoLote.sProduto = objRastreamentoMovtoAux.sProduto
        objRastreamentoLote.iFilialOP = objRastreamentoMovtoAux.iFilialOP
        objRastreamentoLote.sCodigo = GridItens.TextMatrix(iIndice, iGrid_Serie_Col)
        
        lErro = CF("RastreamentoLote_Le", objRastreamentoLote)
        If lErro <> SUCESSO And lErro <> 75710 Then gError 141841
    
        objRastreamentoMovto.dQuantidadeEst = 1
        objRastreamentoMovto.dQuantidade = 1 / dFator
        objRastreamentoMovto.lNumIntDocLote = objRastreamentoLote.lNumIntDoc
        objRastreamentoMovto.dQuantidadeSerie = StrParaDbl(QuantidadeTotal.Caption)
        objRastreamentoMovto.iFilialOP = objRastreamentoMovtoAux.iFilialOP
        objRastreamentoMovto.iTipoDocOrigem = objRastreamentoMovtoAux.iTipoDocOrigem
        objRastreamentoMovto.lNumIntDocLoteSerieIni = objRastreamentoMovtoAux.lNumIntDocLoteSerieIni
        objRastreamentoMovto.lNumIntDocOrigem = objRastreamentoMovtoAux.lNumIntDocOrigem
        objRastreamentoMovto.sLote = objRastreamentoLote.sCodigo
        objRastreamentoMovto.sProduto = objRastreamentoMovtoAux.sProduto
        objRastreamentoMovto.sSiglaUM = objRastreamentoMovtoAux.sSiglaUM
       
        gcolRastreamentoMovto.Add objRastreamentoMovto
        
    Next
    
    If Not gbPodeAlterarQtd Then
        If Abs(StrParaDbl(QuantidadeTotal.Caption) - gobjItemMovEstoque.dQuantidadeEst) > QTDE_ESTOQUE_DELTA Then gError 141943
    End If
          
    If StrParaDbl(QuantidadeTotal.Caption) > (StrParaDbl(QtdAlm.Caption) * dFator) + QTDE_ESTOQUE_DELTA Then gError 141825
    
    gobjItemMovEstoque.dQuantidadeEst = StrParaDbl(QuantidadeTotal.Caption)
    gobjItemMovEstoque.dQuantidade = StrParaDbl(QuantidadeTotal.Caption) / dFator
    
    Move_RastreamentoSerie_Memoria = SUCESSO

    Exit Function

Erro_Move_RastreamentoSerie_Memoria:

    Move_RastreamentoSerie_Memoria = gErr
    
    Select Case gErr
    
        Case 141841
        
        Case 141825
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_LOTE_MAIOR_MOVTO", gErr)

        Case 141943
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RASTRO_NAO_ALTERAVEL", gErr, Formata_Estoque(gobjItemMovEstoque.dQuantidade), QuantidadeTotal.Caption)
    
        Case 177305
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, iIndice)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141842)
    
    End Select
    
    Exit Function
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Número de Série")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Serie.Name)

    iGrid_Serie_Col = 1

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

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

Dim lErro As Long
Dim objRastreamentoSerie As ClassRastreamentoMovto

On Error GoTo Erro_GridItens_KeyDown

    If KeyCode = vbKeyDelete Then
    
        If GridItens.Row <> 0 And gcolRastreamentoMovto.Count <> 0 Then
            
            For Each objRastreamentoSerie In gcolRastreamentoMovto
                Exit For
            Next
        
            If GridItens.TextMatrix(GridItens.Row, iGrid_Serie_Col) = objRastreamentoSerie.sLote Then gError 141851

        End If

        Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    
    End If
   
    Call Calcula_Totais
   
    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 141851
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_SERIE_INICIAL_NAO_EXCLUR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141852)

    End Select
    
    Exit Sub

End Sub

Private Sub Calcula_Totais()
    
    QuantidadeTotal.Caption = Formata_Estoque(objGridItens.iLinhasExistentes)

End Sub

Private Function Escaninho_Obtem(objEscaninho As ClassEscaninho) As Long

On Error GoTo Erro_Escaninho_Obtem

    Select Case objEscaninho.iCodigo
    
        Case ESCANINHO_NOSSO
            objEscaninho.sNome = "Disponível"
        
        Case ESCANINHO_3_EM_CONSIGNACAO
            objEscaninho.sNome = "Consignação - De Terceiros em Nosso Poder"
        
        Case ESCANINHO_NOSSO_EM_CONSIGNACAO
            objEscaninho.sNome = "Consignação - Nosso em Poder de Terceiros"
        
        Case ESCANINHO_3_EM_DEMO
            objEscaninho.sNome = "Demonstração - De Terceiros em Nosso Poder"
        
        Case ESCANINHO_NOSSO_EM_DEMO
            objEscaninho.sNome = "Demonstração - Nosso em Poder de Terceiros"
        
        Case ESCANINHO_3_EM_CONSERTO
            objEscaninho.sNome = "Conserto - De Terceiros em Nosso Poder"
        
        Case ESCANINHO_NOSSO_EM_CONSERTO
            objEscaninho.sNome = "Conserto - Nosso em Poder de Terceiros"
        
        Case ESCANINHO_3_EM_OUTROS
            objEscaninho.sNome = "Outros - De Terceiros em Nosso Poder"
        
        Case ESCANINHO_NOSSO_EM_OUTROS
            objEscaninho.sNome = "Outros - Nosso em Poder de Terceiros"
        
        Case ESCANINHO_3_EM_BENEF
            objEscaninho.sNome = "Beneficiamento - De Terceiros em Nosso Poder"
        
        Case ESCANINHO_NOSSO_EM_BENEF
            objEscaninho.sNome = "Beneficiamento - Nosso em Poder de Terceiros"
        
        Case ESCANINHO_DEFEITUOSO
            objEscaninho.sNome = "Defeituoso"
        
        Case ESCANINHO_INDISPONIVEL
            objEscaninho.sNome = "Outros Indisponíveis"
                        
        Case Else
            objEscaninho.sNome = ""
        
    End Select

    Escaninho_Obtem = SUCESSO

    Exit Function

Erro_Escaninho_Obtem:

    Escaninho_Obtem = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177292)

    End Select
    
    Exit Function

End Function

Private Function Lote_Saida_Celula(ByVal sSerie As String, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim sSerieParteFixa As String
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iEscaninho As Integer
Dim bSaida As Boolean

On Error GoTo Erro_Lote_Saida_Celula
   
   'Se quantidade estiver preenchida
    If Len(Trim(sSerie)) > 0 Then
    
         'Formata o Produto para o BD
        lErro = CF("Produto_Formata", Produto.Caption, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 141843
   
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141936

        objRastroLote.sCodigo = sSerie
        objRastroLote.sProduto = sProdutoFormatado
                
        'Lê o Rastreamento do Lote vinculado ao produto
        lErro = CF("RastreamentoLote_Le", objRastroLote)
        If lErro <> SUCESSO And lErro <> 75710 Then gError 141844
        
        'Se não encontrou --> Erro
        'Não valida na produção entrada porque vai ser gerado
        If gsTelaOrigem = "ProducaoEntrada" Then
        
            If lErro = SUCESSO Then gError 177313

            sSerieParteFixa = Left(objProduto.sSerieProx, Len(objProduto.sSerieProx) - objProduto.iSerieParteNum)
            
            If Len(objProduto.sSerieProx) <> Len(sSerie) Then gError 141937
            
            If Not IsNumeric(Right(sSerie, objProduto.iSerieParteNum)) Then gError 141938
            
            If sSerieParteFixa <> Left(sSerie, Len(sSerieParteFixa)) Then gError 141939
        
        Else
            If lErro = 75710 Then gError 141845
            
            If gobjItemMovEstoque.iTipoMov <> 0 Then
            
                lErro = CF("RastreamentoSerie_Obtem_Escaninho", gobjItemMovEstoque, objTipoMovEstoque, objEstoqueProduto, iEscaninho, bSaida)
                If lErro <> SUCESSO Then gError 177294
                
                lErro = CF("RastreamentoSerie_Valida_Serie", gobjItemMovEstoque, objEstoqueProduto, sSerie)
                If lErro <> SUCESSO Then gError 177295
            
            End If
            
        End If
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If iIndice <> iLinha Then
                If sSerie = GridItens.TextMatrix(iIndice, iGrid_Serie_Col) Then gError 141846
            End If
        Next

        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If
        
    End If

    Call Calcula_Totais

    Lote_Saida_Celula = SUCESSO

    Exit Function

Erro_Lote_Saida_Celula:

    Lote_Saida_Celula = gErr
    
    Select Case gErr
    
        Case 141843, 141844, 141847, 141849, 177294, 177295
        
        Case 141845
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO", gErr, objRastroLote.sProduto, objRastroLote.sCodigo)
        
        Case 141846
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_PROD_ALMOX_JA_UTILIZADO_GRID", gErr, Serie.Text, Produto.Caption, Almoxarifado.Caption)
        
        Case 141937
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_TAMANHO_DIFERENTE", gErr, Len(objProduto.sSerieProx), Len(sSerie))
        
        Case 141938
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, Right(sSerie, objProduto.iSerieParteNum))
       
        Case 141939
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_PARTEFIXA_DIFERENTE", gErr, sSerieParteFixa, Left(sSerie, Len(sSerieParteFixa)))
        
        Case 177313
            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_CADASTRADO", gErr, objRastroLote.sProduto, objRastroLote.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141850)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case Serie.Name
            
            'Não pode alterar se for a série inicial
            If iLinha = 1 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Else
        
            objControl.Enabled = False
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162705)

    End Select

    Exit Sub

End Sub

