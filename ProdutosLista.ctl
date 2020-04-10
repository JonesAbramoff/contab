VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ProdutosLista 
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   11340
   Begin VB.CommandButton BotaoSeleciona 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3975
      Picture         =   "ProdutosLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1830
   End
   Begin VB.CommandButton BotaoFecha 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5880
      Picture         =   "ProdutosLista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridProdutos 
      Height          =   5625
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   9922
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ProdutosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjProdutos As ClassProduto
Dim iAlterado As Integer

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Nome_Col As Integer
Dim iGrid_Codigo_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Referencia_Col As Integer
Dim iGrid_CodigoBarras_Col As Integer
Dim iGrid_ICMS_Col As Integer
Dim iGrid_Preco_Col As Integer
Dim iGrid_Sigla_Col As Integer
Dim iGrid_SitTrib_Col As Integer
Dim iGrid_IAT_Col As Integer
Dim iGrid_IPPT_Col As Integer
Dim iGrid_Qtde_Est_Col As Integer
Dim iGrid_PercICMS_Col As Integer
Dim gcolProdutos As New Collection

Public Sub Form_Load()
    
Dim lRows As Long
    
On Error GoTo Erro_Form_Load

    Set gobjProdutos = New ClassProduto
        
        
    iGrid_Nome_Col = 0
    iGrid_Codigo_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_Referencia_Col = 3
    iGrid_Preco_Col = 4
    iGrid_Sigla_Col = 5
    iGrid_Qtde_Est_Col = 6
    iGrid_CodigoBarras_Col = 7
    iGrid_SitTrib_Col = 8
    iGrid_PercICMS_Col = 9
    iGrid_ICMS_Col = 10
    iGrid_IAT_Col = 11
    iGrid_IPPT_Col = 12
    
    GridProdutos.Cols = 13
    
    GridProdutos.TextMatrix(0, iGrid_Nome_Col) = "Nome"
    GridProdutos.TextMatrix(0, iGrid_Codigo_Col) = "Código"
    GridProdutos.TextMatrix(0, iGrid_Descricao_Col) = "Descrição"
    GridProdutos.TextMatrix(0, iGrid_Referencia_Col) = "Referência"
    GridProdutos.TextMatrix(0, iGrid_Preco_Col) = "Preço"
    GridProdutos.TextMatrix(0, iGrid_Sigla_Col) = "U.M."
    GridProdutos.TextMatrix(0, iGrid_Qtde_Est_Col) = "Quantidade"
    GridProdutos.TextMatrix(0, iGrid_CodigoBarras_Col) = "Código de Barras"
    GridProdutos.TextMatrix(0, iGrid_ICMS_Col) = "Sigla ICMS/ISS"
    GridProdutos.TextMatrix(0, iGrid_SitTrib_Col) = "Situação Tributária"
    GridProdutos.TextMatrix(0, iGrid_IAT_Col) = "IAT"
    GridProdutos.TextMatrix(0, iGrid_IPPT_Col) = "IPPT"
    GridProdutos.TextMatrix(0, iGrid_PercICMS_Col) = "% ICMS/ISS"
    
    GridProdutos.ColWidth(iGrid_Nome_Col) = 2600
    GridProdutos.ColWidth(iGrid_Descricao_Col) = 4000
    GridProdutos.ColWidth(iGrid_CodigoBarras_Col) = 2000
    GridProdutos.ColWidth(iGrid_SitTrib_Col) = 2000
    GridProdutos.ColWidth(iGrid_Codigo_Col) = 1500
    GridProdutos.ColWidth(iGrid_Qtde_Est_Col) = 1500
    GridProdutos.ColWidth(iGrid_PercICMS_Col) = 1500
    
    Call Preenche_Grid_Produtos
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
            
        Case 214843
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 165733)

    End Select

    Exit Sub

End Sub

Function Preenche_Grid_Produtos() As Long

Dim objProdutos As ClassProduto
Dim objProdutosCod As ClassProduto
Dim iIndice As Integer
Dim iLinhas As Integer
Dim bAchou As Boolean
Dim iIndice2 As Integer
Dim objAliquotaICMS As New ClassAliquotaICMS
Dim iAchou As Integer
Dim lErro As Long

On Error GoTo Erro_Preenche_Grid_Produtos

    Set gcolProdutos = New Collection

    lErro = CF_ECF("Produtos_Le_NomeReduzido1", gcolProdutos)
    If lErro <> SUCESSO Then gError 214843


    If gcolProdutos.Count > 8 Then
        GridProdutos.Rows = gcolProdutos.Count + 1
    Else
        GridProdutos.Rows = 9
    End If


    For Each objProdutos In gcolProdutos

        iLinhas = iLinhas + 1
                
        GridProdutos.TextMatrix(iLinhas, iGrid_Nome_Col) = objProdutos.sNomeReduzido
        If objProdutos.colCodBarras.Count > 0 Then GridProdutos.TextMatrix(iLinhas, iGrid_CodigoBarras_Col) = objProdutos.colCodBarras.Item(1)
'        GridProdutos.TextMatrix(iLinhas, iGrid_ICMS_Col) = objProdutos.sICMSAliquota
        GridProdutos.TextMatrix(iLinhas, iGrid_ICMS_Col) = left(objProdutos.sSituacaoTribECF, 1)
        GridProdutos.TextMatrix(iLinhas, iGrid_Preco_Col) = Format(objProdutos.dPrecoLoja, "standard")
        GridProdutos.TextMatrix(iLinhas, iGrid_Referencia_Col) = objProdutos.sReferencia
        GridProdutos.TextMatrix(iLinhas, iGrid_Sigla_Col) = objProdutos.sSiglaUMVenda
        GridProdutos.TextMatrix(iLinhas, iGrid_Qtde_Est_Col) = Format(objProdutos.dQuantEstLoja, "###########,###")
        If objProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ICMS Integral"
        ElseIf objProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ICMS Não Trib."
        ElseIf objProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ICMS Isento"
        ElseIf objProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ICMS Subst. Trib."
        ElseIf objProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ISS Integral"
        ElseIf objProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ISS Não Trib."
        ElseIf objProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_ISENTA Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ISS Isento"
        ElseIf objProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST Then
            GridProdutos.TextMatrix(iLinhas, iGrid_SitTrib_Col) = "ISS Subst. Trib."
        End If
        
'        If objProdutos.sCodigoBarras = "" And objProdutos.sReferencia = "" Then
'            For iIndice2 = 1 To gaobjProdutosCodBarras.Count
'                Set objProdutosCod = gaobjProdutosCodBarras.Item(iIndice2)
'                If objProdutosCod.sNomeReduzido = objProdutos.sNomeReduzido Then
'                     GridProdutos.TextMatrix(iLinhas, iGrid_CodigoBarras_Col) = objProdutosCod.sCodigoBarras
'                     Exit For
'                End If
'            Next
'        End If
        
        GridProdutos.TextMatrix(iLinhas, iGrid_Codigo_Col) = objProdutos.sCodigo
        GridProdutos.TextMatrix(iLinhas, iGrid_Descricao_Col) = objProdutos.sDescricao
        GridProdutos.TextMatrix(iLinhas, iGrid_IAT_Col) = objProdutos.sTruncamento
        GridProdutos.TextMatrix(iLinhas, iGrid_IPPT_Col) = IIf(objProdutos.iCompras = 1, "T", "P")
        
        iAchou = 0
        
        For Each objAliquotaICMS In gobjLojaECF.colAliquotaICMS
        
            If objAliquotaICMS.sSigla = objProdutos.sICMSAliquota Then
                GridProdutos.TextMatrix(iLinhas, iGrid_PercICMS_Col) = Format(objAliquotaICMS.dAliquota * 100, "Standard")
                iAchou = 1
                Exit For
            End If
            
        
        Next

        If iAchou = 0 Then
            GridProdutos.TextMatrix(iLinhas, iGrid_PercICMS_Col) = "0,00"
        End If
        
    Next
    
    Preenche_Grid_Produtos = SUCESSO
    
    Exit Function

Erro_Preenche_Grid_Produtos:

    Preenche_Grid_Produtos = gErr

    Select Case gErr
            
        Case 214843
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 165734)

    End Select

    Exit Function
    
    
End Function

Private Sub BotaoFecha_Click()
    
    Set gobjProdutos = Nothing
    giRetornoTela = vbCancel
    Unload Me
    
End Sub

Private Sub BotaoSeleciona_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim obj1 As Object

On Error GoTo Erro_BotaoSeleciona_Click
    
    If GridProdutos.Row = 0 Or GridProdutos.Row > gcolProdutos.Count Then Exit Sub
    
    gobjProdutos.sNomeReduzido = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Nome_Col)
    gobjProdutos.sCodigoBarras = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_CodigoBarras_Col)
    gobjProdutos.sICMSAliquota = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_ICMS_Col)
    gobjProdutos.dPrecoLoja = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Preco_Col))
    gobjProdutos.sReferencia = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Referencia_Col)
    gobjProdutos.sSiglaUMVenda = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Sigla_Col)
    gobjProdutos.dQuantEstLoja = StrParaDbl(GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Qtde_Est_Col))

    gobjProdutos.sCodigo = GridProdutos.TextMatrix(GridProdutos.Row, iGrid_Codigo_Col)
    
    If GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ICMS Integral" Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL
    ElseIf GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ICMS Não Trib." Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO
    ElseIf GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ICMS Isento" Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA
    ElseIf GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ICMS Subst. Trib." Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST
    ElseIf GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ISS Integral" Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL
    ElseIf GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ISS Não Trib." Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO
    ElseIf GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ISS Isento" Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_ISENTA
    ElseIf GridProdutos.TextMatrix(GridProdutos.Row, iGrid_SitTrib_Col) = "ISS Subst. Trib." Then
        gobjProdutos.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST
    End If
    
    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSeleciona_Click:

    Select Case Err
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 165734)

    End Select

    Exit Sub

End Sub

Private Sub GridProdutos_DblClick()
    
    Call BotaoSeleciona_Click
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

Function Trata_Parametros(objProduto As ClassProduto) As Long

Dim iInicio As Integer
Dim iFim As Integer
Dim iMeio As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjProdutos = objProduto
    
    If Len(Trim(gobjProdutos.sNomeReduzido)) > 0 Then
    
        iInicio = 1
        iFim = gaobjProdutosNome.Count

        Do While iFim >= iInicio

            iMeio = Fix((iInicio + iFim) / 2)

            If UCase(GridProdutos.TextMatrix(iMeio, iGrid_Nome_Col)) > UCase(gobjProdutos.sNomeReduzido) Then
                iFim = iMeio - 1
            Else
                If UCase(GridProdutos.TextMatrix(iMeio, iGrid_Nome_Col)) < UCase(gobjProdutos.sNomeReduzido) Then
                    iInicio = iMeio + 1
                Else
                    iInicio = iFim + 1
                End If
            End If
        Loop
        
        If UCase(GridProdutos.TextMatrix(iMeio, iGrid_Nome_Col)) < UCase(gobjProdutos.sNomeReduzido) And iMeio < iFim Then iMeio = iMeio + 1
            
        GridProdutos.Row = iMeio
        GridProdutos.RowSel = iMeio
        GridProdutos.Col = 0
        GridProdutos.ColSel = GridProdutos.Cols - 1
        SendKeys "{RIGHT}"
        
    End If
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err

        Case Else
        
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 165735)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Produtos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ProdutosLista"
    
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        Call BotaoSeleciona_Click
    End If
    
    
    'Clique em F8
    If KeyCode = vbKeyEscape Then
        Call BotaoFecha_Click
    End If
  
End Sub

