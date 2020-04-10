VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl CupomFiscalLista 
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10875
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   10875
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   735
      Left            =   5625
      Picture         =   "CupomFiscalLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3030
      Width           =   1830
   End
   Begin VB.CommandButton BotaoSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   735
      Left            =   3675
      Picture         =   "CupomFiscalLista.ctx":0272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3030
      Width           =   1860
   End
   Begin MSFlexGridLib.MSFlexGrid GridCupom 
      Height          =   2910
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   5133
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "CupomFiscalLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjCupom As ClassCupomFiscal
Public iAlterado As Integer
Dim gdQuant As Double

Const NUM_GRIDCUPOM_COLS = 6 + 7

'Variáveis Relacionadas as Colunas do Grid
Dim iGrid_ECF_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Hora_Col As Integer
Dim iGrid_COO_Col As Integer
Dim iGrid_ValorTotal_Col As Integer
Dim iGrid_Vendedor_Col As Integer

Dim iGrid_Tipo_Col As Integer
Dim iGrid_DAV_Col As Integer
Dim iGrid_CGC_Col As Integer
Dim iGrid_NomeCli_Col As Integer
Dim iGrid_ChvNFe_Col As Integer
Dim iGrid_Endereco_Col As Integer
Dim iGrid_QtdeItens_Col As Integer

Private Sub BotaoSelecionar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoSelecionar_Click

    If GridCupom.Row = 0 Or GridCupom.Row > gdQuant Then Exit Sub

    gobjCupom.iECF = StrParaInt(GridCupom.TextMatrix(GridCupom.Row, iGrid_ECF_Col))
    gobjCupom.dtDataEmissao = StrParaDate(GridCupom.TextMatrix(GridCupom.Row, iGrid_Data_Col))
    gobjCupom.dHoraEmissao = StrParaDate(GridCupom.TextMatrix(GridCupom.Row, iGrid_Hora_Col))
    gobjCupom.lNumero = StrParaLong(GridCupom.TextMatrix(GridCupom.Row, iGrid_COO_Col))
    gobjCupom.dValorTotal = StrParaDbl(GridCupom.TextMatrix(GridCupom.Row, iGrid_ValorTotal_Col))
    gobjCupom.iVendedor = StrParaInt(GridCupom.TextMatrix(GridCupom.Row, iGrid_Vendedor_Col))
    
    giRetornoTela = vbOK

    Unload Me

    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case Err

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 158427)

    End Select

    Exit Sub

End Sub

Private Sub GridCupom_DblClick()
    
    Call BotaoSelecionar_Click
    
End Sub

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjCupom = New ClassCupomFiscal
    
    iGrid_ECF_Col = 0
    iGrid_COO_Col = 1
    iGrid_Data_Col = 2
    iGrid_Hora_Col = 3
    iGrid_ValorTotal_Col = 4
    iGrid_Vendedor_Col = 5
    
    iGrid_Tipo_Col = 6
    iGrid_DAV_Col = 7
    iGrid_CGC_Col = 8
    iGrid_NomeCli_Col = 9
    iGrid_ChvNFe_Col = 10
    iGrid_Endereco_Col = 11
    iGrid_QtdeItens_Col = 12

    
    GridCupom.Cols = NUM_GRIDCUPOM_COLS

    GridCupom.TextMatrix(0, iGrid_ECF_Col) = "ECF"
    GridCupom.TextMatrix(0, iGrid_COO_Col) = "COO"
    GridCupom.TextMatrix(0, iGrid_Data_Col) = "Data Emissão"
    GridCupom.TextMatrix(0, iGrid_Hora_Col) = "Hora Emissão"
    GridCupom.TextMatrix(0, iGrid_ValorTotal_Col) = "Valor"
    GridCupom.TextMatrix(0, iGrid_Vendedor_Col) = "Vendedor"
        
    GridCupom.TextMatrix(0, iGrid_Tipo_Col) = "Tipo"
    GridCupom.TextMatrix(0, iGrid_DAV_Col) = "DAV"
    GridCupom.TextMatrix(0, iGrid_CGC_Col) = "CPF/CNPJ"
    GridCupom.TextMatrix(0, iGrid_NomeCli_Col) = "Nome.Cli."
    GridCupom.TextMatrix(0, iGrid_ChvNFe_Col) = "Chv.NFCe"
    GridCupom.TextMatrix(0, iGrid_Endereco_Col) = "Endereço"
    GridCupom.TextMatrix(0, iGrid_QtdeItens_Col) = "Qtd.Itens"
    
    GridCupom.ColWidth(iGrid_ECF_Col) = 600
    GridCupom.ColWidth(iGrid_COO_Col) = 600
    GridCupom.ColWidth(iGrid_Data_Col) = 1200
    GridCupom.ColWidth(iGrid_Hora_Col) = 1200
    GridCupom.ColWidth(iGrid_NomeCli_Col) = 2400
    GridCupom.ColWidth(iGrid_ChvNFe_Col) = 4200
    GridCupom.ColWidth(iGrid_Endereco_Col) = 3600
    GridCupom.ColWidth(iGrid_CGC_Col) = 1400
           
    If gcolVendas.Count > 8 Then
        GridCupom.Rows = gcolVendas.Count + 1
    Else
        GridCupom.Rows = 9
    End If
    
    lErro = Preenche_Grid_Cupom()
    If lErro <> SUCESSO Then gError 105351

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 105351

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 158428)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Grid_Cupom() As Long

Dim objCupom As ClassCupomFiscal
Dim iIndice As Integer
Dim objVenda As ClassVenda

On Error GoTo Erro_Preenche_Grid_Cupom
    
    For Each objVenda In gcolVendas
    
        If objVenda.iTipo = OPTION_CF Or (objVenda.iTipo = OPTION_ORCAMENTO And objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO) Then
        
            Set objCupom = objVenda.objCupomFiscal
        
            iIndice = iIndice + 1
            
            GridCupom.TextMatrix(iIndice, iGrid_ECF_Col) = objCupom.iECF
            GridCupom.TextMatrix(iIndice, iGrid_COO_Col) = IIf(objVenda.iTipo = OPTION_CF, objCupom.lNumero, objCupom.lNumOrcamento)
            GridCupom.TextMatrix(iIndice, iGrid_Data_Col) = objCupom.dtDataEmissao
            GridCupom.TextMatrix(iIndice, iGrid_Hora_Col) = CDate(objCupom.dHoraEmissao)
            GridCupom.TextMatrix(iIndice, iGrid_ValorTotal_Col) = Format(objCupom.dValorTotal, "STANDARD")
            GridCupom.TextMatrix(iIndice, iGrid_Vendedor_Col) = objCupom.iVendedor
        
            GridCupom.TextMatrix(iIndice, iGrid_Tipo_Col) = IIf(objVenda.iTipo = OPTION_CF, "CF", "") & IIf(objVenda.iTipo = OPTION_ORCAMENTO, "ORC", "") & IIf(objVenda.iTipo = OPTION_DAV, "DAV", "")
            GridCupom.TextMatrix(iIndice, iGrid_DAV_Col) = objCupom.lNumeroDAV
            GridCupom.TextMatrix(iIndice, iGrid_CGC_Col) = objCupom.sCPFCGC
            GridCupom.TextMatrix(iIndice, iGrid_NomeCli_Col) = objCupom.sNomeCliente
            GridCupom.TextMatrix(iIndice, iGrid_ChvNFe_Col) = objCupom.sNFeChaveAcesso
            GridCupom.TextMatrix(iIndice, iGrid_Endereco_Col) = objCupom.sEndereco
            GridCupom.TextMatrix(iIndice, iGrid_QtdeItens_Col) = CStr(objCupom.colItens.Count)
        
        End If
    
    Next
    
    gdQuant = iIndice
    
    GridCupom.Rows = iIndice + 1
    
    Preenche_Grid_Cupom = SUCESSO
    
    Exit Function

Erro_Preenche_Grid_Cupom:

    Preenche_Grid_Cupom = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 158429)

    
    
    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjCupom = Nothing

End Sub

Function Trata_Parametros(objCupom As ClassCupomFiscal) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjCupom = objCupom
    
    

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else

            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 158430)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Cupons"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CupomLista"

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

