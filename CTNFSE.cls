VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim m_objUserControl As Object


Private WithEvents objEventoLogLote As AdmEvento
Attribute objEventoLogLote.VB_VarHelpID = -1
Private WithEvents objEventoLotes As AdmEvento
Attribute objEventoLotes.VB_VarHelpID = -1
Private WithEvents objEventoRetCanc As AdmEvento
Attribute objEventoRetCanc.VB_VarHelpID = -1
Private WithEvents objEventoRetConsulta As AdmEvento
Attribute objEventoRetConsulta.VB_VarHelpID = -1
Private WithEvents objEventoRetNFe As AdmEvento
Attribute objEventoRetNFe.VB_VarHelpID = -1
Private WithEvents objEventoStatusNFe As AdmEvento
Attribute objEventoStatusNFe.VB_VarHelpID = -1


Private WithEvents objEventoNFiscalDe As AdmEvento
Attribute objEventoNFiscalDe.VB_VarHelpID = -1
Private WithEvents objEventoNFiscalAte As AdmEvento
Attribute objEventoNFiscalAte.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorDe As AdmEvento
Attribute objEventoFornecedorDe.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorAte As AdmEvento
Attribute objEventoFornecedorAte.VB_VarHelpID = -1


Dim iAlterado As Integer
Dim iTabPrincipalAlterado As Integer
Dim iFrameAtual As Integer

Dim asOrdenacao(8) As String
Dim asOrdenacaoString(8) As String


Const TABSELECAO = 1
Const TABNF = 2

Const NUM_MAX_NF_POR_LOTE = 50

Dim objGridNFiscal As AdmGrid
Const GRID_SELECIONA_COL = 1
Const GRID_SERIE_COL = 2
Const GRID_NUMNFISCAL_COL = 3
Const GRID_DATAEMISSAO_COL = 4
Const GRID_TIPONF_COL = 5
Const GRID_CODCLIENTE_COL = 6
Const GRID_NOMECLIENTE_COL = 7
Const GRID_FILIALCLI_COL = 8
Const GRID_CODFORNECEDOR_COL = 9
Const GRID_NOMEFORNECEDOR_COL = 10
Const GRID_FILIALFORN_COL = 11
Const GRID_VALOR_COL = 12
Const GRID_LOTE_COL = 13
Const GRID_STATUS_COL = 14
Const GRID_NUMINTNF_COL = 15


Public Property Get SerieDe() As Object
     Set SerieDe = objUserControl.Controls("SerieDe")
End Property

Public Property Get SerieAte() As Object
     Set SerieAte = objUserControl.Controls("SerieAte")
End Property

Public Property Get NFiscalDe() As Object
     Set NFiscalDe = objUserControl.Controls("NFiscalDe")
End Property

Public Property Get NFiscalAte() As Object
     Set NFiscalAte = objUserControl.Controls("NFiscalAte")
End Property

Public Property Get ClienteDe() As Object
     Set ClienteDe = objUserControl.Controls("ClienteDe")
End Property

Public Property Get ClienteAte() As Object
     Set ClienteAte = objUserControl.Controls("ClienteAte")
End Property

Public Property Get FornecedorDe() As Object
     Set FornecedorDe = objUserControl.Controls("FornecedorDe")
End Property

Public Property Get FornecedorAte() As Object
     Set FornecedorAte = objUserControl.Controls("FornecedorAte")
End Property

Public Property Get DataEmissaoDe() As Object
     Set DataEmissaoDe = objUserControl.Controls("DataEmissaoDe")
End Property

Public Property Get DataEmissaoAte() As Object
     Set DataEmissaoAte = objUserControl.Controls("DataEmissaoAte")
End Property

Public Property Get UpDownEmissaoDe() As Object
     Set UpDownEmissaoDe = objUserControl.Controls("UpDownEmissaoDe")
End Property

Public Property Get UpDownEmissaoAte() As Object
     Set UpDownEmissaoAte = objUserControl.Controls("UpDownEmissaoAte")
End Property

Public Property Get Seleciona() As Object
     Set Seleciona = objUserControl.Controls("Seleciona")
End Property

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
End Property

Public Property Get NumNFiscal() As Object
     Set NumNFiscal = objUserControl.Controls("NumNFiscal")
End Property
    
Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property
    
Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
End Property
    
Public Property Get CodCliente() As Object
     Set CodCliente = objUserControl.Controls("CodCliente")
End Property
    
Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property
    
Public Property Get FilialCli() As Object
     Set FilialCli = objUserControl.Controls("FilialCli")
End Property
    
Public Property Get CodFornecedor() As Object
     Set CodFornecedor = objUserControl.Controls("CodFornecedor")
End Property
    
Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property
    
Public Property Get FilialForn() As Object
     Set FilialForn = objUserControl.Controls("FilialForn")
End Property
    
Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property
    
Public Property Get LoteGrid() As Object
     Set LoteGrid = objUserControl.Controls("LoteGrid")
End Property
    
Public Property Get Status() As Object
     Set Status = objUserControl.Controls("Status")
End Property
    
Public Property Get NumIntNF() As Object
     Set NumIntNF = objUserControl.Controls("NumIntNF")
End Property
    
Public Property Get GridNFiscal() As Object
     Set GridNFiscal = objUserControl.Controls("GridNFiscal")
End Property

Public Property Get Lote() As Object
     Set Lote = objUserControl.Controls("Lote")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get Ordenados() As Object
     Set Ordenados = objUserControl.Controls("Ordenados")
End Property

Public Property Get NaoEnviadas() As Object
     Set NaoEnviadas = objUserControl.Controls("NaoEnviadas")
End Property

Public Property Get EnviadasNaoAceitas() As Object
     Set EnviadasNaoAceitas = objUserControl.Controls("EnviadasNaoAceitas")
End Property

Public Property Get Ambas() As Object
     Set Ambas = objUserControl.Controls("Ambas")
End Property

Public Property Get ItensSelecionados() As Object
     Set ItensSelecionados = objUserControl.Controls("ItensSelecionados")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim bBloqueioFilial As Boolean, iFilialComPermissao As Integer

On Error GoTo Erro_Form_Load

    Set objGridNFiscal = New AdmGrid

    Set objEventoNFiscalDe = New AdmEvento
    Set objEventoNFiscalAte = New AdmEvento
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    Set objEventoFornecedorDe = New AdmEvento
    Set objEventoFornecedorAte = New AdmEvento
    
    Set objEventoLogLote = New AdmEvento
    Set objEventoLotes = New AdmEvento
    Set objEventoRetCanc = New AdmEvento
    Set objEventoRetConsulta = New AdmEvento
    Set objEventoRetNFe = New AdmEvento
    Set objEventoStatusNFe = New AdmEvento
    
    lErro = CF("Verifica_Bloqueios_Filial", giFilialEmpresa, bBloqueioFilial, iFilialComPermissao)
    If lErro <> SUCESSO Then gError 204985
    
    If bBloqueioFilial Then gError 204986
    
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 204987
    
    lErro = Inicializa_GridNFiscal(objGridNFiscal)
    If lErro <> SUCESSO Then gError 204988
    
    asOrdenacao(0) = "NFiscal.Serie"
    asOrdenacao(1) = "NFiscal.NumNotaFiscal"
    asOrdenacao(2) = "NFiscal.DataEmissao"
    asOrdenacao(3) = "TiposDocInfo.Sigla"
    asOrdenacao(4) = "NFiscal.Cliente"
    asOrdenacao(5) = "Clientes.NomeReduzido"
    asOrdenacao(6) = "NFiscal.Fornecedor"
    asOrdenacao(7) = "Fornecedores.NomeReduzido"
    asOrdenacao(8) = "NFiscal.ValorTotal"

    asOrdenacaoString(0) = "Série"
    asOrdenacaoString(1) = "Nota Fiscal"
    asOrdenacaoString(2) = "Data Emissão"
    asOrdenacaoString(3) = "Tipo"
    asOrdenacaoString(4) = "Codigo Cliente"
    asOrdenacaoString(5) = "Cliente"
    asOrdenacaoString(6) = "Codigo Fornecedor"
    asOrdenacaoString(7) = "Fornecedor"
    asOrdenacaoString(8) = "Valor"

    For iIndice = 0 To 8
        Ordenados.AddItem asOrdenacaoString(iIndice)
    Next

    iFrameAtual = 1

    iAlterado = 0
    iTabPrincipalAlterado = REGISTRO_ALTERADO

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 204985, 204987, 204988
        
        Case 204986
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_FILIAL_AUX", gErr, iFilialComPermissao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204989)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGridNFiscal = Nothing

    Set objEventoNFiscalDe = Nothing
    Set objEventoNFiscalAte = Nothing
    Set objEventoClienteDe = Nothing
    Set objEventoClienteAte = Nothing
    Set objEventoFornecedorDe = Nothing
    Set objEventoFornecedorAte = Nothing

    Set objEventoLogLote = Nothing
    Set objEventoLotes = Nothing
    Set objEventoRetCanc = Nothing
    Set objEventoRetConsulta = Nothing
    Set objEventoRetNFe = Nothing
    Set objEventoStatusNFe = Nothing

    Set objUserControl = Nothing
    
End Sub

Private Function Carrega_Serie() As Long
'Carrega as combos de Série e serie de NF original com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 204990

    'Carrega na combo
    For Each objSerie In colSerie
        SerieDe.AddItem objSerie.sSerie
        SerieAte.AddItem objSerie.sSerie
    Next
    
    SerieDe.AddItem " "
    SerieAte.AddItem " "
            
    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 204990

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204991)

    End Select

    Exit Function

End Function

Public Sub SerieDe_Click()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SerieAte_Click()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NaoEnviadas_Click()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EnviadasNaoAceitas_Click()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Ambas_Click()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LabelNFiscalDe_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection

    'Preenche PedidoDe com o pedido da tela
    If Len(Trim(NFiscalDe.ClipText)) > 0 Then objNFiscal.lNumNotaFiscal = CLng(NFiscalDe.Text)

    objNFiscal.iFilialEmpresa = giFilialEmpresa

    'Chama Tela NFiscalLista
    Call Chama_Tela("NFiscalInternaLista", colSelecao, objNFiscal, objEventoNFiscalDe, "RPS = 1")

End Sub

Private Sub objEventoNFiscalDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscalDe_evSelecao

    Set objNFiscal = obj1

    NFiscalDe.Text = CStr(objNFiscal.lNumNotaFiscal)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoNFiscalDe_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204992)

    End Select

    Exit Sub

End Sub

Public Sub LabelNFiscalAte_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection

    'Preenche PedidoDe com o pedido da tela
    If Len(Trim(NFiscalAte.ClipText)) > 0 Then objNFiscal.lNumNotaFiscal = CLng(NFiscalAte.Text)

    objNFiscal.iFilialEmpresa = giFilialEmpresa

    'Chama Tela NFiscalLista
    Call Chama_Tela("NFiscalInternaLista", colSelecao, objNFiscal, objEventoNFiscalAte, "RPS = 1")

End Sub

Private Sub objEventoNFiscalAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscalAte_evSelecao

    Set objNFiscal = obj1

    NFiscalAte.Text = CStr(objNFiscal.lNumNotaFiscal)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoNFiscalAte_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204993)

    End Select

    Exit Sub

End Sub

Public Sub ClienteDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ClienteDe_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(ClienteDe, iAlterado)
    iTabPrincipalAlterado = iTabAux
    
End Sub

Public Sub ClienteAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ClienteAte_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(ClienteAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche ClienteDe com o cliente da tela
    If Len(Trim(ClienteDe.ClipText)) > 0 Then objCliente.lCodigo = CLng(ClienteDe.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)

End Sub

Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente

On Error GoTo Erro_objEventoClienteDe_evSelecao

    Set objCliente = obj1

    ClienteDe.Text = CStr(objCliente.lCodigo)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoClienteDe_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204994)

    End Select

    Exit Sub

End Sub

Public Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche ClienteAte com o cliente da tela
    If Len(Trim(ClienteAte.ClipText)) > 0 Then objCliente.lCodigo = CLng(ClienteAte.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente

On Error GoTo Erro_objEventoClienteAte_evSelecao

    Set objCliente = obj1

    ClienteAte.Text = CStr(objCliente.lCodigo)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoClienteAte_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204995)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FornecedorDe_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    iTabPrincipalAlterado = iTabAux
    
End Sub

Public Sub FornecedorAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FornecedorAte_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub LabelFornecedorDe_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche FornecedorDe com o Fornecedor da tela
    If Len(Trim(FornecedorDe.ClipText)) > 0 Then objFornecedor.lCodigo = CLng(FornecedorDe.Text)

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorDe)

End Sub

Private Sub objEventoFornecedorDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_objEventoFornecedorDe_evSelecao

    Set objFornecedor = obj1

    FornecedorDe.Text = CStr(objFornecedor.lCodigo)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoFornecedorDe_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204996)

    End Select

    Exit Sub

End Sub

Public Sub LabelFornecedorAte_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche FornecedorAte com o Fornecedor da tela
    If Len(Trim(FornecedorAte.ClipText)) > 0 Then objFornecedor.lCodigo = CLng(FornecedorAte.Text)

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorAte)

End Sub

Private Sub objEventoFornecedorAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_objEventoFornecedorAte_evSelecao

    Set objFornecedor = obj1

    FornecedorAte.Text = CStr(objFornecedor.lCodigo)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoFornecedorAte_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204997)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissaoDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissaoDe_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub DataEmissaoDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Se a DataEmissaoDe está preenchida
    If Len(DataEmissaoDe.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoDe é válida
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 204998

    Exit Sub

Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 204998

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204999)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissaoAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissaoAte_GotFocus()

Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub DataEmissaoAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Se a DataEmissaoAte está preenchida
    If Len(DataEmissaoAte.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoDe é válida
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 207000

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 207000

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207001)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 207002

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 207002

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207003)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 207004

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 207004

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207005)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 207006

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 207006

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207007)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 207008

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 207008

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207009)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridNFiscal(objGridInt As AdmGrid) As Long
'Inicializa o GridPV
Dim lErro  As Long

On Error GoTo Erro_Inicializa_GridNFiscal

    'Form do GridPV
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Gera NFse")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Nota Fiscal")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("CodCli")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("CodForn")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Lote")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add (" ")

    'Controles que participam do GridPV
    objGridInt.colCampo.Add (Seleciona.Name)
    objGridInt.colCampo.Add (Serie.Name)
    objGridInt.colCampo.Add (NumNFiscal.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (TipoNFiscal.Name)
    objGridInt.colCampo.Add (CodCliente.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (FilialCli.Name)
    objGridInt.colCampo.Add (CodFornecedor.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (FilialForn.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (LoteGrid.Name)
    objGridInt.colCampo.Add (Status.Name)
    objGridInt.colCampo.Add (NumIntNF.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridNFiscal

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10


    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

'    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'Largura da primeira coluna
    GridNFiscal.ColWidth(0) = 400
    GridNFiscal.ColWidth(15) = 0

    Inicializa_GridNFiscal = SUCESSO

    Exit Function

Erro_Inicializa_GridNFiscal:

    Inicializa_GridNFiscal = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207010)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Col

            Case GRID_SELECIONA_COL

                lErro = Saida_Celula_Seleciona(objGridInt)
                If lErro <> SUCESSO Then gError 207011

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 207012
    
    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 207011

        Case 207012
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207013)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Seleciona(objGridInt As AdmGrid) As Long
'faz a critica da celula GeraOP do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Seleciona

    Set objGridInt.objControle = Seleciona
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 207014

    Saida_Celula_Seleciona = SUCESSO

    Exit Function

Erro_Saida_Celula_Seleciona:

    Saida_Celula_Seleciona = gErr

    Select Case gErr

        Case 207014
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207015)

    End Select

    Exit Function

End Function

Public Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer

On Error GoTo Erro_TabStrip1_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
        
    If (Cancel = False) Then
    
        'Se Frame atual é a seleção
        If TabStrip1.SelectedItem.Index = TABSELECAO Then

            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
                
                Call Grid_Limpa(objGridNFiscal)
                
                lErro = Trata_TabNFiscal()
                If lErro <> SUCESSO Then gError 207016

            End If
        
        End If

    End If
        
    Exit Sub
        
Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 207016

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207017)

    End Select

    Exit Sub

End Sub

Private Function Trata_TabNFiscal() As Long

Dim lErro As Long
Dim objGeracaoNFe As New ClassGeracaoNFe
Dim colNFe1 As New Collection

On Error GoTo Erro_Trata_TabNFiscal

    If Ordenados.ListIndex = -1 Then

        Ordenados.ListIndex = 1

    Else

        lErro = Move_TabSelecao_Memoria(objGeracaoNFe)
        If lErro <> SUCESSO Then gError 207018

        'Preenche a Coleção de Notas Fiscais
        lErro = CF("GeracaoRPSWEB_SelecionarItens", giFilialEmpresa, objGeracaoNFe, colNFe1)
        If lErro <> SUCESSO Then gError 207019

        'Preenche o Grid
        Call GridNFiscal_Preenche(colNFe1)

    End If

    Exit Function

Erro_Trata_TabNFiscal:

    Trata_TabNFiscal = gErr

    Select Case gErr

        Case 207018, 207019

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207020)

    End Select

    Exit Function

End Function

Private Function Move_TabSelecao_Memoria(objGeracaoNFe As ClassGeracaoNFe) As Long

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    lErro = Critica_Campos()
    If lErro <> SUCESSO Then gError 207021
    
    If NaoEnviadas.Value = True Then
        objGeracaoNFe.iOpcaoNFiscais = GERACAONFE_NAO_ENVIADOS
    ElseIf EnviadasNaoAceitas.Value = True Then
        objGeracaoNFe.iOpcaoNFiscais = GERACAONFE_NAO_ACEITOS
    Else
        objGeracaoNFe.iOpcaoNFiscais = GERACAONFE_AMBOS
    End If
        

    If Len(Trim(SerieDe.Text)) > 0 Then
        objGeracaoNFe.sSerieDe = SerieDe.Text
    End If

   If Len(Trim(SerieAte.Text)) > 0 Then
        objGeracaoNFe.sSerieAte = SerieAte.Text
    End If

    objGeracaoNFe.lNFiscalDe = StrParaLong(NFiscalDe.Text)
        
    objGeracaoNFe.lNFiscalAte = StrParaLong(NFiscalAte.Text)
        
    objGeracaoNFe.lClienteDe = StrParaLong(ClienteDe.Text)

    objGeracaoNFe.lClienteAte = StrParaLong(ClienteAte.Text)

    objGeracaoNFe.lFornecedorDe = StrParaLong(FornecedorDe.Text)

    objGeracaoNFe.lFornecedorAte = StrParaLong(FornecedorAte.Text)

    objGeracaoNFe.dtEmissaoDe = StrParaDate(DataEmissaoDe.Text)

    objGeracaoNFe.dtEmissaoAte = StrParaDate(DataEmissaoAte.Text)

    objGeracaoNFe.sOrdenacao = asOrdenacao(Ordenados.ListIndex)

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr
    
    Select Case gErr

        Case 207021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207022)

    End Select

    Exit Function

End Function

Private Function Critica_Campos() As Long

Dim lErro As Long
Dim sProdutoDe As String, sProdutoAte As String
Dim iPreenchidoDe As Integer, iPreenchidoAte As Integer

On Error GoTo Erro_Critica_Campos

    'Se NfiscalDe e NFiscalAte estão preenchidos
    If Len(Trim(NFiscalDe.ClipText)) > 0 And Len(Trim(NFiscalAte.ClipText)) > 0 Then

        'Verifica se NFiscalAte é maior ou igual que NFiscalDe
        If CLng(Trim(NFiscalAte.Text)) < CLng(Trim(NFiscalDe.Text)) Then gError 207023

    End If
    
       'Se ClienteAte e ClienteDe estão preenchidos
    If Len(Trim(ClienteDe.Text)) > 0 And Len(Trim(ClienteAte.Text)) > 0 Then

        'Verifica se ClienteAte é maior ou igual que ClienteDe
        If CLng(Trim(ClienteAte.Text)) < CLng(Trim(ClienteDe.Text)) Then gError 207024

    End If

       'Se FornecedorAte e FornecedorDe estão preenchidos
    If Len(Trim(FornecedorDe.Text)) > 0 And Len(Trim(FornecedorAte.Text)) > 0 Then

        'Verifica se FornecedorAte é maior ou igual que FornecedorDe
        If CLng(Trim(FornecedorAte.Text)) < CLng(Trim(FornecedorDe.Text)) Then gError 207025

    End If


    If Len(Trim(DataEmissaoDe.ClipText)) > 0 And Len(Trim(DataEmissaoAte.ClipText)) > 0 Then

        If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then gError 207026
    
    End If

    Critica_Campos = SUCESSO
    
    Exit Function
    
Erro_Critica_Campos:
    
    Critica_Campos = gErr
    
    Select Case gErr
        
        Case 207023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCALDE_MAIOR_NFISCALATE", gErr)
        
        Case 207024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", gErr)
        
        Case 207025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORDE_MAIOR_FORNECEDORATE", gErr)
        
        Case 207026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMISSAODE_MAIOR_EMISSAOATE", gErr)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207027)
    
    End Select
    
    Exit Function

End Function

Private Sub GridNFiscal_Preenche(colNFe1 As Collection)
'Preenche o Grid Pedido com os dados de colItens

Dim iLinha As Integer
Dim objNFe1 As ClassNFe1
Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objFilialCliente As New ClassFilialCliente
Dim sNomeRedCliente As String

On Error GoTo Erro_GridNFiscal_Preenche

    Call Grid_Limpa(objGridNFiscal)

    'Se o número de itens for maior que o número de linhas do Grid
    
    If colNFe1.Count + 1 > GridNFiscal.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de itens
        GridNFiscal.Rows = colNFe1.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridNFiscal)

    End If

    iLinha = 0

    'Percorre todas as OP da Coleção
    For Each objNFe1 In colNFe1

        iLinha = iLinha + 1

        'Passa para a tela os dados da OP em questão
        GridNFiscal.TextMatrix(iLinha, GRID_SERIE_COL) = objNFe1.objNFiscal.sSerie
        GridNFiscal.TextMatrix(iLinha, GRID_NUMNFISCAL_COL) = objNFe1.objNFiscal.lNumNotaFiscal
        GridNFiscal.TextMatrix(iLinha, GRID_TIPONF_COL) = objNFe1.sTipoNFiscal
        GridNFiscal.TextMatrix(iLinha, GRID_DATAEMISSAO_COL) = objNFe1.objNFiscal.dtDataEmissao
        GridNFiscal.TextMatrix(iLinha, GRID_TIPONF_COL) = objNFe1.sTipoNFiscal
        
        If objNFe1.objNFiscal.lCliente <> 0 Then
            GridNFiscal.TextMatrix(iLinha, GRID_CODCLIENTE_COL) = objNFe1.objNFiscal.lCliente
            GridNFiscal.TextMatrix(iLinha, GRID_NOMECLIENTE_COL) = objNFe1.sClienteNomeReduzido
            GridNFiscal.TextMatrix(iLinha, GRID_FILIALCLI_COL) = objNFe1.objNFiscal.iFilialCli
        End If
        
        If objNFe1.objNFiscal.lFornecedor <> 0 Then
            GridNFiscal.TextMatrix(iLinha, GRID_CODFORNECEDOR_COL) = objNFe1.objNFiscal.lFornecedor
            GridNFiscal.TextMatrix(iLinha, GRID_NOMEFORNECEDOR_COL) = objNFe1.sFornecedorNomeReduzido
            GridNFiscal.TextMatrix(iLinha, GRID_FILIALFORN_COL) = objNFe1.objNFiscal.iFilialForn
        End If
        
        GridNFiscal.TextMatrix(iLinha, GRID_VALOR_COL) = objNFe1.objNFiscal.dValorTotal
        
        If objNFe1.lLote <> 0 Then
            GridNFiscal.TextMatrix(iLinha, GRID_LOTE_COL) = objNFe1.lLote
        End If
        
        If Len(Trim(objNFe1.sstatus)) <> 0 Then
            GridNFiscal.TextMatrix(iLinha, GRID_STATUS_COL) = objNFe1.sstatus
        End If
        
        
        GridNFiscal.TextMatrix(iLinha, GRID_NUMINTNF_COL) = objNFe1.objNFiscal.lNumIntDoc
        
    Next

    Call Grid_Refresh_Checkbox(objGridNFiscal)

    'Passa para o Obj o número de itens passados pela Coleção
    objGridNFiscal.iLinhasExistentes = colNFe1.Count
    
    ItensSelecionados.Caption = ""

    Exit Sub

Erro_GridNFiscal_Preenche:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207028)

    End Select

    Exit Sub

End Sub

Public Sub Seleciona_Click()

Dim iClick As Integer
Dim lErro As Long

On Error GoTo Erro_Seleciona_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se é alguma linha válida
    If GridNFiscal.Row > objGridNFiscal.iLinhasExistentes Then Exit Sub


    'Verifica se está selecionando ou desselecionando
    If Len(Trim(GridNFiscal.TextMatrix(GridNFiscal.Row, GRID_SELECIONA_COL))) > 0 Then
        iClick = CInt(GridNFiscal.TextMatrix(GridNFiscal.Row, GRID_SELECIONA_COL)) = 1
    End If

    If iClick = True Then
        ItensSelecionados.Caption = StrParaInt(ItensSelecionados.Caption) + 1
    Else
        ItensSelecionados.Caption = StrParaInt(ItensSelecionados.Caption) - 1
    End If
    
    If StrParaInt(ItensSelecionados.Caption) > NUM_MAX_NF_POR_LOTE Then
        GridNFiscal.TextMatrix(GridNFiscal.Row, GRID_SELECIONA_COL) = DESMARCADO
        Seleciona.Value = vbUnchecked
        Call Grid_Refresh_Checkbox(objGridNFiscal)
        gError 207028
    End If

    Exit Sub

Erro_Seleciona_Click:

    Select Case gErr

        Case 207028
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_NUM_MAX_LOTE", gErr, NUM_MAX_NF_POR_LOTE)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207029)

    End Select

    Exit Sub

End Sub

Public Sub Seleciona_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscal)
End Sub

Public Sub Seleciona_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscal)
End Sub

Public Sub Seleciona_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridNFiscal.objControle = Seleciona
    lErro = Grid_Campo_Libera_Foco(objGridNFiscal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Lote_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Public Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Ordenados_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objGeracaoNFe As New ClassGeracaoNFe
Dim colNFe1 As New Collection
Dim colNFiscal As New Collection
Dim objNFiscal As ClassNFiscal
Dim iLinha As Integer

On Error GoTo Erro_Ordenados_Click

    If Ordenados.ListIndex = -1 Then Exit Sub

    '**** guarda as notas marcadas *********
    For iLinha = 1 To objGridNFiscal.iLinhasExistentes
        If StrParaInt(GridNFiscal.TextMatrix(iLinha, GRID_SELECIONA_COL)) = 1 Then
            Set objNFiscal = New ClassNFiscal
            objNFiscal.sSerie = GridNFiscal.TextMatrix(iLinha, GRID_SERIE_COL)
            objNFiscal.lNumNotaFiscal = StrParaLong(GridNFiscal.TextMatrix(iLinha, GRID_NUMNFISCAL_COL))
            colNFiscal.Add objNFiscal
        End If
    Next


    '**** carrega as notas com a nova ordenacao **********
    lErro = Move_TabSelecao_Memoria(objGeracaoNFe)
    If lErro <> SUCESSO Then gError 207030

    'Preenche a Coleção de Notas Fiscais
    lErro = CF("GeracaoRPSWEB_SelecionarItens", giFilialEmpresa, objGeracaoNFe, colNFe1)
    If lErro <> SUCESSO Then gError 207031

    'Preenche o Grid
    Call GridNFiscal_Preenche(colNFe1)


    '***** marca as notas que estavam selecionadas **********
    For Each objNFiscal In colNFiscal
        
        For iLinha = 1 To objGridNFiscal.iLinhasExistentes
            If GridNFiscal.TextMatrix(iLinha, GRID_SERIE_COL) = objNFiscal.sSerie And StrParaLong(GridNFiscal.TextMatrix(iLinha, GRID_NUMNFISCAL_COL)) = objNFiscal.lNumNotaFiscal Then
                GridNFiscal.TextMatrix(iLinha, GRID_SELECIONA_COL) = "1"
                Exit For
            End If
        Next
        
    Next
    
    Call Grid_Refresh_Checkbox(objGridNFiscal)
    
    ItensSelecionados.Caption = colNFiscal.Count
    
    Exit Sub

Erro_Ordenados_Click:

    Select Case gErr

        Case 207030, 207031

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207032)

    End Select

    Exit Sub

End Sub

Public Sub ProxNumLote_Click()

Dim lLote As Long
Dim lErro As Long

On Error GoTo Erro_ProxNumLote_Click

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt_Trans", "FATConfig", "NUM_PROX_LOTE_RPSWEB", lLote)
    If lErro <> SUCESSO Then gError 207033

    Lote.Text = lLote

    Exit Sub

Erro_ProxNumLote_Click:

    Select Case gErr

        Case 207033

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207034)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGerarNFe_Click()
'codigo da OP tem que estar preenchido, bem como a data e gobjGeracaoOP.colItens tem que ter itens selecionados

Dim lErro As Long

On Error GoTo Erro_BotaoGerarNFe_Click

    If gobjCRFAT.iUsaNFSE = MARCADO Then

        lErro = Gravar_Registro()
        If lErro <> SUCESSO Then gError 207035
    
        lErro = Limpa_Tela_GeracaoNFe
        If lErro <> SUCESSO Then gError 207036
    
        iTabPrincipalAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_BotaoGerarNFe_Click:

    Select Case gErr

        Case 207035, 207036

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207037)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_GeracaoNFe() As Long

Dim lErro As Long


On Error GoTo Erro_Limpa_Tela_GeracaoNFe

    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridNFiscal)

    ItensSelecionados.Caption = ""

    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = 1
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    iTabPrincipalAlterado = 0
    iAlterado = 0
    

    Limpa_Tela_GeracaoNFe = SUCESSO

    Exit Function

Erro_Limpa_Tela_GeracaoNFe:

    Limpa_Tela_GeracaoNFe = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207038)

    End Select

    Exit Function

End Function

Public Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        'Torna Frame de Pedido visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207039)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is NFiscalDe Then
            Call LabelNFiscalDe_Click
        ElseIf Me.ActiveControl Is NFiscalAte Then
            Call LabelNFiscalAte_Click
        ElseIf Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is FornecedorDe Then
            Call LabelFornecedorDe_Click
        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call LabelFornecedorAte_Click
        End If
    End If

End Sub

Public Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridNFiscal.iLinhasExistentes

        If iLinha <= NUM_MAX_NF_POR_LOTE Then
            'Marca na tela o pedido em questão
            GridNFiscal.TextMatrix(iLinha, GRID_SELECIONA_COL) = S_MARCADO
        Else
            GridNFiscal.TextMatrix(iLinha, GRID_SELECIONA_COL) = S_DESMARCADO
        End If
    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridNFiscal)

    If objGridNFiscal.iLinhasExistentes <= 50 Then
        ItensSelecionados.Caption = objGridNFiscal.iLinhasExistentes
    Else
        ItensSelecionados.Caption = NUM_MAX_NF_POR_LOTE
        Call Rotina_Aviso(vbOK, "AVISO_LOTENFE_LIMITE_NFS")
    End If

End Sub

Public Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    iAlterado = REGISTRO_ALTERADO

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridNFiscal.iLinhasExistentes

        'Desmarca na tela o pedido em questão
        GridNFiscal.TextMatrix(iLinha, GRID_SELECIONA_COL) = S_DESMARCADO

    Next

    ItensSelecionados.Caption = "0"

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridNFiscal)

End Sub

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim sTela As String
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoDocOriginal_Click

    If (GridNFiscal.Row = 0) Then gError 207040
    
    objNFiscal.lNumIntDoc = GridNFiscal.TextMatrix(GridNFiscal.Row, GRID_NUMINTNF_COL)
    
    lErro = CF("TipoDocInfo_Le_NomeTela_NFiscal1", objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 199923 Then gError 207041
    
    'nota fiscal nao encontrada
    If lErro <> SUCESSO Then gError 207042

    'Chama a Tela
    Call Chama_Tela(sTela, objNFiscal)

    Exit Sub

Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 207040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 207041

        Case 207042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207043)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub GridNFiscal_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNFiscal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscal, iAlterado)
    End If

End Sub

Public Sub GridNFiscal_EnterCell()
    Call Grid_Entrada_Celula(objGridNFiscal, iAlterado)
End Sub

Public Sub GridNFiscal_GotFocus()
    Call Grid_Recebe_Foco(objGridNFiscal)
End Sub

Public Sub GridNFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridNFiscal)
End Sub

Public Sub GridNFiscal_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNFiscal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscal, iAlterado)
    End If

End Sub

Public Sub GridNFiscal_LeaveCell()
    Call Saida_Celula(objGridNFiscal)
End Sub

Public Sub GridNFiscal_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridNFiscal)
End Sub

Public Sub GridNFiscal_RowColChange()
    Call Grid_RowColChange(objGridNFiscal)
End Sub

Public Sub GridNFiscal_Scroll()
    Call Grid_Scroll(objGridNFiscal)
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objNFeFedLote As ClassNFeFedLote
Dim colNFeFedLote As New Collection
Dim lLote As Long
Dim sNFSEEXE As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se o codigo foi preenchido
    If Len(Lote.Text) = 0 Then gError 207044
    
    lLote = StrParaLong(Lote.Text)

    'Loop de Validação dos dados do Grid
    For iLinha = 1 To objGridNFiscal.iLinhasExistentes

        If GridNFiscal.TextMatrix(iLinha, GRID_SELECIONA_COL) = S_MARCADO Then
        
            Set objNFeFedLote = New ClassNFeFedLote
            
            objNFeFedLote.lNumIntNF = StrParaLong(GridNFiscal.TextMatrix(iLinha, GRID_NUMINTNF_COL))
            objNFeFedLote.lLote = lLote
            
            colNFeFedLote.Add objNFeFedLote
            
        End If
        
    Next
    
    If colNFeFedLote.Count = 0 Then gError 207045

    lErro = CF("RPSWEBLote_Grava", giFilialEmpresa, colNFeFedLote)
    If lErro <> SUCESSO Then gError 207046

    lErro = CF("NFSE_Obter_EXE", giFilialEmpresa, sNFSEEXE)
    If lErro <> SUCESSO Then gError 201194

    lErro = WinExec(sNFSEEXE & " Envio " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & Lote.Text, SW_NORMAL)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Call Rotina_Aviso(vbOK, "AVISO_INICIO_ENVIONFE", Lote.Text)
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 207044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PREENCHIDO", gErr)
 
        Case 207045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_NOTA_FISCAL_SELECIONADA", gErr)

        Case 207046, 201194

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207047)

    End Select

    Exit Function

End Function

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Envio de Lote de Nota Fiscal de Serviço Eletrônica"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFSE"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
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

Public Sub BotaoLogLote_Click()
    
Dim objRPSWEBLoteLogView As New ClassRPSWEBLoteLogView
Dim colSelecao As Collection

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("RPSWEBLoteLogViewLista", colSelecao, objRPSWEBLoteLogView, objEventoLogLote)

    Exit Sub
End Sub

Public Sub BotaoLotes_Click()

Dim objRPSWEBLoteView As New ClassRPSWEBLoteView
Dim colSelecao As Collection

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("RPSWEBLoteViewLista", colSelecao, objRPSWEBLoteView, objEventoLotes)

    Exit Sub
End Sub

Public Sub BotaoRetCanc_Click()

Dim objRPSWEBRetCancView As New ClassRPSWEBRetCancView
Dim colSelecao As Collection

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("RPSWEBRetCancViewLista", colSelecao, objRPSWEBRetCancView, objEventoRetCanc)

    Exit Sub
End Sub

Public Sub BotaoRetConsulta_Click()

Dim objRPSWEBConsLote As New ClassRPSWEBConsLote
Dim colSelecao As Collection

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("RPSWEBConsLoteLista", colSelecao, objRPSWEBConsLote, objEventoRetConsulta)

    Exit Sub
End Sub

Public Sub BotaoRetNFe_Click()

Dim objRPSWEBRetEnvi As New ClassRPSWEBRetEnvi
Dim colSelecao As Collection

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("RPSWEBRetEnviLista", colSelecao, objRPSWEBRetEnvi, objEventoRetNFe)

    Exit Sub
End Sub

Public Sub BotaoStatusNFe_Click()

Dim objRPSWEBProtView As New ClassRPSWEBProtView
Dim colSelecao As Collection

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("RPSWEBProtViewLista", colSelecao, objRPSWEBProtView, objEventoStatusNFe)

    Exit Sub
End Sub



