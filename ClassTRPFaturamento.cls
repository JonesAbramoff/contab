VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPFaturamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCliente As Long
Private mvarcolFiliais As New Collection
Private mvarcolExcClientes As New Collection
Private mvarcolExcVouchers As New Collection
Private mvarcolTiposDoc As New Collection
Private mvarcolTiposFat As New Collection
Private mvarcolCondPagtos As New Collection
Private mvarcolInfoFiliais As New Collection
Private mvarcolInfoClientes As New Collection
Private mvarcolInfoFaturas As New Collection
Private mvarcolInfoVouchers As New Collection
Private mvarcolAportes As New Collection
Private mvarcolAportesCred As New Collection
Private mvarlModelo As Long
Private mvarsDiretorio As String
Private mvariEnviarPorEmail As Integer
Private mvariIndiceFaturaAtual As Integer
Private mvarcolFatFiliais As New Collection
Private mvarcolFatClientes As New Collection
Private mvarcolFatEnderecos As New Collection
Private mvarcolFatFiliaisCliente As New Collection
Private mvarsModelo As String
Private mvariFaturarIndividualmente As Integer

Public Property Let iFaturarIndividualmente(ByVal vData As Integer)
    mvariFaturarIndividualmente = vData
End Property

Public Property Get iFaturarIndividualmente() As Integer
    iFaturarIndividualmente = mvariFaturarIndividualmente
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Set colFatFiliaisCliente(ByVal vData As Collection)
    Set mvarcolFatFiliaisCliente = vData
End Property

Public Property Get colFatFiliaisCliente() As Collection
    Set colFatFiliaisCliente = mvarcolFatFiliaisCliente
End Property

Public Property Set colFatEnderecos(ByVal vData As Collection)
    Set mvarcolFatEnderecos = vData
End Property

Public Property Get colFatEnderecos() As Collection
    Set colFatEnderecos = mvarcolFatEnderecos
End Property

Public Property Set colFatClientes(ByVal vData As Collection)
    Set mvarcolFatClientes = vData
End Property

Public Property Get colFatClientes() As Collection
    Set colFatClientes = mvarcolFatClientes
End Property

Public Property Set colFatFiliais(ByVal vData As Collection)
    Set mvarcolFatFiliais = vData
End Property

Public Property Get colFatFiliais() As Collection
    Set colFatFiliais = mvarcolFatFiliais
End Property

Public Property Let iIndiceFaturaAtual(ByVal vData As Integer)
    mvariIndiceFaturaAtual = vData
End Property

Public Property Get iIndiceFaturaAtual() As Integer
    iIndiceFaturaAtual = mvariIndiceFaturaAtual
End Property

Public Property Let iEnviarPorEmail(ByVal vData As Integer)
    mvariEnviarPorEmail = vData
End Property

Public Property Get iEnviarPorEmail() As Integer
    iEnviarPorEmail = mvariEnviarPorEmail
End Property

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Public Property Let lModelo(ByVal vData As Long)
    mvarlModelo = vData
End Property

Public Property Get lModelo() As Long
    lModelo = mvarlModelo
End Property

Public Property Set colAportes(ByVal vData As Collection)
    Set mvarcolAportes = vData
End Property

Public Property Get colAportes() As Collection
    Set colAportes = mvarcolAportes
End Property

Public Property Set colAportesCred(ByVal vData As Collection)
    Set mvarcolAportesCred = vData
End Property

Public Property Get colAportesCred() As Collection
    Set colAportesCred = mvarcolAportesCred
End Property

Public Property Set colInfoVouchers(ByVal vData As Collection)
    Set mvarcolInfoVouchers = vData
End Property

Public Property Get colInfoVouchers() As Collection
    Set colInfoVouchers = mvarcolInfoVouchers
End Property

Public Property Set colInfoFaturas(ByVal vData As Collection)
    Set mvarcolInfoFaturas = vData
End Property

Public Property Get colInfoFaturas() As Collection
    Set colInfoFaturas = mvarcolInfoFaturas
End Property

Public Property Set colInfoClientes(ByVal vData As Collection)
    Set mvarcolInfoClientes = vData
End Property

Public Property Get colInfoClientes() As Collection
    Set colInfoClientes = mvarcolInfoClientes
End Property

Public Property Set colInfoFiliais(ByVal vData As Collection)
    Set mvarcolInfoFiliais = vData
End Property

Public Property Get colInfoFiliais() As Collection
    Set colInfoFiliais = mvarcolInfoFiliais
End Property

Public Property Set colCondPagtos(ByVal vData As Collection)
    Set mvarcolCondPagtos = vData
End Property

Public Property Get colCondPagtos() As Collection
    Set colCondPagtos = mvarcolCondPagtos
End Property

Public Property Set colTiposDoc(ByVal vData As Collection)
    Set mvarcolTiposDoc = vData
End Property

Public Property Get colTiposDoc() As Collection
    Set colTiposDoc = mvarcolTiposDoc
End Property

Public Property Set colTiposFat(ByVal vData As Collection)
    Set mvarcolTiposFat = vData
End Property

Public Property Get colTiposFat() As Collection
    Set colTiposFat = mvarcolTiposFat
End Property

Public Property Set colExcClientes(ByVal vData As Collection)
    Set mvarcolExcClientes = vData
End Property

Public Property Get colExcClientes() As Collection
    Set colExcClientes = mvarcolExcClientes
End Property

Public Property Set colExcVouchers(ByVal vData As Collection)
    Set mvarcolExcVouchers = vData
End Property

Public Property Get colExcVouchers() As Collection
    Set colExcVouchers = mvarcolExcVouchers
End Property

Public Property Set colFiliais(ByVal vData As Collection)
    Set mvarcolFiliais = vData
End Property

Public Property Get colFiliais() As Collection
    Set colFiliais = mvarcolFiliais
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional sValorTexto As String) As Long

Dim lErro As Long
Dim sMes As String
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objCliente As ClassCliente
Dim objFilialEmpresa As AdmFiliais
Dim objEnderecoCli As ClassEndereco
Dim objEnderecoEmp As ClassEndereco
Dim objFilialCliente As ClassFilialCliente
Dim sCgc As String
Dim objFATTRP As New ClassFATTRP
Dim sTabela As String
Dim lEspaco As Long

On Error GoTo Erro_Calcula_Mnemonico

    sValorTexto = ""
    
    Set objInfoFatura = colInfoFaturas.Item(iIndiceFaturaAtual)
    Set objCliente = colFatClientes.Item(objInfoFatura.iLinhaFatCliente)
    Set objFilialEmpresa = colFatFiliais.Item(objInfoFatura.iLinhaFatFilial)
    Set objEnderecoCli = colFatEnderecos.Item(objInfoFatura.iLinhaFatEnderecoCli)
    Set objEnderecoEmp = colFatEnderecos.Item(objInfoFatura.iLinhaFatEnderecoEmp)
    Set objFilialCliente = colFatFiliaisCliente.Item(objInfoFatura.iLinhaFatFilialCliente)

    Select Case UCase(objMnemonicoValor.sMnemonico)
    
        Case "BOTAO_GERAR_BOLETO"
        
            If objInfoFatura.dValor > TRP_VALOR_MINIMO_BOLETO Then
                sValorTexto = "<INPUT TYPE=""submit"" ACTION=""http://tripcard.dynalias.org:8084/scripts/cobrebemecommerce.exe"" VALUE=""Gerar Boleto"" METHOD=""GET"" NAME=""B1"">"
            Else
                sValorTexto = ""
            End If
            
            objMnemonicoValor.colValor.Add sValorTexto
                                                      
        Case "TIPO_DOC"
                                           
            If objInfoFatura.dValor < 0 Then
                objMnemonicoValor.colValor.Add "NOTA DE CRÉDITO"
                sValorTexto = "NOTA DE CRÉDITO"
            Else
                objMnemonicoValor.colValor.Add "FATURA"
                sValorTexto = "FATURA"
            End If
            
        Case "CNPJ_FILIALEMPRESA"
        
            sCgc = ""
            Select Case Len(Trim(objFilialEmpresa.sCgc))
                Case STRING_CPF
                    sCgc = Format(objFilialEmpresa.sCgc, "000\.000\.000-00; ; ; ")
        
                Case STRING_CGC
                    sCgc = Format(objFilialEmpresa.sCgc, "00\.000\.000\/0000-00; ; ; ")
            
            End Select

            objMnemonicoValor.colValor.Add sCgc
            sValorTexto = sCgc
        
        Case "BAIRRO_FILIALEMPRESA"
            objMnemonicoValor.colValor.Add objEnderecoEmp.sBairro
            sValorTexto = objEnderecoEmp.sBairro
        
        Case "ENDERECO_FILIALEMPRESA"
            objMnemonicoValor.colValor.Add objEnderecoEmp.sEndereco
            sValorTexto = objEnderecoEmp.sEndereco
        
        Case "CIDADE_FILIALEMPRESA"
            objMnemonicoValor.colValor.Add objEnderecoEmp.sCidade
            sValorTexto = objEnderecoEmp.sCidade
        
        Case "UF_FILIALEMPRESA"
            objMnemonicoValor.colValor.Add objEnderecoEmp.sSiglaEstado
            sValorTexto = objEnderecoEmp.sSiglaEstado
        
        Case "CEP_FILIALEMPRESA"
            objMnemonicoValor.colValor.Add objEnderecoEmp.sCEP
            sValorTexto = objEnderecoEmp.sCEP
        
        Case "TELEFONE_FILIALEMPRESA"
            objMnemonicoValor.colValor.Add objEnderecoEmp.sTelefone1
            sValorTexto = objEnderecoEmp.sTelefone1
        
        Case "FAX_FILIALEMPRESA"
            objMnemonicoValor.colValor.Add objEnderecoEmp.sFax
            sValorTexto = objEnderecoEmp.sFax
        
        Case "DATAEMISSAO"
            objMnemonicoValor.colValor.Add objInfoFatura.dtDataEmissao
            sValorTexto = Format(objInfoFatura.dtDataEmissao, "dd/mm/yyyy")
            
        Case "DATAVENCIMENTO"
            objMnemonicoValor.colValor.Add objInfoFatura.dtDataVencimento
            sValorTexto = Format(objInfoFatura.dtDataVencimento, "dd/mm/yyyy")
            
        Case "DATAEMISSAO_EST"
            Call MesNome(Month(objInfoFatura.dtDataEmissao), sMes)
            objMnemonicoValor.colValor.Add CStr(Day(objInfoFatura.dtDataEmissao)) & " de " & sMes & " de " & CStr(Year(objInfoFatura.dtDataEmissao))
            sValorTexto = CStr(Day(objInfoFatura.dtDataEmissao)) & " de " & sMes & " de " & CStr(Year(objInfoFatura.dtDataEmissao))
            
        Case "DATAVENCIMENTO_EST"
            Call MesNome(Month(objInfoFatura.dtDataVencimento), sMes)
            objMnemonicoValor.colValor.Add CStr(Day(objInfoFatura.dtDataVencimento)) & " de " & sMes & " de " & CStr(Year(objInfoFatura.dtDataVencimento))
            sValorTexto = CStr(Day(objInfoFatura.dtDataVencimento)) & " de " & sMes & " de " & CStr(Year(objInfoFatura.dtDataVencimento))
            
        Case "NOME_CLIENTE"
            objMnemonicoValor.colValor.Add objCliente.sNomeReduzido
            sValorTexto = objCliente.sNomeReduzido
            
        Case "RAZAO_CLIENTE"
            objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
            sValorTexto = objCliente.sRazaoSocial
            
        Case "CODIGO_CLIENTE"
            objMnemonicoValor.colValor.Add objCliente.lCodigo
            sValorTexto = CStr(objCliente.lCodigo)
            
        Case "CGC_CLIENTE"
            sCgc = ""
            Select Case Len(Trim(objFilialCliente.sCgc))
        
                Case STRING_CPF
                    sCgc = Format(objFilialCliente.sCgc, "000\.000\.000-00; ; ; ")
        
                Case STRING_CGC
                    sCgc = Format(objFilialCliente.sCgc, "00\.000\.000\/0000-00; ; ; ")
            
            End Select

            objMnemonicoValor.colValor.Add sCgc
            sValorTexto = sCgc
        
        Case "BAIRRO_CLIENTE"
            objMnemonicoValor.colValor.Add objEnderecoCli.sBairro
            sValorTexto = objEnderecoCli.sBairro
        
        Case "ENDERECO_CLIENTE"
            objMnemonicoValor.colValor.Add objEnderecoCli.sEndereco
            sValorTexto = objEnderecoCli.sEndereco
        
        Case "CIDADE_CLIENTE"
            objMnemonicoValor.colValor.Add objEnderecoCli.sCidade
            sValorTexto = objEnderecoCli.sCidade
        
        Case "UF_CLIENTE"
            objMnemonicoValor.colValor.Add objEnderecoCli.sSiglaEstado
            sValorTexto = objEnderecoCli.sSiglaEstado
        
        Case "CEP_CLIENTE"
            objMnemonicoValor.colValor.Add objEnderecoCli.sCEP
            sValorTexto = objEnderecoCli.sCEP
        
        Case "TELEFONE_CLIENTE"
            objMnemonicoValor.colValor.Add objEnderecoCli.sTelefone1
            sValorTexto = objEnderecoCli.sTelefone1
        
        Case "FAX_CLIENTE"
            objMnemonicoValor.colValor.Add objEnderecoCli.sFax
            sValorTexto = objEnderecoCli.sFax
            
        Case "NUMERO_FATURA"
            objMnemonicoValor.colValor.Add objInfoFatura.lNumeroFaturaReal
            sValorTexto = CStr(objInfoFatura.lNumeroFaturaReal)
                        
        Case "NOSSO_NUMERO"
            objMnemonicoValor.colValor.Add objInfoFatura.lNumeroFaturaReal
            sValorTexto = CStr(objInfoFatura.lNumeroFaturaReal)
                        
        Case "TOTAL_VOUCHER"
            objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValor + objInfoFatura.dValorCreditos + objInfoFatura.dValorDebitos - objInfoFatura.dValorDeducoes - objInfoFatura.dValorTarifa)
            sValorTexto = Format(Abs(objInfoFatura.dValor + objInfoFatura.dValorCreditos + objInfoFatura.dValorDebitos - objInfoFatura.dValorDeducoes - objInfoFatura.dValorTarifa), "STANDARD")
                        
        Case "VALOR_DOCUMENTO"
            objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValor)
            sValorTexto = Format(objInfoFatura.dValor, "STANDARD")
                        
        Case "TOTAL_VOUCHER_BRUTO"
            If objInfoFatura.iCartao = DESMARCADO Then
                objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValor + objInfoFatura.dValorCreditos + objInfoFatura.dValorDebitos + objInfoFatura.dValorComissao)
                sValorTexto = Format(Abs(objInfoFatura.dValor + objInfoFatura.dValorCreditos + objInfoFatura.dValorDebitos + objInfoFatura.dValorComissao), "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValor + objInfoFatura.dValorCreditos + objInfoFatura.dValorDebitos - objInfoFatura.dValorDeducoes - objInfoFatura.dValorTarifa)
                sValorTexto = Format(Abs(objInfoFatura.dValor + objInfoFatura.dValorCreditos + objInfoFatura.dValorDebitos - objInfoFatura.dValorDeducoes - objInfoFatura.dValorTarifa), "STANDARD")
            End If
                        
        Case "TOTAL_DEBITO"
                        
            objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValorDebitos)
            sValorTexto = Format(Abs(objInfoFatura.dValorDebitos), "STANDARD")
            
        Case "TOTAL_CREDITO"
        
            objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValorCreditos)
            sValorTexto = Format(Abs(objInfoFatura.dValorCreditos), "STANDARD")
                        
        Case "TOTAL_TARIFA"
            objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValorTarifa)
            sValorTexto = Format(Abs(objInfoFatura.dValorTarifa), "STANDARD")
                        
        Case "TOTAL_DEDUCOES"
            objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValorDeducoes)
            sValorTexto = Format(Abs(objInfoFatura.dValorDeducoes), "STANDARD")
            
        Case "ESPACO_PARA_FIM_FOLHA"
        
            lEspaco = 1
            
            'Acerta o tamanho do espaço em branco de acordo com as informações a serem exibidas
            If Abs(objInfoFatura.dValorDebitos) <= DELTA_VALORMONETARIO And Abs(objInfoFatura.dValorCreditos) <= DELTA_VALORMONETARIO Then
                lEspaco = lEspaco + 100
            End If
            If Abs(objInfoFatura.dValorTarifa) <= DELTA_VALORMONETARIO Then
                lEspaco = lEspaco + 100
            End If
            If Abs(objInfoFatura.dValorDeducoes) <= DELTA_VALORMONETARIO Then
                lEspaco = lEspaco + 100
            End If

                        
            objMnemonicoValor.colValor.Add lEspaco
            sValorTexto = CStr(lEspaco)
            
        Case "LINHA_CLIENTE_VOU"
        
            sValorTexto = ""
            
            If objInfoFatura.iCartao = MARCADO Or objInfoFatura.iTipoDoc = TRP_TIPODOC_FAT_OVER Then
        
                sValorTexto = sValorTexto & "<table border=""0"" width=""750"" cellspacing=""0"" cellpadding=""0""  align=center>"
                sValorTexto = sValorTexto & "<tr>"
                sValorTexto = sValorTexto & "<td width=""102"" class=""texto2"">Nome Cli Vou:</td>"
                sValorTexto = sValorTexto & "<td width=""488"" class=""texto"" >" & objInfoFatura.colInfoVouchers.Item(1).sNomeCliVou & "</td>"
                sValorTexto = sValorTexto & "<td width=""80""  class=""texto2"">Nº Cli Vou:</td>"
                sValorTexto = sValorTexto & "<td width=""80""  class=""texto"" >" & CStr(objInfoFatura.colInfoVouchers.Item(1).lCliVou) & "</td>"
                sValorTexto = sValorTexto & "</tr>"
                sValorTexto = sValorTexto & "</table>"
            
            End If
            
            objMnemonicoValor.colValor.Add sValorTexto
            
        Case "TOTAL_CABECALHO_LINHA"

            sValorTexto = ""
            
            If Abs(objInfoFatura.dValor - objInfoFatura.dValorCreditos - objInfoFatura.dValorDebitos + objInfoFatura.dValorDeducoes + objInfoFatura.dValorTarifa) > DELTA_VALORMONETARIO Then
            
                'sValorTexto = sValorTexto & "<table border=""1"" width=""750"" cellspacing=""0"" cellpadding=""0""  align=center>"
                sValorTexto = sValorTexto & "<tr>"
                sValorTexto = sValorTexto & "<td width=""467"" height=""100"" class=""texto"">&nbsp;Por serviços profissionais prestados - vouchers emitidos</td>"
                sValorTexto = sValorTexto & "<td width=""69""  height=""100"" class=""texto"" align=center>R$</td>"
                sValorTexto = sValorTexto & "<td height=""100"" class=""texto"" align=right>" & Format(Abs(objInfoFatura.dValor - objInfoFatura.dValorCreditos - objInfoFatura.dValorDebitos + objInfoFatura.dValorDeducoes + objInfoFatura.dValorTarifa), "STANDARD") & "</td>"
                sValorTexto = sValorTexto & "</tr>"
                'sValorTexto = sValorTexto & "</table>"
                
            End If
            
            If Abs(objInfoFatura.dValorCreditos) > DELTA_VALORMONETARIO Then
            
                'sValorTexto = sValorTexto & "<table border=""1"" width=""750"" cellspacing=""0"" cellpadding=""0""  align=center>"
                sValorTexto = sValorTexto & "<tr>"
                sValorTexto = sValorTexto & "<td width=""467"" height=""100"" class=""texto"">&nbsp;Créditos</td>"
                sValorTexto = sValorTexto & "<td width=""69""  height=""100"" class=""texto"" align=center>R$</td>"
                sValorTexto = sValorTexto & "<td height=""100"" class=""texto""  align=right>" & Format(Abs(objInfoFatura.dValorCreditos), "STANDARD") & "</td>"
                sValorTexto = sValorTexto & "</tr>"
                'sValorTexto = sValorTexto & "</table>"
                
            End If
            
            If Abs(objInfoFatura.dValorDebitos) > DELTA_VALORMONETARIO Then

                'sValorTexto = sValorTexto & "<table border=""1"" width=""750"" cellspacing=""0"" cellpadding=""0""  align=center>"
                sValorTexto = sValorTexto & "<tr>"
                sValorTexto = sValorTexto & "<td width=""467"" height=""100"" class=""texto"">&nbsp;Débitos</td>"
                sValorTexto = sValorTexto & "<td width=""69""  height=""100"" class=""texto"" align=center>R$</td>"
                sValorTexto = sValorTexto & "<td height=""100"" class=""texto"" align=right>" & Format(Abs(objInfoFatura.dValorDebitos), "STANDARD") & "</td>"
                sValorTexto = sValorTexto & "</tr>"
                'sValorTexto = sValorTexto & "</table>"
                
            End If
            
            If Abs(objInfoFatura.dValorTarifa) > DELTA_VALORMONETARIO Then
            
                'sValorTexto = sValorTexto & "<table border=""1"" width=""750"" cellspacing=""0"" cellpadding=""0""  align=center>"
                sValorTexto = sValorTexto & "<tr>"
                sValorTexto = sValorTexto & "<td width=""467"" height=""100"" class=""texto"">&nbsp;Tarifas</td>"
                sValorTexto = sValorTexto & "<td width=""69""  height=""100"" class=""texto"" align=center>R$</td>"
                sValorTexto = sValorTexto & "<td height=""100"" class=""texto"" align=right>-" & Format(Abs(objInfoFatura.dValorTarifa), "STANDARD") & "</td>"
                sValorTexto = sValorTexto & "</tr>"
                'sValorTexto = sValorTexto & "</table>"
                
            End If
            
            If Abs(objInfoFatura.dValorDeducoes) > DELTA_VALORMONETARIO Then
            
                'sValorTexto = sValorTexto & "<table border=""1"" width=""750"" cellspacing=""0"" cellpadding=""0""  align=center>"
                sValorTexto = sValorTexto & "<tr>"
                sValorTexto = sValorTexto & "<td width=""467"" height=""100"" class=""texto"">&nbsp;Deduções</td>"
                sValorTexto = sValorTexto & "<td width=""69""  height=""100"" class=""texto"" align=center>R$</td>"
                sValorTexto = sValorTexto & "<td height=""100"" class=""texto"" align=right>-" & Format(Abs(objInfoFatura.dValorDeducoes), "STANDARD") & "</td>"
                sValorTexto = sValorTexto & "</tr>"
                'sValorTexto = sValorTexto & "</table>"
                
            End If

            objMnemonicoValor.colValor.Add sValorTexto
                        
        Case "TOTAL_COMISSAO"
            objMnemonicoValor.colValor.Add Abs(objInfoFatura.dValorComissao)
            sValorTexto = Format(Abs(objInfoFatura.dValorComissao), "STANDARD")
                        
        Case "DIA_ATUAL"
            objMnemonicoValor.colValor.Add Day(gdtDataAtual)
            sValorTexto = CStr(Day(gdtDataAtual))
        
        Case "MES_NOME"
            Call MesNome(Month(gdtDataAtual), sMes)
            objMnemonicoValor.colValor.Add sMes
            sValorTexto = sMes
        
        Case "MES_ATUAL"
            objMnemonicoValor.colValor.Add Month(gdtDataAtual)
            sValorTexto = CStr(Month(gdtDataAtual))
        
        Case "ANO_ATUAL"
            objMnemonicoValor.colValor.Add Year(gdtDataAtual)
            sValorTexto = CStr(Year(gdtDataAtual))
        
        Case "DATA_ATUAL"
            objMnemonicoValor.colValor.Add gdtDataAtual
            sValorTexto = Format(gdtDataAtual, "dd/mm/yyyy")
            
        Case "CSID"
            objMnemonicoValor.colValor.Add "Bradesco 14520"
            sValorTexto = "Bradesco 14520"
            
        Case "IDENTIFIC_PAGINA"
            objMnemonicoValor.colValor.Add "PB"
            sValorTexto = "PB"
        
        Case "USUARIO_BOLETO"
            objMnemonicoValor.colValor.Add "73178342028482089"
            sValorTexto = "73178342028482089"
            
        'Falta desenvolver
        Case "LISTA_ITENS_FATURA_PROD"
        
            lErro = Gera_Lista_Itens_Fatura_Prod(objInfoFatura, sTabela)
            If lErro <> SUCESSO Then gError 192208
        
            objMnemonicoValor.colValor.Add sTabela
            sValorTexto = sTabela
            
        Case "LISTA_ITENS_FATURA"
        
            lErro = Gera_Lista_Itens_Fatura(objInfoFatura, sTabela)
            If lErro <> SUCESSO Then gError 192206
            
            objMnemonicoValor.colValor.Add sTabela
            sValorTexto = sTabela
            
        Case "LISTA_CREDITOS"
        
            lErro = Gera_Lista_Credito(objInfoFatura, sTabela)
            If lErro <> SUCESSO Then gError 192214
        
            objMnemonicoValor.colValor.Add sTabela
            sValorTexto = sTabela
            
        Case "LISTA_DEBITOS"
        
            lErro = Gera_Lista_Debito(objInfoFatura, sTabela)
            If lErro <> SUCESSO Then gError 192214
        
            objMnemonicoValor.colValor.Add sTabela
            sValorTexto = sTabela
            
'        Case "NOME_ARQUIVO"
'            objMnemonicoValor.colValor.Add gsEmpresaTRPHTML & String(6 - Len(CStr(objInfoFatura.lNumeroFaturaReal)), " ") & CStr(objInfoFatura.lNumeroFaturaReal)
'            sValorTexto = gsEmpresaTRPHTML & String(6 - Len(CStr(objInfoFatura.lNumeroFaturaReal)), " ") & CStr(objInfoFatura.lNumeroFaturaReal)
            
        'Não estão sendo preenchidos ....
        
        Case "MOEDA"
            sValorTexto = ""
        Case "QUANTIDADE"
            sValorTexto = ""
        Case "VALOR_DESCONTO"
            sValorTexto = ""
        Case "OUTROS_ACRESCIMOS"
            sValorTexto = ""
        Case "PREFIXO_NOSSO_NUMERO"
            sValorTexto = ""
        Case "ACEITE"
            sValorTexto = ""
        Case "SACADOR_AVALISTA"
            sValorTexto = ""
        Case "ESPECIE_DOCUMENTO"
            sValorTexto = ""
        Case "PERCENTUAL_MULTA"
            sValorTexto = ""
        Case "PERC_JUROS_DIA"
            sValorTexto = ""
        Case "PRAZOVENCIMENTO"
            sValorTexto = ""
        Case "VENCIMENTODIAUTIL"
            sValorTexto = ""
        Case "DEMONSTRATIVO"
            sValorTexto = ""
            
        Case Else
            gError 192195

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
    
    MsgBox objMnemonicoValor.sMnemonico

    Select Case gErr

        Case 192195
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case 192206, 192208, 192214

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192196)

    End Select

    Exit Function

End Function

Private Function Html_Cabec_Tab_ItemFat(ByVal sControle As String) As String

Dim sScript As String

    sScript = sScript & "<tr height:""750"">"
    sScript = sScript & "<td  class=""texto2""><p align=""center""><b><font size=""2"">CONTROLE: " & sControle & "<br></font><u><o:p></o:p></u></b></td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "<tr height:""750"">"
    sScript = sScript & "<td height=""25"" class=""texto2""><p align=center><u><b><font size=""2"">DESCRIÇÕES DOS DOCUMENTOS</font></b></u></p></td>"
    sScript = sScript & "</tr>"

    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center width=""39"" class=""texto3"">Doc. nº</td>"
    sScript = sScript & "<td align=center width=""47"" class=""texto3"">Prod.</td>"
    sScript = sScript & "<td align=center width=""242"" class=""texto3"">Passageiros</td>"
    sScript = sScript & "<td align=center width=""18"" class=""texto3"">Pax</td>"
    sScript = sScript & "<td align=center class=""texto3"">Emissão</td>"
    sScript = sScript & "<td align=center width=""38"" class=""texto3"" width=""44"">Moeda</td>"
    sScript = sScript & "<td align=center width=""55"" class=""texto3"" width=""72"">Valor</td>"
    sScript = sScript & "<td align=center width=""43"" class=""texto3"">Câmbio</td>"
    sScript = sScript & "<td align=center width=""66"" class=""texto3"">Valor R$</td>"
    sScript = sScript & "<td align=center width=""58"" class=""texto3"">Desconto</td>"
    sScript = sScript & "<td align=center width=""60"" class=""texto3"">Líquido</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    
    Html_Cabec_Tab_ItemFat = sScript

End Function

Private Function Html_Linha_Tab_ItemFat(ByVal sDoc As String, ByVal sProd As String, ByVal sTitular As String, ByVal sPax As String, ByVal sEmissao As String, ByVal sMoeda As String, ByVal sValor As String, ByVal sCambio As String, ByVal sValorR As String, ByVal sComissao As String, ByVal sLiquido As String) As String

Dim sScript As String

    sScript = "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "style=""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=center width=""39"" class=""texto0"">" & sDoc & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=center width=""47"" class=""texto0"">" & sProd & "</td>"
    sScript = sScript & "<td align=left width=""242"" class=""texto0"">" & sTitular & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=center class=""texto0"" width=""18"">" & sPax & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=center class=""texto0"">" & sEmissao & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=center class=""texto0"" width=""38"">" & sMoeda & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=right class=""texto0"" width=""55"">" & sValor & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=right class=""texto0"" width=""43"">" & sCambio & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=right class=""texto0"" width=""66"">" & sValorR & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=right class=""texto0"" width=""58"">" & sComissao & "</td>"
    sScript = sScript & "<td " & IIf(StrParaInt(sPax) > 1, "rowspan=" & sPax, "") & " align=right class=""texto0"" width=""60"">" & sLiquido & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & IIf(StrParaInt(sPax) > 1, "", "</table>")

    Html_Linha_Tab_ItemFat = sScript

End Function

Private Function Html_Linha_Tab_ItemFat_Pax(ByVal sTitular As String, ByVal bFechaTable As Boolean) As String

Dim sScript As String

    sScript = sScript & "<tr><td align=left width=""242"" class=""texto0"">" & sTitular & "</td></tr>"
    If bFechaTable Then sScript = sScript & "</table>"

    Html_Linha_Tab_ItemFat_Pax = sScript

End Function

Private Function Html_SubTotal_Tab_ItemFat(ByVal sBruto As String, ByVal sComissao As String, ByVal sLiquido As String) As String

Dim sScript As String
    
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td>&nbsp;</td>"
    sScript = sScript & "<td width=""604"" class=""texto5"">SUBTOTAL</td>"
    sScript = sScript & "<td width=""60""  class=""texto0"" align=right>" & sBruto & "</td>"
    sScript = sScript & "<td width=""60""  class=""texto0"" align=right>" & sComissao & "</td>"
    sScript = sScript & "<td width=""60""  class=""texto0"" align=right>" & sLiquido & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    
    Html_SubTotal_Tab_ItemFat = sScript

End Function

Private Function Html_Total_Tab_ItemFat(ByVal sBruto As String, ByVal sComissao As String, ByVal sLiquido As String) As String

Dim sScript As String

    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid; border-width: 0px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td>&nbsp;</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td>&nbsp;</td>"
    sScript = sScript & "<td width=""525"" class=""texto5"">TOTAL GERAL</td>"
    sScript = sScript & "<td width=""79""  class=""texto5"">R$</td>"
    sScript = sScript & "<td width=""60""  class=""texto5"" align=right>" & sBruto & "</td>"
    sScript = sScript & "<td width=""60""  class=""texto5"" align=right>" & sComissao & "</td>"
    sScript = sScript & "<td width=""60""  class=""texto5"" align=right>" & sLiquido & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    
    Html_Total_Tab_ItemFat = sScript

End Function

Private Function Gera_Lista_Itens_Fatura(ByVal objInfoFatura As ClassTRPFATInfoFatura, sScript As String) As Long

Dim lErro As Long
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim sScriptAux As String
Dim sControleAnt As String
Dim colVouchers As New Collection
Dim colVouchersOrdenados As New Collection
Dim colCampos As New Collection
Dim iIndice As Integer
Dim dValorBruto As Double
Dim dValorComi As Double
Dim dValorLiq As Double
Dim dValorTBruto As Double
Dim dValorTComi As Double
Dim dValorTLiq As Double
Dim sMoeda As String
Dim objVou As ClassTRPVouchers
Dim objPax As ClassTRPVouPassageiros
Dim lComando As Long
Dim iCount As Integer

On Error GoTo Erro_Gera_Lista_Itens_Fatura

    For Each objInfoVoucher In objInfoFatura.colInfoVouchers
        'If objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO And objInfoVoucher.iMarcado = MARCADO Then
        If objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO And objInfoVoucher.iMarcado = MARCADO And objInfoVoucher.dValor > 0 Then
            colVouchers.Add objInfoVoucher
        End If
    Next
    
    colCampos.Add "sControle"
    colCampos.Add "lNumVou"

    Call Ordena_Colecao(colVouchers, colVouchersOrdenados, colCampos)
    
    If colVouchersOrdenados.Count > 0 Then

        sControleAnt = "FORPRINT_3B728_XPWT"
    
        iIndice = 0
        dValorTBruto = 0
        dValorTComi = 0
        dValorTLiq = 0
        For Each objInfoVoucher In colVouchersOrdenados
    
            iIndice = iIndice + 1
                
            If objInfoVoucher.sControle <> sControleAnt Then
            
                sControleAnt = objInfoVoucher.sControle
            
                If iIndice <> 1 Then
                    sScriptAux = sScriptAux & Html_SubTotal_Tab_ItemFat(Format(dValorBruto, "STANDARD"), Format(dValorComi, "STANDARD"), Format(dValorLiq, "STANDARD"))
                
                    dValorBruto = 0
                    dValorComi = 0
                    dValorLiq = 0
                End If
            
                sScriptAux = sScriptAux & Html_Cabec_Tab_ItemFat(objInfoVoucher.sControle)
            End If
            
            If objInfoVoucher.iMoeda = MOEDA_REAL Then
                sMoeda = "BRL"
            ElseIf objInfoVoucher.iMoeda = MOEDA_DOLAR Then
                sMoeda = "USD"
            Else
                sMoeda = "BRL"
            End If
                           
            sScriptAux = sScriptAux & Html_Linha_Tab_ItemFat(CStr(objInfoVoucher.lNumVou), objInfoVoucher.sProduto, objInfoVoucher.sTitular, CStr(objInfoVoucher.iPax), Format(objInfoVoucher.dtDataEmissao, "dd/mm/yy"), sMoeda, Format(objInfoVoucher.dValorCambio, "STANDARD"), Format(objInfoVoucher.dCambio, "##0.00##"), Format(objInfoVoucher.dValorBruto, "STANDARD"), Format(objInfoVoucher.dValorComissao, "STANDARD"), Format(objInfoVoucher.dValor, "STANDARD"))

            Set objVou = New ClassTRPVouchers
            
            objVou.sSerie = objInfoVoucher.sSerie
            objVou.sTipVou = objInfoVoucher.sTipoVou
            objVou.lNumVou = objInfoVoucher.lNumVou
            
            Call CF("TRPVouPassageiros_Le", lComando, objVou)

            iCount = 0
            For Each objPax In objVou.colPassageiros
                iCount = iCount + 1
                If objPax.iTitular = DESMARCADO Then
                    sScriptAux = sScriptAux & Html_Linha_Tab_ItemFat_Pax(objPax.sNome, IIf(iCount = objInfoVoucher.iPax, True, False))
                End If
            Next
            
            dValorTBruto = dValorTBruto + objInfoVoucher.dValorBruto
            dValorTComi = dValorTComi + objInfoVoucher.dValorComissao
            dValorTLiq = dValorTLiq + objInfoVoucher.dValor
    
            dValorBruto = dValorBruto + objInfoVoucher.dValorBruto
            dValorComi = dValorComi + objInfoVoucher.dValorComissao
            dValorLiq = dValorLiq + objInfoVoucher.dValor
    
            If iIndice = colVouchersOrdenados.Count Then
                sScriptAux = sScriptAux & Html_SubTotal_Tab_ItemFat(Format(dValorBruto, "STANDARD"), Format(dValorComi, "STANDARD"), Format(dValorLiq, "STANDARD"))
            
                dValorBruto = 0
                dValorComi = 0
                dValorLiq = 0
            End If
                
        Next
        
        sScriptAux = sScriptAux & Html_Total_Tab_ItemFat(Format(dValorTBruto, "STANDARD"), Format(dValorTComi, "STANDARD"), Format(dValorTLiq, "STANDARD"))
    
    End If
    
    sScript = sScriptAux
            
    Gera_Lista_Itens_Fatura = SUCESSO

    Exit Function

Erro_Gera_Lista_Itens_Fatura:

    Gera_Lista_Itens_Fatura = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192205)

    End Select


    Exit Function

End Function

Private Function Html_Cabec_Tab_ItemFat_Prod() As String

Dim sScript As String

    sScript = sScript & "<br><br>"
    sScript = sScript & "<table border=""0"" width=""750"" cellspacing=""0"" cellpadding=""0"" align=center>"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto5""><u>RESUMO</u></td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    sScript = sScript & "<br>"
    sScript = sScript & "<table border=""0"" width=""750"" cellspacing=""0"" cellpadding=""0"" align=center style=""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto5"">VOUCHERS EMITIDOS</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    sScript = sScript & "<br>"
    
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"" style=""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td width=""190"">&nbsp;&nbsp;&nbsp; Produto</td>"
    sScript = sScript & "<td>QTD</td>"
    sScript = sScript & "<td width=""140"">Valor R$</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Cabec_Tab_ItemFat_Prod = sScript

End Function

Private Function Html_Linha_Tab_ItemFat_Prod(ByVal sProduto As String, ByVal sQtd As String, ByVal sValor As String) As String

Dim sScript As String

    sScript = "<table border=""0"" width=""750"" cellspacing=""0"" cellpadding=""0"" align=center>"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=left class=""texto0"" width=""191"">" & sProduto & "</td>"
    sScript = sScript & "<td align=left class=""texto0"">" & sQtd & "</td>"
    sScript = sScript & "<td align=right class=""texto0"" width=""142"">" & sValor & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Linha_Tab_ItemFat_Prod = sScript

End Function

Private Function Html_Total_Tab_ItemFat_Prod(ByVal sValor As String) As String

Dim sScript As String

    sScript = "<table align=center border=""0"" width=""750"" cellpadding=""0"" Style = ""border-collapse: collapse; border-style: solid; border-width: 0px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td>&nbsp;</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"" Style = ""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td width=""83"">&nbsp;</td>"
    sScript = sScript & "<td class=""texto5"">TOTAL</td>"
    sScript = sScript & "<td class=""texto5"" width=""33"">R$</td>"
    sScript = sScript & "<td width=""108"" class=""texto5"" align=right>" & sValor & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    sScript = sScript & "<br><br>"

    Html_Total_Tab_ItemFat_Prod = sScript

End Function

Private Function Gera_Lista_Itens_Fatura_Prod(ByVal objInfoFatura As ClassTRPFATInfoFatura, sScript As String) As Long

Dim lErro As Long
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim sScriptAux As String
Dim colVouchers As New Collection
Dim colVouchersOrdenados As New Collection
Dim colCampos As New Collection
Dim iIndice As Integer
Dim dValor As Double
Dim dValorT As Double
Dim sProdutoAnt As String
Dim iQtd As Integer
Dim sParte As String

On Error GoTo Erro_Gera_Lista_Itens_Fatura_Prod

    sParte = "01"

    For Each objInfoVoucher In objInfoFatura.colInfoVouchers
        If objInfoVoucher.sTipoDoc = TRP_TIPODOC_VOU_TEXTO And objInfoVoucher.iMarcado = MARCADO Then
        'If objInfoVoucher.iMarcado = MARCADO Then
            colVouchers.Add objInfoVoucher
        End If
    Next
    sParte = "02"
    
    colCampos.Add "sProduto"

    Call Ordena_Colecao(colVouchers, colVouchersOrdenados, colCampos)
    
    sParte = "03"
    
    If colVouchersOrdenados.Count > 0 Then
    
        sParte = "04"
        sScriptAux = sScriptAux & Html_Cabec_Tab_ItemFat_Prod()
        
        sParte = "05"

        iIndice = 0
        dValorT = 0
        For Each objInfoVoucher In colVouchersOrdenados
                
            iIndice = iIndice + 1
        
            sParte = "06 - Vou " & CStr(objInfoVoucher.lNumVou)
                
            If objInfoVoucher.sProduto <> sProdutoAnt Then
                
                sParte = "07 - Vou " & CStr(objInfoVoucher.lNumVou)
                
                If iIndice <> 1 Then sScriptAux = sScriptAux & Html_Linha_Tab_ItemFat_Prod(sProdutoAnt, CStr(iQtd), Format(dValor, "STANDARD"))
    
                sProdutoAnt = objInfoVoucher.sProduto
    
                dValor = 0
                iQtd = 0
            End If
    
            sParte = "08 - Vou " & CStr(objInfoVoucher.lNumVou)
    
            dValor = dValor + objInfoVoucher.dValor
            dValorT = dValorT + objInfoVoucher.dValor
            iQtd = iQtd + 1
                
        Next
        
        sParte = "09"
        sScriptAux = sScriptAux & Html_Linha_Tab_ItemFat_Prod(sProdutoAnt, CStr(iQtd), Format(dValor, "STANDARD"))
        sParte = "10"
        sScriptAux = sScriptAux & Html_Total_Tab_ItemFat_Prod(Format(dValorT, "STANDARD"))
       
    End If
      
    sScript = sScriptAux
            
    Gera_Lista_Itens_Fatura_Prod = SUCESSO

    Exit Function

Erro_Gera_Lista_Itens_Fatura_Prod:

    Gera_Lista_Itens_Fatura_Prod = gErr
    
    MsgBox "Gera_Lista_Itens_Fatura_Prod " & sParte

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192207)

    End Select


    Exit Function

End Function

Private Function Html_Cabec_Tab_Credito() As String

Dim sScript As String
    
    sScript = sScript & "<br>"
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"" Style = ""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto5""></td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto5"">CRÉDITOS</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Cabec_Tab_Credito = sScript

End Function

Private Function Html_Cabec_Tab_Debito() As String

Dim sScript As String

    sScript = sScript & "<br>"
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"" Style = ""border-collapse: collapse; border-style: solid; border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto5""></td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto5"">DÉBITOS</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Cabec_Tab_Debito = sScript

End Function

Private Function Html_Cabec_Tab_Credito_Aporte() As String

Dim sScript As String

    sScript = sScript & "<br>"
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid;"
    sScript = sScript & "border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td width=""190"" class=""texto3"">&nbsp;&nbsp;&nbsp; Tipo Doc</td>"
    sScript = sScript & "<td class=""texto3"">Nº da Parcela</td>"
    sScript = sScript & "<td class=""texto3"" width=""140"">Valor R$</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Cabec_Tab_Credito_Aporte = sScript

End Function

Private Function Html_Cabec_Tab_Credito_Outros(ByVal sControle As String) As String

Dim sScript As String

    sScript = sScript & "<br>"
    sScript = sScript & "<table border=""0"" width=""750"" cellspacing=""0"" cellpadding=""0"" align=center>"
    sScript = sScript & "<tr height:""750"">"
    sScript = sScript & "<td  class=""texto2""><p align=""center""><b><font size=""2"">CONTROLE: " & sControle & "<br></font><u><o:p></o:p></u></b></td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "<tr height:""750"">"
    sScript = sScript & "<td height=""25"" class=""texto2""><p align=center><u><b><font size=""2"">DESCRIÇÕES DOS DOCUMENTOS</font></b></u></p></td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    sScript = sScript & "<br>"
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid;"
    sScript = sScript & "border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto3"" width=""89"">Tipo Doc</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""60"">Doc nº</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""204"">Titular</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""35"">Pax</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""146"">Emissor</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""71"">Emissão</td>"
    sScript = sScript & "<td align=center width=""140"" class=""texto3"">Valor R$</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Cabec_Tab_Credito_Outros = sScript

End Function

Private Function Html_Linha_Tab_Credito(ByVal sTipoDoc As String, ByVal sParcela As String, ByVal sValor As String) As String

Dim sScript As String

    sScript = "<table border=""0"" width=""750"" cellspacing=""0"" cellpadding=""0"""
    sScript = sScript & "align=center>"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=left class=""texto3"" width=""191"">" & sTipoDoc & "</td>"
    sScript = sScript & "<td align=left class=""texto3"">" & sParcela & "</td>"
    sScript = sScript & "<td align=right class=""texto3"" width=""142"">" & sValor & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Linha_Tab_Credito = sScript

End Function

Private Function Html_Linha_Tab_Credito_Outros(ByVal sTipoDoc As String, ByVal sDoc As String, ByVal sCliente As String, ByVal sPax As String, ByVal sEmissor As String, ByVal sDataEmissao As String, ByVal sValor As String) As String

Dim sScript As String

    sScript = sScript & "<table border=""0"" width=""750"" cellspacing=""0"""
    sScript = sScript & "cellpadding=""0"" align=center>"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td align=center class=""texto3"" width=""89"">" & sTipoDoc & "</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""65"">" & sDoc & "</td>"
    sScript = sScript & "<td align=left class=""texto3"" width=""204"">" & sCliente & "</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""35"">" & sPax & "</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""146"">" & sEmissor & "</td>"
    sScript = sScript & "<td align=center class=""texto3"" width=""71"">" & sDataEmissao & "</td>"
    sScript = sScript & "<td align=right class=""texto3"" width=""140"">" & sValor & "</td>"
    sScript = sScript & "</tr></table>"

    Html_Linha_Tab_Credito_Outros = sScript

End Function

Private Function Html_SubTotal_Tab_Credito(ByVal sValor As String) As String

Dim sScript As String
    
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid;"
    sScript = sScript & "border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td>&nbsp;</td>"
    sScript = sScript & "<td width=""604"" class=""texto5"">SUBTOTAL</td>"
    sScript = sScript & "<td width=""59"" class=""texto0"" align=right>" & sValor & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_SubTotal_Tab_Credito = sScript

End Function

Private Function Html_Total_Tab_Credito(ByVal sValor As String) As String

Dim sScript As String

    sScript = "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid;"
    sScript = sScript & "border-width: 0; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td>&nbsp;"
    sScript = sScript & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"
    
    sScript = sScript & "<table align=center border=""0"" width=""750"" cellpadding=""0"""
    sScript = sScript & "Style = ""border-collapse: collapse; border-style: solid;"
    sScript = sScript & "border-width: 1px; padding: 0"" bgcolor=""#FFFFFF"">"
    sScript = sScript & "<tr>"
    sScript = sScript & "<td width=""83"">&nbsp;</td>"
    sScript = sScript & "<td class=""texto5"">TOTAL</td>"
    sScript = sScript & "<td class=""texto5"" width=""33"">R$</td>"
    sScript = sScript & "<td width=""108"" class=""texto5"" align=right>" & sValor & "</td>"
    sScript = sScript & "</tr>"
    sScript = sScript & "</table>"

    Html_Total_Tab_Credito = sScript

End Function

Private Function Gera_Lista_Credito(ByVal objInfoFatura As ClassTRPFATInfoFatura, sScript As String) As Long

Dim lErro As Long
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoVoucherAux As ClassTRPFATInfoVoucher
Dim objVouAporte As Object
Dim objVouAporteAux As Object
Dim objVouAporteAux2 As New ClassTRPVoucherAporte
Dim sScriptAux As String
Dim colVouchers As New Collection
Dim colVouchersAux As New Collection
Dim colVouchersOrdenados As New Collection
Dim colCampos As New Collection
Dim iIndice As Integer
Dim dValor As Double
Dim dValorT As Double
Dim iQtd As Integer
Dim dTotalCred As Double
Dim dSubTotalCred As Double
Dim dTotalAporte As Double
Dim dValorAporte As Double
Dim sTipoDocAnt As String
Dim objFatAporte As ClassTRPFatAporte
Dim bAchou As Boolean
Dim sControleAnt As String
Dim objFaturaAporte As ClassTRPFaturaAporte
Dim iIndiceAux As Integer

On Error GoTo Erro_Gera_Lista_Credito

    'Pega os documento negativos e os com aporte positivo
    For Each objInfoVoucher In objInfoFatura.colInfoVouchersAG
        If (objInfoVoucher.dValorAporte > 0 Or objInfoVoucher.dValor < 0) And objInfoVoucher.iMarcado = MARCADO Then
            colVouchers.Add objInfoVoucher
            'colVouchersAux.Add objInfoVoucher
        End If
    Next
    
'    'Agrupa comissões do mesmo voucher
'    For Each objInfoVoucherAux In colVouchersAux
'        If objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_CMCC_TEXTO Or objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Or objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_OVER_TEXTO Then
'            bAchou = False
'            For Each objInfoVoucher In colVouchers
'                'Se é o mesmo documento do mesmo voucher agrupa os valores
'                If objInfoVoucherAux.sTipoDoc = objInfoVoucher.sTipoDoc And objInfoVoucherAux.sTipoVou = objInfoVoucher.sTipoVou And objInfoVoucherAux.sSerie = objInfoVoucher.sSerie And objInfoVoucherAux.lNumVou = objInfoVoucher.lNumVou Then
'                    objInfoVoucher.dValor = objInfoVoucher.dValor + objInfoVoucherAux.dValor
'                    bAchou = True
'                    Exit For
'                End If
'            Next
'            If Not bAchou Then colVouchers.Add objInfoVoucherAux
'        Else
'            colVouchers.Add objInfoVoucherAux
'        End If
'    Next
    
    If objInfoFatura.colFatAporteCred.Count > 0 Then
    
        Set objInfoVoucher = New ClassTRPFATInfoVoucher
        
        For Each objFaturaAporte In objInfoFatura.colFatAporteCred
        
            Set objVouAporteAux2 = New ClassTRPVoucherAporte
            
            objVouAporteAux2.dValorAporte = objFaturaAporte.dValorAporte
            objVouAporteAux2.iParcelaAporte = objFaturaAporte.iParcelaAporte
            objVouAporteAux2.iNumParcelas = objFaturaAporte.iNumParcelas
            objVouAporteAux2.lCodigoAporte = objFaturaAporte.lCodigoAporte
            objVouAporteAux2.iPagtoParcial = objFaturaAporte.iPagtoParcial
            
            objInfoVoucher.colAportes.Add objVouAporteAux2
            
        Next
    
        colVouchers.Add objInfoVoucher
    
    End If
    
    Set objInfoFatura.colFatAporte = New Collection
        
    colCampos.Add "lCodigoAporte"
    colCampos.Add "sControle"
    colCampos.Add "sTipoDoc"

    Call Ordena_Colecao(colVouchers, colVouchersOrdenados, colCampos)
    
    If colVouchersOrdenados.Count > 0 Then
    
        sScriptAux = sScriptAux & Html_Cabec_Tab_Credito()

        sControleAnt = "FORPRINT_3B728_XPWT"
    
        'Cloca os créditos
        iIndice = 0
        dTotalCred = 0
        dSubTotalCred = 0
        iIndiceAux = 0
        For Each objInfoVoucher In colVouchersOrdenados
        
            iIndiceAux = iIndiceAux + 1
                               
            If objInfoVoucher.dValor < 0 Then
            
                iIndice = iIndice + 1
                
                If objInfoVoucher.sControle <> sControleAnt Then
                
                    sControleAnt = objInfoVoucher.sControle
                
                    If iIndice <> 1 Then
                        sScriptAux = sScriptAux & Html_SubTotal_Tab_Credito(Format(Abs(dSubTotalCred), "STANDARD"))
                    
                        dSubTotalCred = 0
                    End If
                    
                    sScriptAux = sScriptAux & Html_Cabec_Tab_Credito_Outros(objInfoVoucher.sControle)
                End If
                
                sScriptAux = sScriptAux & Html_Linha_Tab_Credito_Outros(objInfoVoucher.sTipoDoc, CStr(objInfoVoucher.lNumVou), objInfoVoucher.sTitular, CStr(objInfoVoucher.iPax), objInfoVoucher.sEmissor, Format(objInfoVoucher.dtDataEmissao, "dd/mm/yy"), Format(Abs(objInfoVoucher.dValor), "STANDARD"))
                
                dTotalCred = dTotalCred + objInfoVoucher.dValor
                dSubTotalCred = dSubTotalCred + objInfoVoucher.dValor
           
            End If
                            
            If iIndiceAux = colVouchersOrdenados.Count Then
                If dSubTotalCred <> 0 Then sScriptAux = sScriptAux & Html_SubTotal_Tab_Credito(Format(Abs(dSubTotalCred), "STANDARD"))
                dSubTotalCred = 0
            End If
            
        Next
        
        iIndice = 0
        For Each objInfoVoucher In colVouchersOrdenados
        
            For Each objVouAporte In objInfoVoucher.colAportes

                bAchou = False
                For Each objFatAporte In objInfoFatura.colFatAporte
                    If objFatAporte.lCodigo = objVouAporte.lCodigoAporte And objFatAporte.iProxParcela = objVouAporte.iParcelaAporte And objFatAporte.iNumParcelas = objVouAporte.iNumParcelas Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou And objVouAporte.lCodigoAporte <> 0 And objVouAporte.dValorAporte > DELTA_VALORMONETARIO Then
                   
                    Set objFatAporte = New ClassTRPFatAporte
                    
                    objFatAporte.lCodigo = objVouAporte.lCodigoAporte
                    objFatAporte.iProxParcela = objVouAporte.iParcelaAporte
                    objFatAporte.iNumParcelas = objVouAporte.iNumParcelas
                    
                    objInfoFatura.colFatAporte.Add objFatAporte
                   
                    'Verifica o valor do aporte
                    dValorAporte = 0
                    For Each objInfoVoucherAux In colVouchersOrdenados
                        For Each objVouAporteAux In objInfoVoucherAux.colAportes
                        If objVouAporteAux.dValorAporte > 0 And objVouAporteAux.lCodigoAporte = objVouAporte.lCodigoAporte And objVouAporteAux.iParcelaAporte = objVouAporte.iParcelaAporte And objVouAporteAux.iNumParcelas = objVouAporte.iNumParcelas Then
                                dValorAporte = dValorAporte + objVouAporteAux.dValorAporte
                            End If
                        Next
                    Next
                    dTotalAporte = dTotalAporte + dValorAporte
                        
                    iIndice = iIndice + 1
                    
                    If iIndice = 1 Then
                        sScriptAux = sScriptAux & Html_Cabec_Tab_Credito_Aporte()
                    End If
                    
                    sScriptAux = sScriptAux & Html_Linha_Tab_Credito("Aporte" & SEPARADOR & CStr(objVouAporte.lCodigoAporte), IIf(objVouAporte.iNumParcelas = 0, CStr(objVouAporte.iParcelaAporte), CStr(objVouAporte.iParcelaAporte) & "\" & CStr(objVouAporte.iNumParcelas)) & IIf(objVouAporte.iPagtoParcial = MARCADO, " - Pagto Parcial", ""), Format(dValorAporte, "STANDARD"))
        
                End If
                
            Next
            
        Next
        
        'Se passou pelo primeiro tem pelo menos um
        If iIndice <> 0 Then
            sScriptAux = sScriptAux & Html_SubTotal_Tab_Credito(Format(dTotalAporte, "STANDARD"))
        End If
        
        sScriptAux = sScriptAux & Html_Total_Tab_Credito(Format(Abs(dTotalCred - dTotalAporte), "STANDARD"))
    
    End If
    
    sScript = sScriptAux
            
    Gera_Lista_Credito = SUCESSO

    Exit Function

Erro_Gera_Lista_Credito:

    Gera_Lista_Credito = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192207)

    End Select

    Exit Function

End Function

Private Function Gera_Lista_Debito(ByVal objInfoFatura As ClassTRPFATInfoFatura, sScript As String) As Long

Dim lErro As Long
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoVoucherAux As ClassTRPFATInfoVoucher
Dim objVouAporte As Object
Dim objVouAporteAux As Object
Dim sScriptAux As String
Dim colVouchers As New Collection
Dim colVouchersAux As New Collection
Dim colVouchersOrdenados As New Collection
Dim colCampos As New Collection
Dim iIndice As Integer
Dim dValor As Double
Dim dValorT As Double
Dim iQtd As Integer
Dim dTotalCred As Double
Dim dSubTotalCred As Double
Dim dTotalAporte As Double
Dim dValorAporte As Double
Dim sTipoDocAnt As String
Dim objFatAporte As ClassTRPFatAporte
Dim bAchou As Boolean
Dim sControleAnt As String
Dim iIndiceAux As Integer

On Error GoTo Erro_Gera_Lista_Debito

    'Pega os documento negativos e os com aporte positivo
    For Each objInfoVoucher In objInfoFatura.colInfoVouchersAG
        If (objInfoVoucher.dValorAporte < 0 Or (objInfoVoucher.dValor > 0 And objInfoVoucher.sTipoDoc <> TRP_TIPODOC_VOU_TEXTO)) And objInfoVoucher.iMarcado = MARCADO Then
            colVouchers.Add objInfoVoucher
            'colVouchersAux.Add objInfoVoucher
        End If
    Next
    
'    'Agrupa comissões do mesmo voucher
'    For Each objInfoVoucherAux In colVouchersAux
'        If objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_CMC_TEXTO Or objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_CMCC_TEXTO Or objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_CMR_TEXTO Or objInfoVoucherAux.sTipoDoc = TRP_TIPODOC_OVER_TEXTO Then
'            bAchou = False
'            For Each objInfoVoucher In colVouchers
'                'Se é o mesmo documento do mesmo voucher agrupa os valores
'                If objInfoVoucherAux.sTipoDoc = objInfoVoucher.sTipoDoc And objInfoVoucherAux.sTipoVou = objInfoVoucher.sTipoVou And objInfoVoucherAux.sSerie = objInfoVoucher.sSerie And objInfoVoucherAux.lNumVou = objInfoVoucher.lNumVou Then
'                    objInfoVoucher.dValor = objInfoVoucher.dValor + objInfoVoucherAux.dValor
'                    bAchou = True
'                    Exit For
'                End If
'            Next
'            If Not bAchou Then colVouchers.Add objInfoVoucherAux
'        Else
'            colVouchers.Add objInfoVoucherAux
'        End If
'    Next
    
    Set objInfoFatura.colFatAporte = New Collection
        
    colCampos.Add "lCodigoAporte"
    colCampos.Add "sControle"
    colCampos.Add "sTipoDoc"

    Call Ordena_Colecao(colVouchers, colVouchersOrdenados, colCampos)
    
    If colVouchersOrdenados.Count > 0 Then

        sScriptAux = sScriptAux & Html_Cabec_Tab_Debito()
    
        sControleAnt = "FORPRINT_3B728_XPWT"
    
        'Cloca os créditos
        iIndice = 0
        dTotalCred = 0
        dSubTotalCred = 0
        iIndiceAux = 0
        For Each objInfoVoucher In colVouchersOrdenados
        
            iIndiceAux = iIndiceAux + 1
                               
            If objInfoVoucher.dValor > 0 And objInfoVoucher.sTipoDoc <> TRP_TIPODOC_VOU_TEXTO Then
            
                iIndice = iIndice + 1
                
                If objInfoVoucher.sControle <> sControleAnt Then

                    sControleAnt = objInfoVoucher.sControle
                
                    If iIndice <> 1 Then
                        sScriptAux = sScriptAux & Html_SubTotal_Tab_Credito(Format(dSubTotalCred, "STANDARD"))
                    
                        dSubTotalCred = 0
                    End If

                    sScriptAux = sScriptAux & Html_Cabec_Tab_Credito_Outros(objInfoVoucher.sControle)
                End If
                    
                sScriptAux = sScriptAux & Html_Linha_Tab_Credito_Outros(objInfoVoucher.sTipoDoc, CStr(objInfoVoucher.lNumVou), objInfoVoucher.sTitular, CStr(objInfoVoucher.iPax), objInfoVoucher.sEmissor, Format(objInfoVoucher.dtDataEmissao, "dd/mm/yy"), Format(objInfoVoucher.dValor, "STANDARD"))
                
                dTotalCred = dTotalCred + objInfoVoucher.dValor
                dSubTotalCred = dSubTotalCred + objInfoVoucher.dValor
                
            End If
            
            If iIndiceAux = colVouchersOrdenados.Count Then
                If dSubTotalCred <> 0 Then sScriptAux = sScriptAux & Html_SubTotal_Tab_Credito(Format(dSubTotalCred, "STANDARD"))
                dSubTotalCred = 0
            End If
                            
        Next
        
        iIndice = 0
        For Each objInfoVoucher In colVouchersOrdenados

            For Each objVouAporte In objInfoVoucher.colAportes

                bAchou = False
                For Each objFatAporte In objInfoFatura.colFatAporte
                    If objFatAporte.lCodigo = objVouAporte.lCodigoAporte Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou And objVouAporte.lCodigoAporte <> 0 And objVouAporte.dValorAporte < -DELTA_VALORMONETARIO Then

                    Set objFatAporte = New ClassTRPFatAporte
                    
                    objFatAporte.lCodigo = objVouAporte.lCodigoAporte
                    
                    objInfoFatura.colFatAporte.Add objFatAporte
                    
                    'Verifica o valor do aporte
                    dValorAporte = 0
                    For Each objInfoVoucherAux In colVouchersOrdenados
                        For Each objVouAporteAux In objInfoVoucherAux.colAportes
                            If objVouAporteAux.dValorAporte < 0 And objVouAporteAux.lCodigoAporte = objVouAporte.lCodigoAporte Then
                                dValorAporte = dValorAporte + objVouAporteAux.dValorAporte
                            End If
                        Next
                    Next
                    dTotalAporte = dTotalAporte + dValorAporte
                        
                    iIndice = iIndice + 1
                    
                    If iIndice = 1 Then
                        sScriptAux = sScriptAux & Html_Cabec_Tab_Credito_Aporte()
                    End If
                        
                    sScriptAux = sScriptAux & Html_Linha_Tab_Credito("Aporte" & SEPARADOR & CStr(objVouAporte.lCodigoAporte), CStr(objVouAporte.iParcelaAporte), Format(-dValorAporte, "STANDARD"))
        
                End If
                
            Next
            
        Next
        
        'Se passou pelo primeiro tem pelo menos um
        If iIndice <> 0 Then
            sScriptAux = sScriptAux & Html_SubTotal_Tab_Credito(Format(-dTotalAporte, "STANDARD"))
        End If
        
        sScriptAux = sScriptAux & Html_Total_Tab_Credito(Format(dTotalCred - dTotalAporte, "STANDARD"))
    
    End If
    
    sScript = sScriptAux
            
    Gera_Lista_Debito = SUCESSO

    Exit Function

Erro_Gera_Lista_Debito:

    Gera_Lista_Debito = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192207)

    End Select

    Exit Function

End Function
