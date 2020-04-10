VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVBEnvioEmailX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCobreBemX As CobreBemX.ContaCorrente

Public Function Finaliza()
    Set gobjCobreBemX = Nothing
End Function
     
Public Function Email_Enviar(ByVal sArqLic As String, ByVal sAg As String, ByVal sCta As String, ByVal sCed As String, ByVal sCedNome As String, ByVal sLogo As String, ByVal sIDCed As String, ByVal sSMPTHost As String, ByVal sFrom As String, ByVal sFromName As String, ByVal sUser As String, ByVal sPassword As String, ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sCc As String, ByVal sNumDoc As String, ByVal sNomeSac As String, ByVal sCPFSac As String, ByVal sEndSac As String, ByVal sEndBairroSac As String, ByVal sEndCidSac As String, ByVal sEndUFSac As String, ByVal sEndCepSac As String, ByVal sDataVenc As String, ByVal sVlr As String, ByVal sJuros As String, ByVal sMulta As String, ByVal sDesc As String, ByVal sOutAcr As String, ByVal sDem As String, ByVal sInstCx As String, ByVal sNossoNum As String, ByVal sImagens As String, ByVal sURLImgCB As String, sURLLogo As String) As Long

Dim objBoleto As Object
Dim objEmail As Object

On Error GoTo Erro_Email_Enviar

    Set gobjCobreBemX = New CobreBemX.ContaCorrente
    
    With gobjCobreBemX
        
        ' Início dos parâmetros obrigatórios da conta corrente
        .ArquivoLicenca = sArqLic
        .CodigoAgencia = sAg
        .NumeroContaCorrente = sCta
        .CodigoCedente = sCed
        .InicioNossoNumero = 1
        .FimNossoNumero = 999999
        .ProximoNossoNumero = 1
        ' Término dos parâmetros obrigatórios da conta corrente

        ' Início da configuração dos dados do Cedente para envio de boletos por email
        .PadroesBoleto.PadroesBoletoEmail.SMTP.Servidor = sSMPTHost ' Trocar para apontar para o seu servidor SMTP
        .PadroesBoleto.PadroesBoletoEmail.SMTP.Porta = 25
        .PadroesBoleto.PadroesBoletoEmail.SMTP.Usuario = sUser 'utilizar esta propriedade para acesso a servidores SMTP seguros
        .PadroesBoleto.PadroesBoletoEmail.SMTP.Senha = sPassword 'utilizar esta propriedade para acesso a servidores SMTP seguros
        .PadroesBoleto.PadroesBoletoEmail.URLImagensCodigoBarras = sURLImgCB
        .PadroesBoleto.PadroesBoletoEmail.URLLogotipo = sURLLogo
        .PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Assunto = sSubject
        .PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Endereco = sFrom
        .PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Nome = sFromName
        .PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = feeSMTPBoletoHTML
        .PadroesBoleto.PadroesBoletoEmail.PadroesEmail.CopiaFrom = False
        .PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Mensagem = sTexto
        .PadroesBoleto.PadroesBoletoEmail.PadroesEmail.SolicitaConfirmacaoRecebimento = True
        .PadroesBoleto.PadroesBoletoEmail.CaminhoAnexos = sAnexo
        ' Término da configuração dos dados do Cedente para envio de boletos por email

        ' Início da configuração dos dados do Cedente para impressão de boletos
        ' A próxima linha só é necessária se for solicitada a impressão de boletos e
        ' se desejar colocar o logotipo do Cedente na parte superior do boleto
        .PadroesBoleto.PadroesBoletoImpresso.ArquivoLogotipo = sLogo
        .PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = sImagens
        
        ' Utilize o parâmetro abaixo para efetuar ajustes na impressão do boleto subindo ou descendo o mesmo na folha de papel
        '  Os valores devem ser informados em milímetros e quanto maior o valor mais para baixo será iniciado o boleto
        '  Se este parâmetro não for passado será assumido o valor 15 que é o indicado para a maioria das impressoras Jato de Tinta }
        '.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 3
        ' Término da configuração dos dados do Cedente para impressão de boletos
     
        ' A próxima linha é utilizada para exibir um texto do lado direito do logotipo nos boletos impressos ou
        ' enviados por email
        .PadroesBoleto.IdentificacaoCedente = sIDCed
         
        .DocumentosCobranca.Clear
        
        Set objBoleto = .DocumentosCobranca.Add
        
    End With
        
    With objBoleto

        .NumeroDocumento = sNumDoc
        .NomeSacado = sNomeSac
        If Len(Trim(sCPFSac)) > 11 Then
            .CNPJSacado = sCPFSac
        Else
            .CPFSacado = sCPFSac
        End If
        .EnderecoSacado = sEndSac
        .BairroSacado = sEndBairroSac
        .CidadeSacado = sEndCidSac
        .EstadoSacado = sEndUFSac
        .CepSacado = sEndCepSac
        .DataDocumento = Format$(Date, "dd/mm/yyyy")
        .DataVencimento = sDataVenc
        .DataProcessamento = Format$(Date, "dd/mm/yyyy")
        .ValorDocumento = sVlr
'        .PercentualJurosDiaAtraso = sJuros
'        .PercentualMultaAtraso = sMulta
'        .PercentualDesconto = sDesc
'        .ValorOutrosAcrescimos = sOutAcr
        .PadroesBoleto.Demonstrativo = sDem
        .PadroesBoleto.InstrucoesCaixa = sInstCx
        
        Set objEmail = .EnderecosEmailSacado.Add
        objEmail.Nome = .NomeSacado
        objEmail.Endereco = sTo & ";" & sCc
        
        .ControleProcessamentoDocumento.EnviaEmail = scpExecutar
        
        .NossoNumero = sNossoNum
        .CalculaDacNossoNumero = True
     
    End With
     
    gobjCobreBemX.EnviaBoletosPorEmail
    
    If Len(Trim(gobjCobreBemX.UltimaMensagemErro)) > 0 Then MsgBox gobjCobreBemX.UltimaMensagemErro

    Set gobjCobreBemX = Nothing

    Exit Function

Erro_Email_Enviar:

    Email_Enviar = Err

    Select Case Err

        Case Else
            MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err
        
    End Select

    Exit Function

End Function

