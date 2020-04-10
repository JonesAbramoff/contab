Imports System
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports System.Xml.Schema

Public Class Form1
    Public Const NFE_AMBIENTE_HOMOLOGACAO As Integer = 2
    Public Const NFE_AMBIENTE_PRODUCAO As Integer = 1

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Interval = 1000
        Timer1.Start()


    End Sub



    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick


        Dim sEmpresa As String
        Dim lLote As Long
        Dim sOperacao As String

        Dim lNumIntNF As Long
        Dim sMotivo As String
        Dim iFilialEmpresa As Integer

        Dim objEnvioNFe As ClassEnvioNFe = New ClassEnvioNFe
        Dim objCancelaNFe As ClassCancelaNFe = New ClassCancelaNFe
        Dim objConsultaLoteNFe As ClassConsultaLoteNFe = New ClassConsultaLoteNFe


        Dim arguments As [String]() = Environment.GetCommandLineArgs()

        Try

            Timer1.Stop()

            'os valores abaixo sao setados para depuracao  
            'simulando a chamada pela aplicacao vb6.
            'De acordo com o tipo da operacao descomente e atribua os valores devidos,
            'sempre preencha sOperacao, sEmpresa e iFilialEmpresa
            'Para sOperacao "Envio" ou "Consulta" preencha o lote
            'Para sOperacao "Cancela" preencha o NumIntNF e o motivo

            'p/todas as operacoes
            sOperacao = "Consulta" '"Envio"  '"Consulta"   'ou Envio ou Cancela(NF)
            sEmpresa = 1
            iFilialEmpresa = 1

            ''p/envio ou consulta de lote
            lLote = 235

            'p/cancelamento
            'lNumIntNF = 942
            'sMotivo = "teste"

            'os valores abaixo vem da aplicacao normal em vb6
            'comente as linhas abaixo para depuracao
            'sOperacao = arguments(1)
            'sEmpresa = arguments(2)
            'iFilialEmpresa = CInt(arguments(3))

            If sOperacao = "Envio" Then


                lLote = CLng(arguments(4))

                Lote.Text = lLote

                objEnvioNFe.Envia_Lote_NFe(sEmpresa, lLote, iFilialEmpresa)

            ElseIf sOperacao = "Cancela" Then

                lNumIntNF = CLng(arguments(4))

                sMotivo = arguments(5)

                objCancelaNFe.Cancela_NFe(sEmpresa, lNumIntNF, sMotivo, iFilialEmpresa)


            ElseIf sOperacao = "Consulta" Then

                '    lLote = CLng(arguments(4))

                Lote.Text = lLote

                objConsultaLoteNFe.Consulta_Lote_NFe(sEmpresa, lLote, iFilialEmpresa)


            End If

        Catch ex As Exception
            Msg.Items.Add("Erro na execucao")
        End Try

    End Sub
End Class

'Public Class AssinaturaDigital

'    Public Function Assinar(ByVal XMLString As String, ByVal RefUri As String, ByVal X509Cert As X509Certificate2) As Integer
'        '     Entradas:
'        '         XMLString: string XML a ser assinada
'        '         RefUri   : Referência da URI a ser assinada (Ex. infNFe
'        '         X509Cert : certificado digital a ser utilizado na assinatura digital
'        ' 
'        '     Retornos:
'        '         Assinar : 0 - Assinatura realizada com sucesso
'        '                   1 - Erro: Problema ao acessar o certificado digital - %exceção%
'        '                   2 - Problemas no certificado digital
'        '                  3 - XML mal formado + exceção
'        '                   4 - A tag de assinatura %RefUri% inexiste
'        '                   5 - A tag de assinatura %RefUri% não é unica
'        '                   6 - Erro Ao assinar o documento - ID deve ser string %RefUri(Atributo)%
'        '                   7 - Erro: Ao assinar o documento - %exceção%
'        ' 
'        '        XMLStringAssinado : string XML assinada
'        ' 
'        '         XMLDocAssinado    : XMLDocument do XML assinado
'        '

'        Dim resultado As Integer = 0
'        msgResultado = "Assinatura realizada com sucesso"
'        Try
'            '   certificado para ser utilizado na assinatura
'            '
'            Dim _xnome As String = ""
'            If (Not X509Cert Is Nothing) Then
'                _xnome = X509Cert.Subject.ToString()
'            End If

'            Dim _X509Cert As X509Certificate2 = New X509Certificate2()
'            Dim store As X509Store = New X509Store("MY", StoreLocation.CurrentUser)

'            store.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)
'            Dim collection As X509Certificate2Collection = store.Certificates
'            '(X509Certificate2Collection)
'            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindBySubjectDistinguishedName, _xnome, False)
'            '(X509Certificate2Collection)
'            If (collection1.Count = 0) Then
'                resultado = 2
'                msgResultado = "Problemas no certificado digital"
'            Else
'                ' certificado ok
'                _X509Cert = collection1(0)
'                Dim x As String
'                x = _X509Cert.GetKeyAlgorithm().ToString()
'                ' Create a new XML document.
'                Dim doc As XmlDocument = New XmlDocument()

'                ' Format the document to ignore white spaces.
'                doc.PreserveWhitespace = False

'                ' Load the passed XML file using it's name.
'                Try
'                    doc.LoadXml(XMLString)

'                    ' Verifica se a tag a ser assinada existe é única
'                    Dim qtdeRefUri As Integer = doc.GetElementsByTagName(RefUri).Count

'                    If (qtdeRefUri = 0) Then
'                        '  a URI indicada não existe
'                        resultado = 4
'                        msgResultado = "A tag de assinatura " + RefUri.Trim() + " inexiste"
'                        ' Exsiste mais de uma tag a ser assinada
'                    Else

'                        If (qtdeRefUri > 1) Then
'                            ' existe mais de uma URI indicada
'                            resultado = 5
'                            msgResultado = "A tag de assinatura " + RefUri.Trim() + " não é unica"

'                            '//else if (_listaNum.IndexOf(doc.GetElementsByTagName(RefUri).Item(0).Attributes.ToString().Substring(1,1))>0)
'                            '//{
'                            '//    resultado = 6;
'                            '//    msgResultado = "Erro: Ao assinar o documento - ID deve ser string (" + doc.GetElementsByTagName(RefUri).Item(0).Attributes + ")";
'                            '//}
'                        Else
'                            Try

'                                ' Create a SignedXml object.
'                                Dim SignedXml As SignedXml = New SignedXml(doc)

'                                ' Add the key to the SignedXml document 



'                                SignedXml.SigningKey = _X509Cert.PrivateKey

'                                ' Create a reference to be signed
'                                Dim reference As Reference = New Reference()
'                                ' pega o uri que deve ser assinada
'                                Dim _Uri As XmlAttributeCollection = doc.GetElementsByTagName(RefUri).Item(0).Attributes
'                                Dim _atributo As XmlAttribute
'                                For Each _atributo In _Uri
'                                    If (_atributo.Name = "Id") Then
'                                        reference.Uri = "#" + _atributo.InnerText
'                                    End If
'                                Next

'                                ' Add an enveloped transformation to the reference.
'                                Dim env As XmlDsigEnvelopedSignatureTransform = New XmlDsigEnvelopedSignatureTransform()
'                                reference.AddTransform(env)

'                                Dim c14 As XmlDsigC14NTransform = New XmlDsigC14NTransform()
'                                reference.AddTransform(c14)

'                                ' Add the reference to the SignedXml object.
'                                SignedXml.AddReference(reference)

'                                '// Create a new KeyInfo object
'                                Dim keyInfo As KeyInfo = New KeyInfo()

'                                '// Load the certificate into a KeyInfoX509Data object
'                                '// and add it to the KeyInfo object.
'                                keyInfo.AddClause(New KeyInfoX509Data(_X509Cert))

'                                '// Add the KeyInfo object to the SignedXml object.
'                                SignedXml.KeyInfo = keyInfo

'                                SignedXml.ComputeSignature()

'                                '// Get the XML representation of the signature and save
'                                '// it to an XmlElement object.
'                                Dim xmlDigitalSignature As XmlElement = SignedXml.GetXml()

'                                '// Append the element to the XML document.
'                                doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, True))
'                                XMLDoc = New XmlDocument()
'                                XMLDoc.PreserveWhitespace = False
'                                XMLDoc = doc

'                            Catch caught As Exception
'                                resultado = 7
'                                msgResultado = "Erro: Ao assinar o documento - " + caught.Message
'                            End Try
'                        End If
'                    End If
'                Catch caught As Exception
'                    resultado = 3
'                    msgResultado = "Erro: XML mal formado - " + caught.Message
'                End Try
'            End If
'        Catch caught As Exception

'            resultado = 1
'            msgResultado = "Erro: Problema ao acessar o certificado digital" + caught.Message
'        End Try
'        Assinar = resultado

'    End Function
'    '//
'    '// mensagem de Retorno
'    '//
'    Private msgResultado As String
'    Private XMLDoc As XmlDocument

'    Public Function XMLDocAssinado() As XmlDocument
'        XMLDocAssinado = XMLDoc
'    End Function

'    Public Function XMLStringAssinado() As String
'        XMLStringAssinado = XMLDoc.OuterXml
'    End Function

'    Public Function mensagemResultado() As String
'        mensagemResultado = msgResultado
'    End Function

'End Class

'Public Class Certificado

'    Public Function BuscaNome(ByVal Nome As String) As X509Certificate2

'        Dim _X509Cert As X509Certificate2 = New X509Certificate2()
'        Try

'            Dim store As X509Store = New X509Store("MY", StoreLocation.CurrentUser)
'            store.Open(OpenFlags.OpenExistingOnly Or OpenFlags.IncludeArchived Or OpenFlags.ReadWrite)
'            Dim collection As X509Certificate2Collection = store.Certificates
'            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, False)
'            Dim collection2 As X509Certificate2Collection = collection.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, False)
'            If Nome = "" Then
'                Dim scollection As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(collection1, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
'                If (scollection.Count = 0) Then
'                    _X509Cert.Reset()
'                    Console.WriteLine("Nenhum certificado escolhido", "Atenção")
'                Else
'                    _X509Cert = scollection(0)
'                End If
'            Else
'                Dim scollection As X509Certificate2Collection = collection2.Find(X509FindType.FindBySubjectName, Nome, False)
'                If (scollection.Count = 0) Then
'                    Console.WriteLine("Nenhum certificado válido foi encontrado com o nome informado: " + Nome, "Atenção")
'                    _X509Cert.Reset()
'                Else
'                    _X509Cert = scollection(0)
'                End If
'            End If
'            store.Close()
'            BuscaNome = _X509Cert

'        Catch ex As SystemException
'            Console.WriteLine(ex.Message)
'            BuscaNome = _X509Cert
'        End Try
'    End Function

'    Public Function BuscaNroSerie(ByVal NroSerie As String) As X509Certificate2
'        Dim _X509Cert As X509Certificate2 = New X509Certificate2()
'        Try

'            Dim store As X509Store = New X509Store("My", StoreLocation.CurrentUser)
'            store.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)
'            Dim collection As X509Certificate2Collection = store.Certificates
'            Dim collection1 As X509Certificate2Collection = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, True)
'            Dim collection2 As X509Certificate2Collection = collection1.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, True)
'            If (NroSerie = "") Then
'                Dim scollection As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(collection2, "Certificados Digitais", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection)
'                If (scollection.Count = 0) Then
'                    _X509Cert.Reset()
'                    Console.WriteLine("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção")
'                Else
'                    _X509Cert = scollection(0)
'                End If
'            Else
'                Dim scollection As X509Certificate2Collection = collection2.Find(X509FindType.FindBySerialNumber, NroSerie, True)
'                If (scollection.Count = 0) Then
'                    _X509Cert.Reset()
'                    Console.WriteLine("Nenhum certificado válido foi encontrado com o número de série informado: " + NroSerie, "Atenção")
'                Else
'                    _X509Cert = scollection(0)
'                End If
'            End If
'            store.Close()
'            Return _X509Cert
'        Catch ex As System.Exception
'            Console.WriteLine(ex.Message)
'            Return _X509Cert
'        End Try

'    End Function


'End Class

'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe"), _
' System.Xml.Serialization.XmlRootAttribute("enviNFe", [Namespace]:="http://www.portalfiscal.inf.br/nfe", IsNullable:=False)> _
'Partial Public Class TEnviNFe

'    Private idLoteField As String

'    Private nFeField() As TNFe

'    Private versaoField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property idLote() As String
'        Get
'            Return Me.idLoteField
'        End Get
'        Set(ByVal value As String)
'            Me.idLoteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("NFe")> _
'    Public Property NFe() As TNFe()
'        Get
'            Return Me.nFeField
'        End Get
'        Set(ByVal value As TNFe())
'            Me.nFeField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="token")> _
'    Public Property versao() As String
'        Get
'            Return Me.versaoField
'        End Get
'        Set(ByVal value As String)
'            Me.versaoField = value
'        End Set
'    End Property
'End Class


'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe"), _
' System.Xml.Serialization.XmlRootAttribute("NFe", [Namespace]:="http://www.portalfiscal.inf.br/nfe", IsNullable:=False)> _
'Partial Public Class TNFe
'    Private infNFeField As TNFeInfNFe

'    '''<remarks/>
'    Public Property infNFe() As TNFeInfNFe

'        Get
'            Return Me.infNFeField
'        End Get
'        Set(ByVal value As TNFeInfNFe)
'            Me.infNFeField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFe

'    Private ideField As TNFeInfNFeIde

'    Private emitField As TNFeInfNFeEmit

'    Private avulsaField As TNFeInfNFeAvulsa

'    Private destField As TNFeInfNFeDest

'    Private retiradaField As TLocal

'    Private entregaField As TLocal

'    Private detField() As TNFeInfNFeDet

'    Private totalField As TNFeInfNFeTotal

'    Private transpField As TNFeInfNFeTransp

'    Private cobrField As TNFeInfNFeCobr

'    Private infAdicField As TNFeInfNFeInfAdic

'    Private exportaField As TNFeInfNFeExporta

'    Private compraField As TNFeInfNFeCompra

'    Private versaoField As String

'    Private idField As String

'    '''<remarks/>
'    Public Property ide() As TNFeInfNFeIde
'        Get
'            Return Me.ideField
'        End Get
'        Set(ByVal value As TNFeInfNFeIde)
'            Me.ideField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property emit() As TNFeInfNFeEmit
'        Get
'            Return Me.emitField
'        End Get
'        Set(ByVal value As TNFeInfNFeEmit)
'            Me.emitField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property avulsa() As TNFeInfNFeAvulsa
'        Get
'            Return Me.avulsaField
'        End Get
'        Set(ByVal value As TNFeInfNFeAvulsa)
'            Me.avulsaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dest() As TNFeInfNFeDest
'        Get
'            Return Me.destField
'        End Get
'        Set(ByVal value As TNFeInfNFeDest)
'            Me.destField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property retirada() As TLocal
'        Get
'            Return Me.retiradaField
'        End Get
'        Set(ByVal value As TLocal)
'            Me.retiradaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property entrega() As TLocal
'        Get
'            Return Me.entregaField
'        End Get
'        Set(ByVal value As TLocal)
'            Me.entregaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("det")> _
'    Public Property det() As TNFeInfNFeDet()
'        Get
'            Return Me.detField
'        End Get
'        Set(ByVal value As TNFeInfNFeDet())
'            Me.detField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property total() As TNFeInfNFeTotal
'        Get
'            Return Me.totalField
'        End Get
'        Set(ByVal value As TNFeInfNFeTotal)
'            Me.totalField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property transp() As TNFeInfNFeTransp
'        Get
'            Return Me.transpField
'        End Get
'        Set(ByVal value As TNFeInfNFeTransp)
'            Me.transpField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cobr() As TNFeInfNFeCobr
'        Get
'            Return Me.cobrField
'        End Get
'        Set(ByVal value As TNFeInfNFeCobr)
'            Me.cobrField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property infAdic() As TNFeInfNFeInfAdic
'        Get
'            Return Me.infAdicField
'        End Get
'        Set(ByVal value As TNFeInfNFeInfAdic)
'            Me.infAdicField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property exporta() As TNFeInfNFeExporta
'        Get
'            Return Me.exportaField
'        End Get
'        Set(ByVal value As TNFeInfNFeExporta)
'            Me.exportaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property compra() As TNFeInfNFeCompra
'        Get
'            Return Me.compraField
'        End Get
'        Set(ByVal value As TNFeInfNFeCompra)
'            Me.compraField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="token")> _
'    Public Property versao() As String
'        Get
'            Return Me.versaoField
'        End Get
'        Set(ByVal value As String)
'            Me.versaoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="ID")> _
'    Public Property Id() As String
'        Get
'            Return Me.idField
'        End Get
'        Set(ByVal value As String)
'            Me.idField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeIde

'    Private cUFField As TCodUfIBGE

'    Private cNFField As String

'    Private natOpField As String

'    Private indPagField As TNFeInfNFeIdeIndPag

'    Private modField As TMod

'    Private serieField As String

'    Private nNFField As String

'    Private dEmiField As String

'    Private dSaiEntField As String

'    Private tpNFField As TNFeInfNFeIdeTpNF

'    Private cMunFGField As String

'    Private nFrefField() As TNFeInfNFeIdeNFref

'    Private tpImpField As TNFeInfNFeIdeTpImp

'    Private tpEmisField As TNFeInfNFeIdeTpEmis

'    Private cDVField As String

'    Private tpAmbField As TAmb

'    Private finNFeField As TFinNFe

'    Private procEmiField As TProcEmi

'    Private verProcField As String

'    Public Sub New()
'        MyBase.New()
'        Me.modField = TMod.Item55
'    End Sub

'    '''<remarks/>
'    Public Property cUF() As TCodUfIBGE
'        Get
'            Return Me.cUFField
'        End Get
'        Set(ByVal value As TCodUfIBGE)
'            Me.cUFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cNF() As String
'        Get
'            Return Me.cNFField
'        End Get
'        Set(ByVal value As String)
'            Me.cNFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property natOp() As String
'        Get
'            Return Me.natOpField
'        End Get
'        Set(ByVal value As String)
'            Me.natOpField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property indPag() As TNFeInfNFeIdeIndPag
'        Get
'            Return Me.indPagField
'        End Get
'        Set(ByVal value As TNFeInfNFeIdeIndPag)
'            Me.indPagField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property [mod]() As TMod
'        Get
'            Return Me.modField
'        End Get
'        Set(ByVal value As TMod)
'            Me.modField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property serie() As String
'        Get
'            Return Me.serieField
'        End Get
'        Set(ByVal value As String)
'            Me.serieField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property nNF() As String
'        Get
'            Return Me.nNFField
'        End Get
'        Set(ByVal value As String)
'            Me.nNFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dEmi() As String
'        Get
'            Return Me.dEmiField
'        End Get
'        Set(ByVal value As String)
'            Me.dEmiField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dSaiEnt() As String
'        Get
'            Return Me.dSaiEntField
'        End Get
'        Set(ByVal value As String)
'            Me.dSaiEntField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property tpNF() As TNFeInfNFeIdeTpNF
'        Get
'            Return Me.tpNFField
'        End Get
'        Set(ByVal value As TNFeInfNFeIdeTpNF)
'            Me.tpNFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cMunFG() As String
'        Get
'            Return Me.cMunFGField
'        End Get
'        Set(ByVal value As String)
'            Me.cMunFGField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("NFref")> _
'    Public Property NFref() As TNFeInfNFeIdeNFref()
'        Get
'            Return Me.nFrefField
'        End Get
'        Set(ByVal value As TNFeInfNFeIdeNFref())
'            Me.nFrefField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property tpImp() As TNFeInfNFeIdeTpImp
'        Get
'            Return Me.tpImpField
'        End Get
'        Set(ByVal value As TNFeInfNFeIdeTpImp)
'            Me.tpImpField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property tpEmis() As TNFeInfNFeIdeTpEmis
'        Get
'            Return Me.tpEmisField
'        End Get
'        Set(ByVal value As TNFeInfNFeIdeTpEmis)
'            Me.tpEmisField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cDV() As String
'        Get
'            Return Me.cDVField
'        End Get
'        Set(ByVal value As String)
'            Me.cDVField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property tpAmb() As TAmb
'        Get
'            Return Me.tpAmbField
'        End Get
'        Set(ByVal value As TAmb)
'            Me.tpAmbField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property finNFe() As TFinNFe
'        Get
'            Return Me.finNFeField
'        End Get
'        Set(ByVal value As TFinNFe)
'            Me.finNFeField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property procEmi() As TProcEmi
'        Get
'            Return Me.procEmiField
'        End Get
'        Set(ByVal value As TProcEmi)
'            Me.procEmiField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property verProc() As String
'        Get
'            Return Me.verProcField
'        End Get
'        Set(ByVal value As String)
'            Me.verProcField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TCodUfIBGE

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("11")> _
'    Item11

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("12")> _
'    Item12

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("13")> _
'    Item13

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("14")> _
'    Item14

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("15")> _
'    Item15

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("16")> _
'    Item16

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("17")> _
'    Item17

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("21")> _
'    Item21

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("22")> _
'    Item22

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("23")> _
'    Item23

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("24")> _
'    Item24

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("25")> _
'    Item25

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("26")> _
'    Item26

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("27")> _
'    Item27

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("28")> _
'    Item28

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("29")> _
'    Item29

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("31")> _
'    Item31

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("32")> _
'    Item32

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("33")> _
'    Item33

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("35")> _
'    Item35

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("41")> _
'    Item41

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("42")> _
'    Item42

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("43")> _
'    Item43

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("50")> _
'    Item50

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("51")> _
'    Item51

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("52")> _
'    Item52

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("53")> _
'    Item53
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeIdeIndPag

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TMod

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("55")> _
'    Item55
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeIdeTpNF

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeIdeNFref

'    Private itemField As Object

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("refNF", GetType(TNFeInfNFeIdeNFrefRefNF)), _
'     System.Xml.Serialization.XmlElementAttribute("refNFe", GetType(String))> _
'    Public Property Item() As Object
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As Object)
'            Me.itemField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeIdeNFrefRefNF

'    Private cUFField As TCodUfIBGE

'    Private aAMMField As String

'    Private cNPJField As String

'    Private modField As String

'    Private serieField As String

'    Private nNFField As String

'    Public Sub New()
'        MyBase.New()
'        Me.modField = "01"
'    End Sub

'    '''<remarks/>
'    Public Property cUF() As TCodUfIBGE
'        Get
'            Return Me.cUFField
'        End Get
'        Set(ByVal value As TCodUfIBGE)
'            Me.cUFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property AAMM() As String
'        Get
'            Return Me.aAMMField
'        End Get
'        Set(ByVal value As String)
'            Me.aAMMField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CNPJ() As String
'        Get
'            Return Me.cNPJField
'        End Get
'        Set(ByVal value As String)
'            Me.cNPJField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property [mod]() As String
'        Get
'            Return Me.modField
'        End Get
'        Set(ByVal value As String)
'            Me.modField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property serie() As String
'        Get
'            Return Me.serieField
'        End Get
'        Set(ByVal value As String)
'            Me.serieField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property nNF() As String
'        Get
'            Return Me.nNFField
'        End Get
'        Set(ByVal value As String)
'            Me.nNFField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TVeiculo

'    Private placaField As String

'    Private ufField As TUf

'    Private rNTCField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property placa() As String
'        Get
'            Return Me.placaField
'        End Get
'        Set(ByVal value As String)
'            Me.placaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UF() As TUf
'        Get
'            Return Me.ufField
'        End Get
'        Set(ByVal value As TUf)
'            Me.ufField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property RNTC() As String
'        Get
'            Return Me.rNTCField
'        End Get
'        Set(ByVal value As String)
'            Me.rNTCField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TUf

'    '''<remarks/>
'    AC

'    '''<remarks/>
'    AL

'    '''<remarks/>
'    AM

'    '''<remarks/>
'    AP

'    '''<remarks/>
'    BA

'    '''<remarks/>
'    CE

'    '''<remarks/>
'    DF

'    '''<remarks/>
'    ES

'    '''<remarks/>
'    GO

'    '''<remarks/>
'    MA

'    '''<remarks/>
'    MG

'    '''<remarks/>
'    MS

'    '''<remarks/>
'    MT

'    '''<remarks/>
'    PA

'    '''<remarks/>
'    PB

'    '''<remarks/>
'    PE

'    '''<remarks/>
'    PI

'    '''<remarks/>
'    PR

'    '''<remarks/>
'    RJ

'    '''<remarks/>
'    RN

'    '''<remarks/>
'    RO

'    '''<remarks/>
'    RR

'    '''<remarks/>
'    RS

'    '''<remarks/>
'    SC

'    '''<remarks/>
'    SE

'    '''<remarks/>
'    SP

'    '''<remarks/>
'    [TO]

'    '''<remarks/>
'    EX
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TLocal

'    Private cNPJField As String

'    Private xLgrField As String

'    Private nroField As String

'    Private xCplField As String

'    Private xBairroField As String

'    Private cMunField As String

'    Private xMunField As String

'    Private ufField As TUf

'    '''<remarks/>
'    Public Property CNPJ() As String
'        Get
'            Return Me.cNPJField
'        End Get
'        Set(ByVal value As String)
'            Me.cNPJField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xLgr() As String
'        Get
'            Return Me.xLgrField
'        End Get
'        Set(ByVal value As String)
'            Me.xLgrField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property nro() As String
'        Get
'            Return Me.nroField
'        End Get
'        Set(ByVal value As String)
'            Me.nroField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xCpl() As String
'        Get
'            Return Me.xCplField
'        End Get
'        Set(ByVal value As String)
'            Me.xCplField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xBairro() As String
'        Get
'            Return Me.xBairroField
'        End Get
'        Set(ByVal value As String)
'            Me.xBairroField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cMun() As String
'        Get
'            Return Me.cMunField
'        End Get
'        Set(ByVal value As String)
'            Me.cMunField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xMun() As String
'        Get
'            Return Me.xMunField
'        End Get
'        Set(ByVal value As String)
'            Me.xMunField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UF() As TUf
'        Get
'            Return Me.ufField
'        End Get
'        Set(ByVal value As TUf)
'            Me.ufField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TEndereco

'    Private xLgrField As String

'    Private nroField As String

'    Private xCplField As String

'    Private xBairroField As String

'    Private cMunField As String

'    Private xMunField As String

'    Private ufField As TUf

'    Private cEPField As String

'    Private cPaisField As String

'    Private xPaisField As String

'    Private foneField As String

'    '''<remarks/>
'    Public Property xLgr() As String
'        Get
'            Return Me.xLgrField
'        End Get
'        Set(ByVal value As String)
'            Me.xLgrField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property nro() As String
'        Get
'            Return Me.nroField
'        End Get
'        Set(ByVal value As String)
'            Me.nroField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xCpl() As String
'        Get
'            Return Me.xCplField
'        End Get
'        Set(ByVal value As String)
'            Me.xCplField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xBairro() As String
'        Get
'            Return Me.xBairroField
'        End Get
'        Set(ByVal value As String)
'            Me.xBairroField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cMun() As String
'        Get
'            Return Me.cMunField
'        End Get
'        Set(ByVal value As String)
'            Me.cMunField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xMun() As String
'        Get
'            Return Me.xMunField
'        End Get
'        Set(ByVal value As String)
'            Me.xMunField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UF() As TUf
'        Get
'            Return Me.ufField
'        End Get
'        Set(ByVal value As TUf)
'            Me.ufField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CEP() As String
'        Get
'            Return Me.cEPField
'        End Get
'        Set(ByVal value As String)
'            Me.cEPField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cPais() As String
'        Get
'            Return Me.cPaisField
'        End Get
'        Set(ByVal value As String)
'            Me.cPaisField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xPais() As String
'        Get
'            Return Me.xPaisField
'        End Get
'        Set(ByVal value As String)
'            Me.xPaisField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property fone() As String
'        Get
'            Return Me.foneField
'        End Get
'        Set(ByVal value As String)
'            Me.foneField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TEndEmi

'    Private xLgrField As String

'    Private nroField As String

'    Private xCplField As String

'    Private xBairroField As String

'    Private cMunField As String

'    Private xMunField As String

'    Private ufField As TUf

'    Private cEPField As String

'    Private cPaisField As String

'    Private xPaisField As String

'    Private foneField As String

'    '''<remarks/>
'    Public Property xLgr() As String
'        Get
'            Return Me.xLgrField
'        End Get
'        Set(ByVal value As String)
'            Me.xLgrField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property nro() As String
'        Get
'            Return Me.nroField
'        End Get
'        Set(ByVal value As String)
'            Me.nroField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xCpl() As String
'        Get
'            Return Me.xCplField
'        End Get
'        Set(ByVal value As String)
'            Me.xCplField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xBairro() As String
'        Get
'            Return Me.xBairroField
'        End Get
'        Set(ByVal value As String)
'            Me.xBairroField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cMun() As String
'        Get
'            Return Me.cMunField
'        End Get
'        Set(ByVal value As String)
'            Me.cMunField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xMun() As String
'        Get
'            Return Me.xMunField
'        End Get
'        Set(ByVal value As String)
'            Me.xMunField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UF() As TUf
'        Get
'            Return Me.ufField
'        End Get
'        Set(ByVal value As TUf)
'            Me.ufField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CEP() As String
'        Get
'            Return Me.cEPField
'        End Get
'        Set(ByVal value As String)
'            Me.cEPField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cPais() As String
'        Get
'            Return Me.cPaisField
'        End Get
'        Set(ByVal value As String)
'            Me.cPaisField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xPais() As String
'        Get
'            Return Me.xPaisField
'        End Get
'        Set(ByVal value As String)
'            Me.xPaisField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property fone() As String
'        Get
'            Return Me.foneField
'        End Get
'        Set(ByVal value As String)
'            Me.foneField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeIdeTpImp

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeIdeTpEmis

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TAmb

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TFinNFe

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TProcEmi

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeEmit

'    Private itemField As String

'    Private itemElementNameField As ItemChoiceType

'    Private xNomeField As String

'    Private xFantField As String

'    Private enderEmitField As TEndEmi

'    Private ieField As String

'    Private iESTField As String

'    Private imField As String

'    Private cNAEField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("CNPJ", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("CPF", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemElementName")> _
'    Public Property Item() As String
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As String)
'            Me.itemField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemElementName() As ItemChoiceType
'        Get
'            Return Me.itemElementNameField
'        End Get
'        Set(ByVal value As ItemChoiceType)
'            Me.itemElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xNome() As String
'        Get
'            Return Me.xNomeField
'        End Get
'        Set(ByVal value As String)
'            Me.xNomeField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xFant() As String
'        Get
'            Return Me.xFantField
'        End Get
'        Set(ByVal value As String)
'            Me.xFantField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property enderEmit() As TEndEmi
'        Get
'            Return Me.enderEmitField
'        End Get
'        Set(ByVal value As TEndEmi)
'            Me.enderEmitField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property IE() As String
'        Get
'            Return Me.ieField
'        End Get
'        Set(ByVal value As String)
'            Me.ieField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property IEST() As String
'        Get
'            Return Me.iESTField
'        End Get
'        Set(ByVal value As String)
'            Me.iESTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property IM() As String
'        Get
'            Return Me.imField
'        End Get
'        Set(ByVal value As String)
'            Me.imField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CNAE() As String
'        Get
'            Return Me.cNAEField
'        End Get
'        Set(ByVal value As String)
'            Me.cNAEField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemChoiceType

'    '''<remarks/>
'    CNPJ

'    '''<remarks/>
'    CPF
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeAvulsa

'    Private cNPJField As String

'    Private xOrgaoField As String

'    Private matrField As String

'    Private xAgenteField As String

'    Private foneField As String

'    Private ufField As TUf

'    Private nDARField As String

'    Private dEmiField As String

'    Private vDARField As String

'    Private repEmiField As String

'    Private dPagField As String

'    '''<remarks/>
'    Public Property CNPJ() As String
'        Get
'            Return Me.cNPJField
'        End Get
'        Set(ByVal value As String)
'            Me.cNPJField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xOrgao() As String
'        Get
'            Return Me.xOrgaoField
'        End Get
'        Set(ByVal value As String)
'            Me.xOrgaoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property matr() As String
'        Get
'            Return Me.matrField
'        End Get
'        Set(ByVal value As String)
'            Me.matrField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xAgente() As String
'        Get
'            Return Me.xAgenteField
'        End Get
'        Set(ByVal value As String)
'            Me.xAgenteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property fone() As String
'        Get
'            Return Me.foneField
'        End Get
'        Set(ByVal value As String)
'            Me.foneField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UF() As TUf
'        Get
'            Return Me.ufField
'        End Get
'        Set(ByVal value As TUf)
'            Me.ufField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nDAR() As String
'        Get
'            Return Me.nDARField
'        End Get
'        Set(ByVal value As String)
'            Me.nDARField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dEmi() As String
'        Get
'            Return Me.dEmiField
'        End Get
'        Set(ByVal value As String)
'            Me.dEmiField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vDAR() As String
'        Get
'            Return Me.vDARField
'        End Get
'        Set(ByVal value As String)
'            Me.vDARField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property repEmi() As String
'        Get
'            Return Me.repEmiField
'        End Get
'        Set(ByVal value As String)
'            Me.repEmiField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dPag() As String
'        Get
'            Return Me.dPagField
'        End Get
'        Set(ByVal value As String)
'            Me.dPagField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDest

'    Private itemField As String

'    Private itemElementNameField As ItemChoiceType1

'    Private xNomeField As String

'    Private enderDestField As TEndereco

'    Private ieField As String

'    Private iSUFField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("CNPJ", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("CPF", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemElementName")> _
'    Public Property Item() As String
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As String)
'            Me.itemField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemElementName() As ItemChoiceType1
'        Get
'            Return Me.itemElementNameField
'        End Get
'        Set(ByVal value As ItemChoiceType1)
'            Me.itemElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xNome() As String
'        Get
'            Return Me.xNomeField
'        End Get
'        Set(ByVal value As String)
'            Me.xNomeField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property enderDest() As TEndereco
'        Get
'            Return Me.enderDestField
'        End Get
'        Set(ByVal value As TEndereco)
'            Me.enderDestField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property IE() As String
'        Get
'            Return Me.ieField
'        End Get
'        Set(ByVal value As String)
'            Me.ieField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property ISUF() As String
'        Get
'            Return Me.iSUFField
'        End Get
'        Set(ByVal value As String)
'            Me.iSUFField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemChoiceType1

'    '''<remarks/>
'    CNPJ

'    '''<remarks/>
'    CPF
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDet

'    Private prodField As TNFeInfNFeDetProd

'    Private impostoField As TNFeInfNFeDetImposto

'    Private infAdProdField As String

'    Private nItemField As String

'    '''<remarks/>
'    Public Property prod() As TNFeInfNFeDetProd
'        Get
'            Return Me.prodField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProd)
'            Me.prodField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property imposto() As TNFeInfNFeDetImposto
'        Get
'            Return Me.impostoField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImposto)
'            Me.impostoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property infAdProd() As String
'        Get
'            Return Me.infAdProdField
'        End Get
'        Set(ByVal value As String)
'            Me.infAdProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="token")> _
'    Public Property nItem() As String
'        Get
'            Return Me.nItemField
'        End Get
'        Set(ByVal value As String)
'            Me.nItemField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProd

'    Private cProdField As String

'    Private cEANField As String

'    Private xProdField As String

'    Private nCMField As String

'    Private eXTIPIField As String

'    Private generoField As String

'    Private cFOPField As String

'    Private uComField As String

'    Private qComField As String

'    Private vUnComField As String

'    Private vProdField As String

'    Private cEANTribField As String

'    Private uTribField As String

'    Private qTribField As String

'    Private vUnTribField As String

'    Private vFreteField As String

'    Private vSegField As String

'    Private vDescField As String

'    Private diField() As TNFeInfNFeDetProdDI

'    Private itemsField() As Object

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cProd() As String
'        Get
'            Return Me.cProdField
'        End Get
'        Set(ByVal value As String)
'            Me.cProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cEAN() As String
'        Get
'            Return Me.cEANField
'        End Get
'        Set(ByVal value As String)
'            Me.cEANField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xProd() As String
'        Get
'            Return Me.xProdField
'        End Get
'        Set(ByVal value As String)
'            Me.xProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property NCM() As String
'        Get
'            Return Me.nCMField
'        End Get
'        Set(ByVal value As String)
'            Me.nCMField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property EXTIPI() As String
'        Get
'            Return Me.eXTIPIField
'        End Get
'        Set(ByVal value As String)
'            Me.eXTIPIField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property genero() As String
'        Get
'            Return Me.generoField
'        End Get
'        Set(ByVal value As String)
'            Me.generoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CFOP() As String
'        Get
'            Return Me.cFOPField
'        End Get
'        Set(ByVal value As String)
'            Me.cFOPField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property uCom() As String
'        Get
'            Return Me.uComField
'        End Get
'        Set(ByVal value As String)
'            Me.uComField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property qCom() As String
'        Get
'            Return Me.qComField
'        End Get
'        Set(ByVal value As String)
'            Me.qComField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vUnCom() As String
'        Get
'            Return Me.vUnComField
'        End Get
'        Set(ByVal value As String)
'            Me.vUnComField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vProd() As String
'        Get
'            Return Me.vProdField
'        End Get
'        Set(ByVal value As String)
'            Me.vProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cEANTrib() As String
'        Get
'            Return Me.cEANTribField
'        End Get
'        Set(ByVal value As String)
'            Me.cEANTribField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property uTrib() As String
'        Get
'            Return Me.uTribField
'        End Get
'        Set(ByVal value As String)
'            Me.uTribField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property qTrib() As String
'        Get
'            Return Me.qTribField
'        End Get
'        Set(ByVal value As String)
'            Me.qTribField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vUnTrib() As String
'        Get
'            Return Me.vUnTribField
'        End Get
'        Set(ByVal value As String)
'            Me.vUnTribField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vFrete() As String
'        Get
'            Return Me.vFreteField
'        End Get
'        Set(ByVal value As String)
'            Me.vFreteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vSeg() As String
'        Get
'            Return Me.vSegField
'        End Get
'        Set(ByVal value As String)
'            Me.vSegField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vDesc() As String
'        Get
'            Return Me.vDescField
'        End Get
'        Set(ByVal value As String)
'            Me.vDescField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("DI")> _
'    Public Property DI() As TNFeInfNFeDetProdDI()
'        Get
'            Return Me.diField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdDI())
'            Me.diField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("arma", GetType(TNFeInfNFeDetProdArma)), _
'     System.Xml.Serialization.XmlElementAttribute("comb", GetType(TNFeInfNFeDetProdComb)), _
'     System.Xml.Serialization.XmlElementAttribute("med", GetType(TNFeInfNFeDetProdMed)), _
'     System.Xml.Serialization.XmlElementAttribute("veicProd", GetType(TNFeInfNFeDetProdVeicProd))> _
'    Public Property Items() As Object()
'        Get
'            Return Me.itemsField
'        End Get
'        Set(ByVal value As Object())
'            Me.itemsField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdDI

'    Private nDIField As String

'    Private dDIField As String

'    Private xLocDesembField As String

'    Private uFDesembField As TUf

'    Private dDesembField As String

'    Private cExportadorField As String

'    Private adiField() As TNFeInfNFeDetProdDIAdi

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nDI() As String
'        Get
'            Return Me.nDIField
'        End Get
'        Set(ByVal value As String)
'            Me.nDIField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dDI() As String
'        Get
'            Return Me.dDIField
'        End Get
'        Set(ByVal value As String)
'            Me.dDIField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xLocDesemb() As String
'        Get
'            Return Me.xLocDesembField
'        End Get
'        Set(ByVal value As String)
'            Me.xLocDesembField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UFDesemb() As TUf
'        Get
'            Return Me.uFDesembField
'        End Get
'        Set(ByVal value As TUf)
'            Me.uFDesembField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dDesemb() As String
'        Get
'            Return Me.dDesembField
'        End Get
'        Set(ByVal value As String)
'            Me.dDesembField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cExportador() As String
'        Get
'            Return Me.cExportadorField
'        End Get
'        Set(ByVal value As String)
'            Me.cExportadorField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("adi")> _
'    Public Property adi() As TNFeInfNFeDetProdDIAdi()
'        Get
'            Return Me.adiField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdDIAdi())
'            Me.adiField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdDIAdi

'    Private nAdicaoField As String

'    Private nSeqAdicField As String

'    Private cFabricanteField As String

'    Private vDescDIField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nAdicao() As String
'        Get
'            Return Me.nAdicaoField
'        End Get
'        Set(ByVal value As String)
'            Me.nAdicaoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nSeqAdic() As String
'        Get
'            Return Me.nSeqAdicField
'        End Get
'        Set(ByVal value As String)
'            Me.nSeqAdicField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cFabricante() As String
'        Get
'            Return Me.cFabricanteField
'        End Get
'        Set(ByVal value As String)
'            Me.cFabricanteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vDescDI() As String
'        Get
'            Return Me.vDescDIField
'        End Get
'        Set(ByVal value As String)
'            Me.vDescDIField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdArma

'    Private tpArmaField As TNFeInfNFeDetProdArmaTpArma

'    Private nSerieField As String

'    Private nCanoField As String

'    Private descrField As String

'    '''<remarks/>
'    Public Property tpArma() As TNFeInfNFeDetProdArmaTpArma
'        Get
'            Return Me.tpArmaField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdArmaTpArma)
'            Me.tpArmaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nSerie() As String
'        Get
'            Return Me.nSerieField
'        End Get
'        Set(ByVal value As String)
'            Me.nSerieField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nCano() As String
'        Get
'            Return Me.nCanoField
'        End Get
'        Set(ByVal value As String)
'            Me.nCanoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property descr() As String
'        Get
'            Return Me.descrField
'        End Get
'        Set(ByVal value As String)
'            Me.descrField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetProdArmaTpArma

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdComb

'    Private cProdANPField As String

'    Private cODIFField As String

'    Private qTempField As String

'    Private cIDEField As TNFeInfNFeDetProdCombCIDE

'    Private iCMSCombField As TNFeInfNFeDetProdCombICMSComb

'    Private iCMSInterField As TNFeInfNFeDetProdCombICMSInter

'    Private iCMSConsField As TNFeInfNFeDetProdCombICMSCons

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cProdANP() As String
'        Get
'            Return Me.cProdANPField
'        End Get
'        Set(ByVal value As String)
'            Me.cProdANPField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CODIF() As String
'        Get
'            Return Me.cODIFField
'        End Get
'        Set(ByVal value As String)
'            Me.cODIFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property qTemp() As String
'        Get
'            Return Me.qTempField
'        End Get
'        Set(ByVal value As String)
'            Me.qTempField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CIDE() As TNFeInfNFeDetProdCombCIDE
'        Get
'            Return Me.cIDEField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdCombCIDE)
'            Me.cIDEField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property ICMSComb() As TNFeInfNFeDetProdCombICMSComb
'        Get
'            Return Me.iCMSCombField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdCombICMSComb)
'            Me.iCMSCombField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property ICMSInter() As TNFeInfNFeDetProdCombICMSInter
'        Get
'            Return Me.iCMSInterField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdCombICMSInter)
'            Me.iCMSInterField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property ICMSCons() As TNFeInfNFeDetProdCombICMSCons
'        Get
'            Return Me.iCMSConsField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdCombICMSCons)
'            Me.iCMSConsField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdCombCIDE

'    Private qBCProdField As String

'    Private vAliqProdField As String

'    Private vCIDEField As String

'    '''<remarks/>
'    Public Property qBCProd() As String
'        Get
'            Return Me.qBCProdField
'        End Get
'        Set(ByVal value As String)
'            Me.qBCProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vAliqProd() As String
'        Get
'            Return Me.vAliqProdField
'        End Get
'        Set(ByVal value As String)
'            Me.vAliqProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vCIDE() As String
'        Get
'            Return Me.vCIDEField
'        End Get
'        Set(ByVal value As String)
'            Me.vCIDEField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdCombICMSComb

'    Private vBCICMSField As String

'    Private vICMSField As String

'    Private vBCICMSSTField As String

'    Private vICMSSTField As String

'    '''<remarks/>
'    Public Property vBCICMS() As String
'        Get
'            Return Me.vBCICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCICMSST() As String
'        Get
'            Return Me.vBCICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCICMSSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSST() As String
'        Get
'            Return Me.vICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdCombICMSInter

'    Private vBCICMSSTDestField As String

'    Private vICMSSTDestField As String

'    '''<remarks/>
'    Public Property vBCICMSSTDest() As String
'        Get
'            Return Me.vBCICMSSTDestField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCICMSSTDestField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSSTDest() As String
'        Get
'            Return Me.vICMSSTDestField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTDestField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdCombICMSCons

'    Private vBCICMSSTConsField As String

'    Private vICMSSTConsField As String

'    Private uFConsField As TUf

'    '''<remarks/>
'    Public Property vBCICMSSTCons() As String
'        Get
'            Return Me.vBCICMSSTConsField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCICMSSTConsField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSSTCons() As String
'        Get
'            Return Me.vICMSSTConsField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTConsField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UFCons() As TUf
'        Get
'            Return Me.uFConsField
'        End Get
'        Set(ByVal value As TUf)
'            Me.uFConsField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdMed

'    Private nLoteField As String

'    Private qLoteField As String

'    Private dFabField As String

'    Private dValField As String

'    Private vPMCField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nLote() As String
'        Get
'            Return Me.nLoteField
'        End Get
'        Set(ByVal value As String)
'            Me.nLoteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property qLote() As String
'        Get
'            Return Me.qLoteField
'        End Get
'        Set(ByVal value As String)
'            Me.qLoteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dFab() As String
'        Get
'            Return Me.dFabField
'        End Get
'        Set(ByVal value As String)
'            Me.dFabField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dVal() As String
'        Get
'            Return Me.dValField
'        End Get
'        Set(ByVal value As String)
'            Me.dValField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vPMC() As String
'        Get
'            Return Me.vPMCField
'        End Get
'        Set(ByVal value As String)
'            Me.vPMCField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetProdVeicProd

'    Private tpOpField As TNFeInfNFeDetProdVeicProdTpOp

'    Private chassiField As String

'    Private cCorField As String

'    Private xCorField As String

'    Private potField As String

'    Private cM3Field As String

'    Private pesoLField As String

'    Private pesoBField As String

'    Private nSerieField As String

'    Private tpCombField As String

'    Private nMotorField As String

'    Private cMKGField As String

'    Private distField As String

'    Private rENAVAMField As String

'    Private anoModField As String

'    Private anoFabField As String

'    Private tpPintField As String

'    Private tpVeicField As String

'    Private espVeicField As String

'    Private vINField As String

'    Private condVeicField As TNFeInfNFeDetProdVeicProdCondVeic

'    Private cModField As String

'    '''<remarks/>
'    Public Property tpOp() As TNFeInfNFeDetProdVeicProdTpOp
'        Get
'            Return Me.tpOpField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdVeicProdTpOp)
'            Me.tpOpField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property chassi() As String
'        Get
'            Return Me.chassiField
'        End Get
'        Set(ByVal value As String)
'            Me.chassiField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cCor() As String
'        Get
'            Return Me.cCorField
'        End Get
'        Set(ByVal value As String)
'            Me.cCorField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xCor() As String
'        Get
'            Return Me.xCorField
'        End Get
'        Set(ByVal value As String)
'            Me.xCorField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property pot() As String
'        Get
'            Return Me.potField
'        End Get
'        Set(ByVal value As String)
'            Me.potField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CM3() As String
'        Get
'            Return Me.cM3Field
'        End Get
'        Set(ByVal value As String)
'            Me.cM3Field = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property pesoL() As String
'        Get
'            Return Me.pesoLField
'        End Get
'        Set(ByVal value As String)
'            Me.pesoLField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property pesoB() As String
'        Get
'            Return Me.pesoBField
'        End Get
'        Set(ByVal value As String)
'            Me.pesoBField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nSerie() As String
'        Get
'            Return Me.nSerieField
'        End Get
'        Set(ByVal value As String)
'            Me.nSerieField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property tpComb() As String
'        Get
'            Return Me.tpCombField
'        End Get
'        Set(ByVal value As String)
'            Me.tpCombField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nMotor() As String
'        Get
'            Return Me.nMotorField
'        End Get
'        Set(ByVal value As String)
'            Me.nMotorField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CMKG() As String
'        Get
'            Return Me.cMKGField
'        End Get
'        Set(ByVal value As String)
'            Me.cMKGField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property dist() As String
'        Get
'            Return Me.distField
'        End Get
'        Set(ByVal value As String)
'            Me.distField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property RENAVAM() As String
'        Get
'            Return Me.rENAVAMField
'        End Get
'        Set(ByVal value As String)
'            Me.rENAVAMField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property anoMod() As String
'        Get
'            Return Me.anoModField
'        End Get
'        Set(ByVal value As String)
'            Me.anoModField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property anoFab() As String
'        Get
'            Return Me.anoFabField
'        End Get
'        Set(ByVal value As String)
'            Me.anoFabField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property tpPint() As String
'        Get
'            Return Me.tpPintField
'        End Get
'        Set(ByVal value As String)
'            Me.tpPintField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property tpVeic() As String
'        Get
'            Return Me.tpVeicField
'        End Get
'        Set(ByVal value As String)
'            Me.tpVeicField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property espVeic() As String
'        Get
'            Return Me.espVeicField
'        End Get
'        Set(ByVal value As String)
'            Me.espVeicField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property VIN() As String
'        Get
'            Return Me.vINField
'        End Get
'        Set(ByVal value As String)
'            Me.vINField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property condVeic() As TNFeInfNFeDetProdVeicProdCondVeic
'        Get
'            Return Me.condVeicField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetProdVeicProdCondVeic)
'            Me.condVeicField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cMod() As String
'        Get
'            Return Me.cModField
'        End Get
'        Set(ByVal value As String)
'            Me.cModField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetProdVeicProdTpOp

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetProdVeicProdCondVeic

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImposto

'    Private iCMSField As TNFeInfNFeDetImpostoICMS

'    Private iPIField As TNFeInfNFeDetImpostoIPI

'    Private iiField As TNFeInfNFeDetImpostoII

'    Private pISField As TNFeInfNFeDetImpostoPIS

'    Private pISSTField As TNFeInfNFeDetImpostoPISST

'    Private cOFINSField As TNFeInfNFeDetImpostoCOFINS

'    Private cOFINSSTField As TNFeInfNFeDetImpostoCOFINSST

'    Private iSSQNField As TNFeInfNFeDetImpostoISSQN

'    '''<remarks/>
'    Public Property ICMS() As TNFeInfNFeDetImpostoICMS
'        Get
'            Return Me.iCMSField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMS)
'            Me.iCMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property IPI() As TNFeInfNFeDetImpostoIPI
'        Get
'            Return Me.iPIField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoIPI)
'            Me.iPIField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property II() As TNFeInfNFeDetImpostoII
'        Get
'            Return Me.iiField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoII)
'            Me.iiField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property PIS() As TNFeInfNFeDetImpostoPIS
'        Get
'            Return Me.pISField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoPIS)
'            Me.pISField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property PISST() As TNFeInfNFeDetImpostoPISST
'        Get
'            Return Me.pISSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoPISST)
'            Me.pISSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property COFINS() As TNFeInfNFeDetImpostoCOFINS
'        Get
'            Return Me.cOFINSField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoCOFINS)
'            Me.cOFINSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property COFINSST() As TNFeInfNFeDetImpostoCOFINSST
'        Get
'            Return Me.cOFINSSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoCOFINSST)
'            Me.cOFINSSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property ISSQN() As TNFeInfNFeDetImpostoISSQN
'        Get
'            Return Me.iSSQNField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoISSQN)
'            Me.iSSQNField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMS

'    Private itemField As Object

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("ICMS00", GetType(TNFeInfNFeDetImpostoICMSICMS00)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS10", GetType(TNFeInfNFeDetImpostoICMSICMS10)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS20", GetType(TNFeInfNFeDetImpostoICMSICMS20)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS30", GetType(TNFeInfNFeDetImpostoICMSICMS30)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS40", GetType(TNFeInfNFeDetImpostoICMSICMS40)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS51", GetType(TNFeInfNFeDetImpostoICMSICMS51)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS60", GetType(TNFeInfNFeDetImpostoICMSICMS60)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS70", GetType(TNFeInfNFeDetImpostoICMSICMS70)), _
'     System.Xml.Serialization.XmlElementAttribute("ICMS90", GetType(TNFeInfNFeDetImpostoICMSICMS90))> _
'    Public Property Item() As Object
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As Object)
'            Me.itemField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS00

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS00CST

'    Private modBCField As TNFeInfNFeDetImpostoICMSICMS00ModBC

'    Private vBCField As String

'    Private pICMSField As String

'    Private vICMSField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS00CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS00CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBC() As TNFeInfNFeDetImpostoICMSICMS00ModBC
'        Get
'            Return Me.modBCField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS00ModBC)
'            Me.modBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMS() As String
'        Get
'            Return Me.pICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum Torig

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS00CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("00")> _
'    Item00
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS00ModBC

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS10

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS10CST

'    Private modBCField As TNFeInfNFeDetImpostoICMSICMS10ModBC

'    Private vBCField As String

'    Private pICMSField As String

'    Private vICMSField As String

'    Private modBCSTField As TNFeInfNFeDetImpostoICMSICMS10ModBCST

'    Private pMVASTField As String

'    Private pRedBCSTField As String

'    Private vBCSTField As String

'    Private pICMSSTField As String

'    Private vICMSSTField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS10CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS10CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBC() As TNFeInfNFeDetImpostoICMSICMS10ModBC
'        Get
'            Return Me.modBCField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS10ModBC)
'            Me.modBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMS() As String
'        Get
'            Return Me.pICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBCST() As TNFeInfNFeDetImpostoICMSICMS10ModBCST
'        Get
'            Return Me.modBCSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS10ModBCST)
'            Me.modBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pMVAST() As String
'        Get
'            Return Me.pMVASTField
'        End Get
'        Set(ByVal value As String)
'            Me.pMVASTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBCST() As String
'        Get
'            Return Me.pRedBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCST() As String
'        Get
'            Return Me.vBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMSST() As String
'        Get
'            Return Me.pICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSST() As String
'        Get
'            Return Me.vICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS10CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("10")> _
'    Item10
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS10ModBC

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS10ModBCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("4")> _
'    Item4

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("5")> _
'    Item5
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS20

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS20CST

'    Private modBCField As TNFeInfNFeDetImpostoICMSICMS20ModBC

'    Private pRedBCField As String

'    Private vBCField As String

'    Private pICMSField As String

'    Private vICMSField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS20CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS20CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBC() As TNFeInfNFeDetImpostoICMSICMS20ModBC
'        Get
'            Return Me.modBCField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS20ModBC)
'            Me.modBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBC() As String
'        Get
'            Return Me.pRedBCField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMS() As String
'        Get
'            Return Me.pICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS20CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("20")> _
'    Item20
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS20ModBC

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS30

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS30CST

'    Private modBCSTField As TNFeInfNFeDetImpostoICMSICMS30ModBCST

'    Private pMVASTField As String

'    Private pRedBCSTField As String

'    Private vBCSTField As String

'    Private pICMSSTField As String

'    Private vICMSSTField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS30CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS30CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBCST() As TNFeInfNFeDetImpostoICMSICMS30ModBCST
'        Get
'            Return Me.modBCSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS30ModBCST)
'            Me.modBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pMVAST() As String
'        Get
'            Return Me.pMVASTField
'        End Get
'        Set(ByVal value As String)
'            Me.pMVASTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBCST() As String
'        Get
'            Return Me.pRedBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCST() As String
'        Get
'            Return Me.vBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMSST() As String
'        Get
'            Return Me.pICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSST() As String
'        Get
'            Return Me.vICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS30CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("30")> _
'    Item30
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS30ModBCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("4")> _
'    Item4

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("5")> _
'    Item5
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS40

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS40CST

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS40CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS40CST)
'            Me.cSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS40CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("40")> _
'    Item40

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("41")> _
'    Item41

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("50")> _
'    Item50
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS51

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS51CST

'    Private modBCField As TNFeInfNFeDetImpostoICMSICMS51ModBC

'    Private modBCFieldSpecified As Boolean

'    Private pRedBCField As String

'    Private vBCField As String

'    Private pICMSField As String

'    Private vICMSField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS51CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS51CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBC() As TNFeInfNFeDetImpostoICMSICMS51ModBC
'        Get
'            Return Me.modBCField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS51ModBC)
'            Me.modBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property modBCSpecified() As Boolean
'        Get
'            Return Me.modBCFieldSpecified
'        End Get
'        Set(ByVal value As Boolean)
'            Me.modBCFieldSpecified = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBC() As String
'        Get
'            Return Me.pRedBCField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMS() As String
'        Get
'            Return Me.pICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS51CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("51")> _
'    Item51
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS51ModBC

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS60

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS60CST

'    Private vBCSTField As String

'    Private vICMSSTField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS60CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS60CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCST() As String
'        Get
'            Return Me.vBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSST() As String
'        Get
'            Return Me.vICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS60CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("60")> _
'    Item60
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS70

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS70CST

'    Private modBCField As TNFeInfNFeDetImpostoICMSICMS70ModBC

'    Private pRedBCField As String

'    Private vBCField As String

'    Private pICMSField As String

'    Private vICMSField As String

'    Private modBCSTField As TNFeInfNFeDetImpostoICMSICMS70ModBCST

'    Private pMVASTField As String

'    Private pRedBCSTField As String

'    Private vBCSTField As String

'    Private pICMSSTField As String

'    Private vICMSSTField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS70CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS70CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBC() As TNFeInfNFeDetImpostoICMSICMS70ModBC
'        Get
'            Return Me.modBCField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS70ModBC)
'            Me.modBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBC() As String
'        Get
'            Return Me.pRedBCField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMS() As String
'        Get
'            Return Me.pICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBCST() As TNFeInfNFeDetImpostoICMSICMS70ModBCST
'        Get
'            Return Me.modBCSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS70ModBCST)
'            Me.modBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pMVAST() As String
'        Get
'            Return Me.pMVASTField
'        End Get
'        Set(ByVal value As String)
'            Me.pMVASTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBCST() As String
'        Get
'            Return Me.pRedBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCST() As String
'        Get
'            Return Me.vBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMSST() As String
'        Get
'            Return Me.pICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSST() As String
'        Get
'            Return Me.vICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS70CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("70")> _
'    Item70
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS70ModBC

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS70ModBCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("4")> _
'    Item4

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("5")> _
'    Item5
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoICMSICMS90

'    Private origField As Torig

'    Private cSTField As TNFeInfNFeDetImpostoICMSICMS90CST

'    Private modBCField As TNFeInfNFeDetImpostoICMSICMS90ModBC

'    Private vBCField As String

'    Private pRedBCField As String

'    Private pICMSField As String

'    Private vICMSField As String

'    Private modBCSTField As TNFeInfNFeDetImpostoICMSICMS90ModBCST

'    Private pMVASTField As String

'    Private pRedBCSTField As String

'    Private vBCSTField As String

'    Private pICMSSTField As String

'    Private vICMSSTField As String

'    '''<remarks/>
'    Public Property orig() As Torig
'        Get
'            Return Me.origField
'        End Get
'        Set(ByVal value As Torig)
'            Me.origField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoICMSICMS90CST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS90CST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBC() As TNFeInfNFeDetImpostoICMSICMS90ModBC
'        Get
'            Return Me.modBCField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS90ModBC)
'            Me.modBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBC() As String
'        Get
'            Return Me.pRedBCField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMS() As String
'        Get
'            Return Me.pICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property modBCST() As TNFeInfNFeDetImpostoICMSICMS90ModBCST
'        Get
'            Return Me.modBCSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoICMSICMS90ModBCST)
'            Me.modBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pMVAST() As String
'        Get
'            Return Me.pMVASTField
'        End Get
'        Set(ByVal value As String)
'            Me.pMVASTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pRedBCST() As String
'        Get
'            Return Me.pRedBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pRedBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCST() As String
'        Get
'            Return Me.vBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMSST() As String
'        Get
'            Return Me.pICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSST() As String
'        Get
'            Return Me.vICMSSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS90CST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("90")> _
'    Item90
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS90ModBC

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoICMSICMS90ModBCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("4")> _
'    Item4

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("5")> _
'    Item5
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoIPI

'    Private clEnqField As String

'    Private cNPJProdField As String

'    Private cSeloField As String

'    Private qSeloField As String

'    Private cEnqField As String

'    Private itemField As Object

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property clEnq() As String
'        Get
'            Return Me.clEnqField
'        End Get
'        Set(ByVal value As String)
'            Me.clEnqField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property CNPJProd() As String
'        Get
'            Return Me.cNPJProdField
'        End Get
'        Set(ByVal value As String)
'            Me.cNPJProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cSelo() As String
'        Get
'            Return Me.cSeloField
'        End Get
'        Set(ByVal value As String)
'            Me.cSeloField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property qSelo() As String
'        Get
'            Return Me.qSeloField
'        End Get
'        Set(ByVal value As String)
'            Me.qSeloField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property cEnq() As String
'        Get
'            Return Me.cEnqField
'        End Get
'        Set(ByVal value As String)
'            Me.cEnqField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("IPINT", GetType(TNFeInfNFeDetImpostoIPIIPINT)), _
'     System.Xml.Serialization.XmlElementAttribute("IPITrib", GetType(TNFeInfNFeDetImpostoIPIIPITrib))> _
'    Public Property Item() As Object
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As Object)
'            Me.itemField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoIPIIPINT

'    Private cSTField As TNFeInfNFeDetImpostoIPIIPINTCST

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoIPIIPINTCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoIPIIPINTCST)
'            Me.cSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoIPIIPINTCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("01")> _
'    Item01

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("02")> _
'    Item02

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("03")> _
'    Item03

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("04")> _
'    Item04

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("05")> _
'    Item05

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("51")> _
'    Item51

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("52")> _
'    Item52

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("53")> _
'    Item53

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("54")> _
'    Item54

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("55")> _
'    Item55
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoIPIIPITrib

'    Private cSTField As TNFeInfNFeDetImpostoIPIIPITribCST

'    Private itemsField() As String

'    Private itemsElementNameField() As ItemsChoiceType

'    Private vIPIField As String

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoIPIIPITribCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoIPIIPITribCST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("pIPI", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("qUnid", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vBC", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vUnid", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")> _
'    Public Property Items() As String()
'        Get
'            Return Me.itemsField
'        End Get
'        Set(ByVal value As String())
'            Me.itemsField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("ItemsElementName"), _
'     System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemsElementName() As ItemsChoiceType()
'        Get
'            Return Me.itemsElementNameField
'        End Get
'        Set(ByVal value As ItemsChoiceType())
'            Me.itemsElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vIPI() As String
'        Get
'            Return Me.vIPIField
'        End Get
'        Set(ByVal value As String)
'            Me.vIPIField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoIPIIPITribCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("00")> _
'    Item00

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("49")> _
'    Item49

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("50")> _
'    Item50

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("99")> _
'    Item99
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemsChoiceType

'    '''<remarks/>
'    pIPI

'    '''<remarks/>
'    qUnid

'    '''<remarks/>
'    vBC

'    '''<remarks/>
'    vUnid
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoII

'    Private vBCField As String

'    Private vDespAduField As String

'    Private vIIField As String

'    Private vIOFField As String

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vDespAdu() As String
'        Get
'            Return Me.vDespAduField
'        End Get
'        Set(ByVal value As String)
'            Me.vDespAduField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vII() As String
'        Get
'            Return Me.vIIField
'        End Get
'        Set(ByVal value As String)
'            Me.vIIField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vIOF() As String
'        Get
'            Return Me.vIOFField
'        End Get
'        Set(ByVal value As String)
'            Me.vIOFField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoPIS

'    Private itemField As Object

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("PISAliq", GetType(TNFeInfNFeDetImpostoPISPISAliq)), _
'     System.Xml.Serialization.XmlElementAttribute("PISNT", GetType(TNFeInfNFeDetImpostoPISPISNT)), _
'     System.Xml.Serialization.XmlElementAttribute("PISOutr", GetType(TNFeInfNFeDetImpostoPISPISOutr)), _
'     System.Xml.Serialization.XmlElementAttribute("PISQtde", GetType(TNFeInfNFeDetImpostoPISPISQtde))> _
'    Public Property Item() As Object
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As Object)
'            Me.itemField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoPISPISAliq

'    Private cSTField As TNFeInfNFeDetImpostoPISPISAliqCST

'    Private vBCField As String

'    Private pPISField As String

'    Private vPISField As String

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoPISPISAliqCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoPISPISAliqCST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pPIS() As String
'        Get
'            Return Me.pPISField
'        End Get
'        Set(ByVal value As String)
'            Me.pPISField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vPIS() As String
'        Get
'            Return Me.vPISField
'        End Get
'        Set(ByVal value As String)
'            Me.vPISField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoPISPISAliqCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("01")> _
'    Item01

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("02")> _
'    Item02
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoPISPISNT

'    Private cSTField As TNFeInfNFeDetImpostoPISPISNTCST

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoPISPISNTCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoPISPISNTCST)
'            Me.cSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoPISPISNTCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("04")> _
'    Item04

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("06")> _
'    Item06

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("07")> _
'    Item07

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("08")> _
'    Item08

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("09")> _
'    Item09
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoPISPISOutr

'    Private cSTField As TNFeInfNFeDetImpostoPISPISOutrCST

'    Private itemsField() As String

'    Private itemsElementNameField() As ItemsChoiceType1

'    Private vPISField As String

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoPISPISOutrCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoPISPISOutrCST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("pPIS", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("qBCProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vAliqProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vBC", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")> _
'    Public Property Items() As String()
'        Get
'            Return Me.itemsField
'        End Get
'        Set(ByVal value As String())
'            Me.itemsField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("ItemsElementName"), _
'     System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemsElementName() As ItemsChoiceType1()
'        Get
'            Return Me.itemsElementNameField
'        End Get
'        Set(ByVal value As ItemsChoiceType1())
'            Me.itemsElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vPIS() As String
'        Get
'            Return Me.vPISField
'        End Get
'        Set(ByVal value As String)
'            Me.vPISField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoPISPISOutrCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("99")> _
'    Item99
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemsChoiceType1

'    '''<remarks/>
'    pPIS

'    '''<remarks/>
'    qBCProd

'    '''<remarks/>
'    vAliqProd

'    '''<remarks/>
'    vBC
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoPISPISQtde

'    Private cSTField As TNFeInfNFeDetImpostoPISPISQtdeCST

'    Private qBCProdField As String

'    Private vAliqProdField As String

'    Private vPISField As String

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoPISPISQtdeCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoPISPISQtdeCST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property qBCProd() As String
'        Get
'            Return Me.qBCProdField
'        End Get
'        Set(ByVal value As String)
'            Me.qBCProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vAliqProd() As String
'        Get
'            Return Me.vAliqProdField
'        End Get
'        Set(ByVal value As String)
'            Me.vAliqProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vPIS() As String
'        Get
'            Return Me.vPISField
'        End Get
'        Set(ByVal value As String)
'            Me.vPISField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoPISPISQtdeCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("03")> _
'    Item03
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoPISST

'    Private itemsField() As String

'    Private itemsElementNameField() As ItemsChoiceType2

'    Private vPISField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("pPIS", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("qBCProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vAliqProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vBC", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")> _
'    Public Property Items() As String()
'        Get
'            Return Me.itemsField
'        End Get
'        Set(ByVal value As String())
'            Me.itemsField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("ItemsElementName"), _
'     System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemsElementName() As ItemsChoiceType2()
'        Get
'            Return Me.itemsElementNameField
'        End Get
'        Set(ByVal value As ItemsChoiceType2())
'            Me.itemsElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vPIS() As String
'        Get
'            Return Me.vPISField
'        End Get
'        Set(ByVal value As String)
'            Me.vPISField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemsChoiceType2

'    '''<remarks/>
'    pPIS

'    '''<remarks/>
'    qBCProd

'    '''<remarks/>
'    vAliqProd

'    '''<remarks/>
'    vBC
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoCOFINS

'    Private itemField As Object

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("COFINSAliq", GetType(TNFeInfNFeDetImpostoCOFINSCOFINSAliq)), _
'     System.Xml.Serialization.XmlElementAttribute("COFINSNT", GetType(TNFeInfNFeDetImpostoCOFINSCOFINSNT)), _
'     System.Xml.Serialization.XmlElementAttribute("COFINSOutr", GetType(TNFeInfNFeDetImpostoCOFINSCOFINSOutr)), _
'     System.Xml.Serialization.XmlElementAttribute("COFINSQtde", GetType(TNFeInfNFeDetImpostoCOFINSCOFINSQtde))> _
'    Public Property Item() As Object
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As Object)
'            Me.itemField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoCOFINSCOFINSAliq

'    Private cSTField As TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST

'    Private vBCField As String

'    Private pCOFINSField As String

'    Private vCOFINSField As String

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pCOFINS() As String
'        Get
'            Return Me.pCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.pCOFINSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vCOFINS() As String
'        Get
'            Return Me.vCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.vCOFINSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoCOFINSCOFINSAliqCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("01")> _
'    Item01

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("02")> _
'    Item02
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoCOFINSCOFINSNT

'    Private cSTField As TNFeInfNFeDetImpostoCOFINSCOFINSNTCST

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoCOFINSCOFINSNTCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoCOFINSCOFINSNTCST)
'            Me.cSTField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoCOFINSCOFINSNTCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("04")> _
'    Item04

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("06")> _
'    Item06

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("07")> _
'    Item07

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("08")> _
'    Item08

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("09")> _
'    Item09
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoCOFINSCOFINSOutr

'    Private cSTField As TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST

'    Private itemsField() As String

'    Private itemsElementNameField() As ItemsChoiceType3

'    Private vCOFINSField As String

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("pCOFINS", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("qBCProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vAliqProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vBC", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")> _
'    Public Property Items() As String()
'        Get
'            Return Me.itemsField
'        End Get
'        Set(ByVal value As String())
'            Me.itemsField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("ItemsElementName"), _
'     System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemsElementName() As ItemsChoiceType3()
'        Get
'            Return Me.itemsElementNameField
'        End Get
'        Set(ByVal value As ItemsChoiceType3())
'            Me.itemsElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vCOFINS() As String
'        Get
'            Return Me.vCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.vCOFINSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoCOFINSCOFINSOutrCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("99")> _
'    Item99
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemsChoiceType3

'    '''<remarks/>
'    pCOFINS

'    '''<remarks/>
'    qBCProd

'    '''<remarks/>
'    vAliqProd

'    '''<remarks/>
'    vBC
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoCOFINSCOFINSQtde

'    Private cSTField As TNFeInfNFeDetImpostoCOFINSCOFINSQtdeCST

'    Private qBCProdField As String

'    Private vAliqProdField As String

'    Private vCOFINSField As String

'    '''<remarks/>
'    Public Property CST() As TNFeInfNFeDetImpostoCOFINSCOFINSQtdeCST
'        Get
'            Return Me.cSTField
'        End Get
'        Set(ByVal value As TNFeInfNFeDetImpostoCOFINSCOFINSQtdeCST)
'            Me.cSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property qBCProd() As String
'        Get
'            Return Me.qBCProdField
'        End Get
'        Set(ByVal value As String)
'            Me.qBCProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vAliqProd() As String
'        Get
'            Return Me.vAliqProdField
'        End Get
'        Set(ByVal value As String)
'            Me.vAliqProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vCOFINS() As String
'        Get
'            Return Me.vCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.vCOFINSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeDetImpostoCOFINSCOFINSQtdeCST

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("03")> _
'    Item03
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoCOFINSST

'    Private itemsField() As String

'    Private itemsElementNameField() As ItemsChoiceType4

'    Private vCOFINSField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("pCOFINS", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("qBCProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vAliqProd", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("vBC", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")> _
'    Public Property Items() As String()
'        Get
'            Return Me.itemsField
'        End Get
'        Set(ByVal value As String())
'            Me.itemsField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("ItemsElementName"), _
'     System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemsElementName() As ItemsChoiceType4()
'        Get
'            Return Me.itemsElementNameField
'        End Get
'        Set(ByVal value As ItemsChoiceType4())
'            Me.itemsElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vCOFINS() As String
'        Get
'            Return Me.vCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.vCOFINSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemsChoiceType4

'    '''<remarks/>
'    pCOFINS

'    '''<remarks/>
'    qBCProd

'    '''<remarks/>
'    vAliqProd

'    '''<remarks/>
'    vBC
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeDetImpostoISSQN

'    Private vBCField As String

'    Private vAliqField As String

'    Private vISSQNField As String

'    Private cMunFGField As String

'    Private cListServField As TCListServ

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vAliq() As String
'        Get
'            Return Me.vAliqField
'        End Get
'        Set(ByVal value As String)
'            Me.vAliqField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vISSQN() As String
'        Get
'            Return Me.vISSQNField
'        End Get
'        Set(ByVal value As String)
'            Me.vISSQNField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cMunFG() As String
'        Get
'            Return Me.cMunFGField
'        End Get
'        Set(ByVal value As String)
'            Me.cMunFGField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cListServ() As TCListServ
'        Get
'            Return Me.cListServField
'        End Get
'        Set(ByVal value As TCListServ)
'            Me.cListServField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TCListServ

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("101")> _
'    Item101

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("102")> _
'    Item102

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("103")> _
'    Item103

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("104")> _
'    Item104

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("105")> _
'    Item105

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("106")> _
'    Item106

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("107")> _
'    Item107

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("108")> _
'    Item108

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("201")> _
'    Item201

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("302")> _
'    Item302

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("303")> _
'    Item303

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("304")> _
'    Item304

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("305")> _
'    Item305

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("401")> _
'    Item401

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("402")> _
'    Item402

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("403")> _
'    Item403

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("404")> _
'    Item404

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("405")> _
'    Item405

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("406")> _
'    Item406

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("407")> _
'    Item407

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("408")> _
'    Item408

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("409")> _
'    Item409

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("410")> _
'    Item410

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("411")> _
'    Item411

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("412")> _
'    Item412

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("413")> _
'    Item413

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("414")> _
'    Item414

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("415")> _
'    Item415

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("416")> _
'    Item416

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("417")> _
'    Item417

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("418")> _
'    Item418

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("419")> _
'    Item419

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("420")> _
'    Item420

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("421")> _
'    Item421

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("422")> _
'    Item422

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("423")> _
'    Item423

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("501")> _
'    Item501

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("502")> _
'    Item502

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("503")> _
'    Item503

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("504")> _
'    Item504

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("505")> _
'    Item505

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("506")> _
'    Item506

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("507")> _
'    Item507

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("508")> _
'    Item508

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("509")> _
'    Item509

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("601")> _
'    Item601

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("602")> _
'    Item602

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("603")> _
'    Item603

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("604")> _
'    Item604

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("605")> _
'    Item605

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("701")> _
'    Item701

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("702")> _
'    Item702

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("703")> _
'    Item703

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("704")> _
'    Item704

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("705")> _
'    Item705

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("706")> _
'    Item706

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("707")> _
'    Item707

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("708")> _
'    Item708

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("709")> _
'    Item709

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("710")> _
'    Item710

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("711")> _
'    Item711

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("712")> _
'    Item712

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("713")> _
'    Item713

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("716")> _
'    Item716

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("717")> _
'    Item717

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("718")> _
'    Item718

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("719")> _
'    Item719

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("720")> _
'    Item720

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("721")> _
'    Item721

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("722")> _
'    Item722

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("801")> _
'    Item801

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("802")> _
'    Item802

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("901")> _
'    Item901

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("902")> _
'    Item902

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("903")> _
'    Item903

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1001")> _
'    Item1001

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1002")> _
'    Item1002

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1003")> _
'    Item1003

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1004")> _
'    Item1004

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1005")> _
'    Item1005

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1006")> _
'    Item1006

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1007")> _
'    Item1007

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1008")> _
'    Item1008

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1009")> _
'    Item1009

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1010")> _
'    Item1010

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1101")> _
'    Item1101

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1102")> _
'    Item1102

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1103")> _
'    Item1103

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1104")> _
'    Item1104

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1201")> _
'    Item1201

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1202")> _
'    Item1202

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1203")> _
'    Item1203

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1204")> _
'    Item1204

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1205")> _
'    Item1205

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1206")> _
'    Item1206

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1207")> _
'    Item1207

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1208")> _
'    Item1208

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1209")> _
'    Item1209

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1210")> _
'    Item1210

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1211")> _
'    Item1211

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1212")> _
'    Item1212

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1213")> _
'    Item1213

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1214")> _
'    Item1214

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1215")> _
'    Item1215

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1216")> _
'    Item1216

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1217")> _
'    Item1217

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1302")> _
'    Item1302

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1303")> _
'    Item1303

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1304")> _
'    Item1304

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1305")> _
'    Item1305

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1401")> _
'    Item1401

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1402")> _
'    Item1402

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1403")> _
'    Item1403

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1404")> _
'    Item1404

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1405")> _
'    Item1405

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1406")> _
'    Item1406

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1407")> _
'    Item1407

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1408")> _
'    Item1408

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1409")> _
'    Item1409

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1410")> _
'    Item1410

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1411")> _
'    Item1411

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1412")> _
'    Item1412

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1413")> _
'    Item1413

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1501")> _
'    Item1501

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1502")> _
'    Item1502

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1503")> _
'    Item1503

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1504")> _
'    Item1504

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1505")> _
'    Item1505

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1506")> _
'    Item1506

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1507")> _
'    Item1507

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1508")> _
'    Item1508

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1509")> _
'    Item1509

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1510")> _
'    Item1510

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1511")> _
'    Item1511

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1512")> _
'    Item1512

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1513")> _
'    Item1513

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1514")> _
'    Item1514

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1515")> _
'    Item1515

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1516")> _
'    Item1516

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1517")> _
'    Item1517

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1518")> _
'    Item1518

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1601")> _
'    Item1601

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1701")> _
'    Item1701

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1702")> _
'    Item1702

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1703")> _
'    Item1703

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1704")> _
'    Item1704

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1705")> _
'    Item1705

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1706")> _
'    Item1706

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1708")> _
'    Item1708

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1709")> _
'    Item1709

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1710")> _
'    Item1710

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1711")> _
'    Item1711

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1712")> _
'    Item1712

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1713")> _
'    Item1713

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1714")> _
'    Item1714

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1715")> _
'    Item1715

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1716")> _
'    Item1716

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1717")> _
'    Item1717

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1718")> _
'    Item1718

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1719")> _
'    Item1719

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1720")> _
'    Item1720

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1721")> _
'    Item1721

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1722")> _
'    Item1722

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1723")> _
'    Item1723

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1724")> _
'    Item1724

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1801")> _
'    Item1801

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1901")> _
'    Item1901

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2001")> _
'    Item2001

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2002")> _
'    Item2002

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2003")> _
'    Item2003

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2101")> _
'    Item2101

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2201")> _
'    Item2201

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2301")> _
'    Item2301

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2401")> _
'    Item2401

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2501")> _
'    Item2501

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2502")> _
'    Item2502

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2503")> _
'    Item2503

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2504")> _
'    Item2504

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2601")> _
'    Item2601

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2701")> _
'    Item2701

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2801")> _
'    Item2801

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2901")> _
'    Item2901

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3001")> _
'    Item3001

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3101")> _
'    Item3101

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3201")> _
'    Item3201

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3301")> _
'    Item3301

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3401")> _
'    Item3401

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3501")> _
'    Item3501

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3601")> _
'    Item3601

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3701")> _
'    Item3701

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3801")> _
'    Item3801

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3901")> _
'    Item3901

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("4001")> _
'    Item4001
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTotal

'    Private iCMSTotField As TNFeInfNFeTotalICMSTot

'    Private iSSQNtotField As TNFeInfNFeTotalISSQNtot

'    Private retTribField As TNFeInfNFeTotalRetTrib

'    '''<remarks/>
'    Public Property ICMSTot() As TNFeInfNFeTotalICMSTot
'        Get
'            Return Me.iCMSTotField
'        End Get
'        Set(ByVal value As TNFeInfNFeTotalICMSTot)
'            Me.iCMSTotField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property ISSQNtot() As TNFeInfNFeTotalISSQNtot
'        Get
'            Return Me.iSSQNtotField
'        End Get
'        Set(ByVal value As TNFeInfNFeTotalISSQNtot)
'            Me.iSSQNtotField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property retTrib() As TNFeInfNFeTotalRetTrib
'        Get
'            Return Me.retTribField
'        End Get
'        Set(ByVal value As TNFeInfNFeTotalRetTrib)
'            Me.retTribField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTotalICMSTot

'    Private vBCField As String

'    Private vICMSField As String

'    Private vBCSTField As String

'    Private vSTField As String

'    Private vProdField As String

'    Private vFreteField As String

'    Private vSegField As String

'    Private vDescField As String

'    Private vIIField As String

'    Private vIPIField As String

'    Private vPISField As String

'    Private vCOFINSField As String

'    Private vOutroField As String

'    Private vNFField As String

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMS() As String
'        Get
'            Return Me.vICMSField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCST() As String
'        Get
'            Return Me.vBCSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vST() As String
'        Get
'            Return Me.vSTField
'        End Get
'        Set(ByVal value As String)
'            Me.vSTField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vProd() As String
'        Get
'            Return Me.vProdField
'        End Get
'        Set(ByVal value As String)
'            Me.vProdField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vFrete() As String
'        Get
'            Return Me.vFreteField
'        End Get
'        Set(ByVal value As String)
'            Me.vFreteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vSeg() As String
'        Get
'            Return Me.vSegField
'        End Get
'        Set(ByVal value As String)
'            Me.vSegField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vDesc() As String
'        Get
'            Return Me.vDescField
'        End Get
'        Set(ByVal value As String)
'            Me.vDescField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vII() As String
'        Get
'            Return Me.vIIField
'        End Get
'        Set(ByVal value As String)
'            Me.vIIField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vIPI() As String
'        Get
'            Return Me.vIPIField
'        End Get
'        Set(ByVal value As String)
'            Me.vIPIField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vPIS() As String
'        Get
'            Return Me.vPISField
'        End Get
'        Set(ByVal value As String)
'            Me.vPISField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vCOFINS() As String
'        Get
'            Return Me.vCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.vCOFINSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vOutro() As String
'        Get
'            Return Me.vOutroField
'        End Get
'        Set(ByVal value As String)
'            Me.vOutroField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vNF() As String
'        Get
'            Return Me.vNFField
'        End Get
'        Set(ByVal value As String)
'            Me.vNFField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTotalISSQNtot

'    Private vServField As String

'    Private vBCField As String

'    Private vISSField As String

'    Private vPISField As String

'    Private vCOFINSField As String

'    '''<remarks/>
'    Public Property vServ() As String
'        Get
'            Return Me.vServField
'        End Get
'        Set(ByVal value As String)
'            Me.vServField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBC() As String
'        Get
'            Return Me.vBCField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vISS() As String
'        Get
'            Return Me.vISSField
'        End Get
'        Set(ByVal value As String)
'            Me.vISSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vPIS() As String
'        Get
'            Return Me.vPISField
'        End Get
'        Set(ByVal value As String)
'            Me.vPISField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vCOFINS() As String
'        Get
'            Return Me.vCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.vCOFINSField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTotalRetTrib

'    Private vRetPISField As String

'    Private vRetCOFINSField As String

'    Private vRetCSLLField As String

'    Private vBCIRRFField As String

'    Private vIRRFField As String

'    Private vBCRetPrevField As String

'    Private vRetPrevField As String

'    '''<remarks/>
'    Public Property vRetPIS() As String
'        Get
'            Return Me.vRetPISField
'        End Get
'        Set(ByVal value As String)
'            Me.vRetPISField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vRetCOFINS() As String
'        Get
'            Return Me.vRetCOFINSField
'        End Get
'        Set(ByVal value As String)
'            Me.vRetCOFINSField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vRetCSLL() As String
'        Get
'            Return Me.vRetCSLLField
'        End Get
'        Set(ByVal value As String)
'            Me.vRetCSLLField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCIRRF() As String
'        Get
'            Return Me.vBCIRRFField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCIRRFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vIRRF() As String
'        Get
'            Return Me.vIRRFField
'        End Get
'        Set(ByVal value As String)
'            Me.vIRRFField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCRetPrev() As String
'        Get
'            Return Me.vBCRetPrevField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCRetPrevField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vRetPrev() As String
'        Get
'            Return Me.vRetPrevField
'        End Get
'        Set(ByVal value As String)
'            Me.vRetPrevField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTransp

'    Private modFreteField As TNFeInfNFeTranspModFrete

'    Private transportaField As TNFeInfNFeTranspTransporta

'    Private retTranspField As TNFeInfNFeTranspRetTransp

'    Private veicTranspField As TVeiculo

'    Private reboqueField() As TVeiculo

'    Private volField() As TNFeInfNFeTranspVol

'    '''<remarks/>
'    Public Property modFrete() As TNFeInfNFeTranspModFrete
'        Get
'            Return Me.modFreteField
'        End Get
'        Set(ByVal value As TNFeInfNFeTranspModFrete)
'            Me.modFreteField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property transporta() As TNFeInfNFeTranspTransporta
'        Get
'            Return Me.transportaField
'        End Get
'        Set(ByVal value As TNFeInfNFeTranspTransporta)
'            Me.transportaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property retTransp() As TNFeInfNFeTranspRetTransp
'        Get
'            Return Me.retTranspField
'        End Get
'        Set(ByVal value As TNFeInfNFeTranspRetTransp)
'            Me.retTranspField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property veicTransp() As TVeiculo
'        Get
'            Return Me.veicTranspField
'        End Get
'        Set(ByVal value As TVeiculo)
'            Me.veicTranspField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("reboque")> _
'    Public Property reboque() As TVeiculo()
'        Get
'            Return Me.reboqueField
'        End Get
'        Set(ByVal value As TVeiculo())
'            Me.reboqueField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("vol")> _
'    Public Property vol() As TNFeInfNFeTranspVol()
'        Get
'            Return Me.volField
'        End Get
'        Set(ByVal value As TNFeInfNFeTranspVol())
'            Me.volField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeTranspModFrete

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTranspTransporta

'    Private itemField As String

'    Private itemElementNameField As ItemChoiceType2

'    Private xNomeField As String

'    Private ieField As String

'    Private xEnderField As String

'    Private xMunField As String

'    Private ufField As TUf

'    Private ufFieldSpecified As Boolean

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("CNPJ", GetType(String)), _
'     System.Xml.Serialization.XmlElementAttribute("CPF", GetType(String)), _
'     System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemElementName")> _
'    Public Property Item() As String
'        Get
'            Return Me.itemField
'        End Get
'        Set(ByVal value As String)
'            Me.itemField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property ItemElementName() As ItemChoiceType2
'        Get
'            Return Me.itemElementNameField
'        End Get
'        Set(ByVal value As ItemChoiceType2)
'            Me.itemElementNameField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xNome() As String
'        Get
'            Return Me.xNomeField
'        End Get
'        Set(ByVal value As String)
'            Me.xNomeField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property IE() As String
'        Get
'            Return Me.ieField
'        End Get
'        Set(ByVal value As String)
'            Me.ieField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xEnder() As String
'        Get
'            Return Me.xEnderField
'        End Get
'        Set(ByVal value As String)
'            Me.xEnderField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xMun() As String
'        Get
'            Return Me.xMunField
'        End Get
'        Set(ByVal value As String)
'            Me.xMunField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property UF() As TUf
'        Get
'            Return Me.ufField
'        End Get
'        Set(ByVal value As TUf)
'            Me.ufField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlIgnoreAttribute()> _
'    Public Property UFSpecified() As Boolean
'        Get
'            Return Me.ufFieldSpecified
'        End Get
'        Set(ByVal value As Boolean)
'            Me.ufFieldSpecified = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IncludeInSchema:=False)> _
'Public Enum ItemChoiceType2

'    '''<remarks/>
'    CNPJ

'    '''<remarks/>
'    CPF
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTranspRetTransp

'    Private vServField As String

'    Private vBCRetField As String

'    Private pICMSRetField As String

'    Private vICMSRetField As String

'    Private cFOPField As String

'    Private cMunFGField As String

'    '''<remarks/>
'    Public Property vServ() As String
'        Get
'            Return Me.vServField
'        End Get
'        Set(ByVal value As String)
'            Me.vServField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vBCRet() As String
'        Get
'            Return Me.vBCRetField
'        End Get
'        Set(ByVal value As String)
'            Me.vBCRetField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pICMSRet() As String
'        Get
'            Return Me.pICMSRetField
'        End Get
'        Set(ByVal value As String)
'            Me.pICMSRetField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vICMSRet() As String
'        Get
'            Return Me.vICMSRetField
'        End Get
'        Set(ByVal value As String)
'            Me.vICMSRetField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property CFOP() As String
'        Get
'            Return Me.cFOPField
'        End Get
'        Set(ByVal value As String)
'            Me.cFOPField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property cMunFG() As String
'        Get
'            Return Me.cMunFGField
'        End Get
'        Set(ByVal value As String)
'            Me.cMunFGField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTranspVol

'    Private qVolField As String

'    Private espField As String

'    Private marcaField As String

'    Private nVolField As String

'    Private pesoLField As String

'    Private pesoBField As String

'    Private lacresField() As TNFeInfNFeTranspVolLacres

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property qVol() As String
'        Get
'            Return Me.qVolField
'        End Get
'        Set(ByVal value As String)
'            Me.qVolField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property esp() As String
'        Get
'            Return Me.espField
'        End Get
'        Set(ByVal value As String)
'            Me.espField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property marca() As String
'        Get
'            Return Me.marcaField
'        End Get
'        Set(ByVal value As String)
'            Me.marcaField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nVol() As String
'        Get
'            Return Me.nVolField
'        End Get
'        Set(ByVal value As String)
'            Me.nVolField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pesoL() As String
'        Get
'            Return Me.pesoLField
'        End Get
'        Set(ByVal value As String)
'            Me.pesoLField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property pesoB() As String
'        Get
'            Return Me.pesoBField
'        End Get
'        Set(ByVal value As String)
'            Me.pesoBField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("lacres")> _
'    Public Property lacres() As TNFeInfNFeTranspVolLacres()
'        Get
'            Return Me.lacresField
'        End Get
'        Set(ByVal value As TNFeInfNFeTranspVolLacres())
'            Me.lacresField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeTranspVolLacres

'    Private nLacreField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nLacre() As String
'        Get
'            Return Me.nLacreField
'        End Get
'        Set(ByVal value As String)
'            Me.nLacreField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeCobr

'    Private fatField As TNFeInfNFeCobrFat

'    Private dupField() As TNFeInfNFeCobrDup

'    '''<remarks/>
'    Public Property fat() As TNFeInfNFeCobrFat
'        Get
'            Return Me.fatField
'        End Get
'        Set(ByVal value As TNFeInfNFeCobrFat)
'            Me.fatField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("dup")> _
'    Public Property dup() As TNFeInfNFeCobrDup()
'        Get
'            Return Me.dupField
'        End Get
'        Set(ByVal value As TNFeInfNFeCobrDup())
'            Me.dupField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeCobrFat

'    Private nFatField As String

'    Private vOrigField As String

'    Private vDescField As String

'    Private vLiqField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nFat() As String
'        Get
'            Return Me.nFatField
'        End Get
'        Set(ByVal value As String)
'            Me.nFatField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vOrig() As String
'        Get
'            Return Me.vOrigField
'        End Get
'        Set(ByVal value As String)
'            Me.vOrigField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vDesc() As String
'        Get
'            Return Me.vDescField
'        End Get
'        Set(ByVal value As String)
'            Me.vDescField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vLiq() As String
'        Get
'            Return Me.vLiqField
'        End Get
'        Set(ByVal value As String)
'            Me.vLiqField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeCobrDup

'    Private nDupField As String

'    Private dVencField As String

'    Private vDupField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nDup() As String
'        Get
'            Return Me.nDupField
'        End Get
'        Set(ByVal value As String)
'            Me.nDupField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property dVenc() As String
'        Get
'            Return Me.dVencField
'        End Get
'        Set(ByVal value As String)
'            Me.dVencField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property vDup() As String
'        Get
'            Return Me.vDupField
'        End Get
'        Set(ByVal value As String)
'            Me.vDupField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeInfAdic

'    Private infAdFiscoField As String

'    Private infCplField As String

'    Private obsContField() As TNFeInfNFeInfAdicObsCont

'    Private obsFiscoField() As TNFeInfNFeInfAdicObsFisco

'    Private procRefField() As TNFeInfNFeInfAdicProcRef

'    '''<remarks/>
'    Public Property infAdFisco() As String
'        Get
'            Return Me.infAdFiscoField
'        End Get
'        Set(ByVal value As String)
'            Me.infAdFiscoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property infCpl() As String
'        Get
'            Return Me.infCplField
'        End Get
'        Set(ByVal value As String)
'            Me.infCplField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("obsCont")> _
'    Public Property obsCont() As TNFeInfNFeInfAdicObsCont()
'        Get
'            Return Me.obsContField
'        End Get
'        Set(ByVal value As TNFeInfNFeInfAdicObsCont())
'            Me.obsContField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("obsFisco")> _
'    Public Property obsFisco() As TNFeInfNFeInfAdicObsFisco()
'        Get
'            Return Me.obsFiscoField
'        End Get
'        Set(ByVal value As TNFeInfNFeInfAdicObsFisco())
'            Me.obsFiscoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute("procRef")> _
'    Public Property procRef() As TNFeInfNFeInfAdicProcRef()
'        Get
'            Return Me.procRefField
'        End Get
'        Set(ByVal value As TNFeInfNFeInfAdicProcRef())
'            Me.procRefField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeInfAdicObsCont

'    Private xTextoField As String

'    Private xCampoField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xTexto() As String
'        Get
'            Return Me.xTextoField
'        End Get
'        Set(ByVal value As String)
'            Me.xTextoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="token")> _
'    Public Property xCampo() As String
'        Get
'            Return Me.xCampoField
'        End Get
'        Set(ByVal value As String)
'            Me.xCampoField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeInfAdicObsFisco

'    Private xTextoField As String

'    Private xCampoField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xTexto() As String
'        Get
'            Return Me.xTextoField
'        End Get
'        Set(ByVal value As String)
'            Me.xTextoField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="token")> _
'    Public Property xCampo() As String
'        Get
'            Return Me.xCampoField
'        End Get
'        Set(ByVal value As String)
'            Me.xCampoField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeInfAdicProcRef

'    Private nProcField As String

'    Private indProcField As TNFeInfNFeInfAdicProcRefIndProc

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property nProc() As String
'        Get
'            Return Me.nProcField
'        End Get
'        Set(ByVal value As String)
'            Me.nProcField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property indProc() As TNFeInfNFeInfAdicProcRefIndProc
'        Get
'            Return Me.indProcField
'        End Get
'        Set(ByVal value As TNFeInfNFeInfAdicProcRefIndProc)
'            Me.indProcField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Public Enum TNFeInfNFeInfAdicProcRefIndProc

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("0")> _
'    Item0

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("1")> _
'    Item1

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("2")> _
'    Item2

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("3")> _
'    Item3

'    '''<remarks/>
'    <System.Xml.Serialization.XmlEnumAttribute("9")> _
'    Item9
'End Enum

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeExporta

'    Private uFEmbarqField As TUf

'    Private xLocEmbarqField As String

'    '''<remarks/>
'    Public Property UFEmbarq() As TUf
'        Get
'            Return Me.uFEmbarqField
'        End Get
'        Set(ByVal value As TUf)
'            Me.uFEmbarqField = value
'        End Set
'    End Property

'    '''<remarks/>
'    Public Property xLocEmbarq() As String
'        Get
'            Return Me.xLocEmbarqField
'        End Get
'        Set(ByVal value As String)
'            Me.xLocEmbarqField = value
'        End Set
'    End Property
'End Class

''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe")> _
'Partial Public Class TNFeInfNFeCompra

'    Private xNEmpField As String

'    Private xPedField As String

'    Private xContField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xNEmp() As String
'        Get
'            Return Me.xNEmpField
'        End Get
'        Set(ByVal value As String)
'            Me.xNEmpField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xPed() As String
'        Get
'            Return Me.xPedField
'        End Get
'        Set(ByVal value As String)
'            Me.xPedField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property xCont() As String
'        Get
'            Return Me.xContField
'        End Get
'        Set(ByVal value As String)
'            Me.xContField = value
'        End Set
'    End Property
'End Class




'
'This source code was auto-generated by xsd, Version=2.0.50727.1432.
'



''''<remarks/>
'<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"), _
' System.SerializableAttribute(), _
' System.Diagnostics.DebuggerStepThroughAttribute(), _
' System.ComponentModel.DesignerCategoryAttribute("code"), _
' System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=True, [Namespace]:="http://www.portalfiscal.inf.br/nfe"), _
' System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe", IsNullable:=False)> _
'Partial Public Class cabecMsg

'    Private versaoDadosField As String

'    Private versaoField As String

'    '''<remarks/>
'    <System.Xml.Serialization.XmlElementAttribute(DataType:="token")> _
'    Public Property versaoDados() As String
'        Get
'            Return Me.versaoDadosField
'        End Get
'        Set(ByVal value As String)
'            Me.versaoDadosField = value
'        End Set
'    End Property

'    '''<remarks/>
'    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="token")> _
'    Public Property versao() As String
'        Get
'            Return Me.versaoField
'        End Get
'        Set(ByVal value As String)
'            Me.versaoField = value
'        End Set
'    End Property
'End Class

''******************fim da cabecMsg ******************


'    Const SUCESSO As Integer = 0

'Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click


'    'Dim coll As NameValueCollection
'    'Dim col1 As New NameValueCollection
'    ''        lErro = wsConvert.WS_ConsultaRecibo("teste", lRet, sRet)
'    'coll = Request.QueryString
'    'Dim x As New ServiceReference1.nfeRetRecepcaoRequestBody
'    'Dim z As New ServiceReference1.nfeRetRecepcaoRequest
'    'Dim x1 As New ServiceReference1.NfeRetRecepcaoSoapClient

'    'Dim a1 As New ServiceReference2.nfeRecepcaoLoteRequest
'    'Dim a2 As New ServiceReference2.nfeRecepcaoLoteRequestBody
'    'Dim a3 As New ServiceReference2.NfeRecepcaoSoapClient
'    Dim s1 As String
'    Dim a4 As cabecMsg = New cabecMsg

'    Dim sArquivo As String

'    Dim Assinatura As AssinaturaDigital

'    Dim i1 As Integer
'    Dim iIndice As Integer

'    Dim XMLStream As MemoryStream = New MemoryStream(10000)
'    Dim XMLStream1 As MemoryStream = New MemoryStream(10000)
'    Dim XMLStreamCabec As MemoryStream = New MemoryStream(10000)
'    Dim XMLStreamRet As MemoryStream = New MemoryStream(10000)
'    Dim XMLStreamDados As MemoryStream = New MemoryStream(10000)



'    Dim db1 As DataClassesDataContext = New DataClassesDataContext

'    Dim dic As DataClasses2DataContext = New DataClasses2DataContext


'    Dim results As IEnumerable(Of NFeNFiscal)
'    Dim resFiliaisClientes As IEnumerable(Of FiliaisCliente)
'    Dim resFiliaisFornecedores As IEnumerable(Of FiliaisFornecedore)
'    Dim resEndereco As IEnumerable(Of Endereco)
'    Dim resEstado As IEnumerable(Of Estado)
'    Dim resTitRec As IEnumerable(Of TitulosRecTodo)
'    Dim resTitPag As IEnumerable(Of TitulosPagTodo)
'    Dim resFilialEmpresa As IEnumerable(Of FiliaisEmpresa)
'    Dim resCidade As IEnumerable(Of Cidade)
'    Dim resPais As IEnumerable(Of Paise)
'    Dim resCliente As IEnumerable(Of Cliente)
'    Dim resFornecedor As IEnumerable(Of Fornecedore)
'    Dim resEndDest As IEnumerable(Of Endereco)
'    Dim resItemNF As IEnumerable(Of ItensNFiscal)
'    Dim resProduto As IEnumerable(Of Produto)
'    Dim resTribItemNF As IEnumerable(Of TributacaoItemNF)
'    Dim resTipoTribICMS As IEnumerable(Of TiposTribICM)
'    Dim resTipoTribIPI As IEnumerable(Of TiposTribIPI)
'    Dim resItemAdicaoDIItemNF As IEnumerable(Of ItemAdicaoDIItemNF)
'    Dim resItensAdicaoDI As IEnumerable(Of ItensAdicaoDI)
'    Dim resTributacaoNF As IEnumerable(Of TributacaoNF)
'    Dim resDIInfo As IEnumerable(Of DIInfo)
'    Dim resAdicaoDI As IEnumerable(Of AdicaoDI)
'    Dim resTransp As IEnumerable(Of Transportadora)
'    Dim resParcPag As IEnumerable(Of ParcelasPagToda)
'    Dim resParcRec As IEnumerable(Of ParcelasRecToda)
'    Dim resFatConfig As IEnumerable(Of FATConfig)


'    Dim objFatConfig As FATConfig
'    Dim objParcRec As ParcelasRecToda
'    Dim objParcPag As ParcelasPagToda
'    Dim objTransp As Transportadora
'    Dim objAdicaoDI As AdicaoDI
'    Dim objDIInfo As DIInfo
'    Dim objTributacaoNF As TributacaoNF
'    Dim objItensAdicaoDI As ItensAdicaoDI
'    Dim objItemAdicaoDiItemNF As ItemAdicaoDIItemNF
'    Dim objTipoTribIPI As TiposTribIPI
'    Dim objTipoTribICMS As TiposTribICM
'    Dim objTribItemNF As TributacaoItemNF
'    Dim objProduto As Produto
'    Dim objItemNF As ItensNFiscal
'    Dim objCliente As Cliente
'    Dim objFornecedor As Fornecedore
'    Dim objPais As Paise
'    Dim objCidade As Cidade
'    Dim objFilialEmpresa As FiliaisEmpresa
'    Dim objTitRec As TitulosRecTodo
'    Dim objTitPag As TitulosPagTodo
'    Dim objEstado As Estado
'    Dim objEndereco As Endereco
'    Dim objNFiscal As NFeNFiscal
'    Dim objFiliaisClientes As FiliaisCliente
'    Dim objFiliaisFornecedores As FiliaisFornecedore
'    Dim lEndereco As Long
'    Dim lEndDest As Long
'    Dim objNFiscalParam As NFeNFiscal

'    Dim objRefNfe As TNFeInfNFeIdeNFref

'    Dim iCST As Integer
'    Dim lErro As Long
'    Dim dPISAliquota As Double
'    Dim dCOFINSAliquota As Double
'    Dim lNumIntDIInfo As Long
'    Dim sideCUF As String

'    Dim colNFiscal As Collection = New Collection

'    Dim envioNFe As TEnviNFe = New TEnviNFe

'    Dim resdicEmpresa As IEnumerable(Of Empresa)

'    Dim objEmpresa As Empresa

'    Dim XMLString As String
'    Dim XMLString1 As String
'    Dim XMLString2 As String
'    Dim XMLStringNFes As String

'    Dim XMLStringCabec As String

'    Dim IPITrib As TNFeInfNFeDetImpostoIPIIPITrib = New TNFeInfNFeDetImpostoIPIIPITrib

'    Dim iResult As Integer

'    Dim cert As X509Certificate2 = New X509Certificate2
'    Dim certificado As Certificado = New Certificado


'    Dim NfeRecepcao As New br.gov.rs.sefazvirtual.nfe.homologacao.NfeRecepcao
'    Dim NFeRetRecepCao As New br.gov.rs.sefazvirtual.nfe.homologacao1.NfeRetRecepcao

'    Dim lNumIntNF As Long

'    Dim objValidaXML As ClassValidaXML = New ClassValidaXML

'    Try


'        resFatConfig = db1.ExecuteQuery(Of FATConfig) _
'        ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_PROX_LOTE_NFE")

'        objFatConfig = resFatConfig(0)

'        iResult = db1.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", CLng(objFatConfig.Conteudo + 1), "NUM_PROX_LOTE_NFE")

'        envioNFe.idLote = objFatConfig.Conteudo


'        '        On Error GoTo Erro_Button1_Click

'        '        db1.Connection.ConnectionString = "DSN=SGEDados5;UID=sa;PWD=SAPWD;"
'        db1.Connection.Open()
'        '        dic.Connection.ConnectionString = "DSN=SGEDic5;UID=sa;PWD=SAPWD;"
'        dic.Connection.Open()

'        '
'        '  seleciona certificado do repositório MY do windows
'        '
'        cert = Certificado.BuscaNome("")

'        db1.Transaction = db1.Connection.BeginTransaction()




'        a4.versao = "1.02"
'        a4.versaoDados = "1.10"

'        Dim mySerializercabec As New XmlSerializer(GetType(cabecMsg))

'        XMLStreamCabec = New MemoryStream(10000)

'        mySerializercabec.Serialize(XMLStreamCabec, a4)

'        Dim doccabec As XmlDocument = New XmlDocument
'        XMLStreamCabec.Position = 0
'        doccabec.Load(XMLStreamCabec)
'        doccabec.Save("c:\nfe\XmlCabec.xml")

'        Call objValidaXML.validaXML("c:\nfe\Xmlcabec.xml", "c:\nfe\cabecMsg_v1.02.xsd")

'        Dim xmcabec As Byte()

'        xmcabec = XMLStreamCabec.ToArray

'        XMLStringCabec = System.Text.Encoding.UTF8.GetString(xmcabec)

'        XMLStringCabec = Mid(XMLStringCabec, 1, 19) & " encoding=""utf-8"" " & Mid(XMLStringCabec, 20)

'        lNumIntNF = NumIntNF.Text

'        'lNumIntNFiscalParam
'        results = db1.ExecuteQuery(Of NFeNFiscal) _
'        ("SELECT * FROM NFeNFiscal WHERE NumIntDoc = {0} ", lNumIntNF)

'        For Each objNFiscal In results

'            colNFiscal.Add(objNFiscal)

'            Dim a5 As TNFe = New TNFe

'            Dim infNFe As TNFeInfNFe = New TNFeInfNFe
'            a5.infNFe = infNFe

'            a5.infNFe.versao = "1.10"

'            Dim infNFeIde As TNFeInfNFeIde = New TNFeInfNFeIde
'            a5.infNFe.ide = infNFeIde

'            lNumIntDIInfo = 0

'            resFilialEmpresa = db1.ExecuteQuery(Of FiliaisEmpresa) _
'            ("SELECT * FROM FiliaisEmpresa WHERE FilialEmpresa = {0} ", objNFiscal.FilialEmpresa)

'            For Each objFilialEmpresa In resFilialEmpresa
'                lEndereco = objFilialEmpresa.Endereco
'                Exit For
'            Next

'            resEndereco = db1.ExecuteQuery(Of Endereco) _
'            ("SELECT * FROM Enderecos WHERE Codigo = {0}", lEndereco)

'            objEndereco = resEndereco(0)

'            resEstado = db1.ExecuteQuery(Of Estado) _
'                ("SELECT * FROM Estados WHERE Sigla = {0}", objEndereco.SiglaEstado)

'            'a5.infNFe.ide.cUF = resEstado(0).CodIBGE

'            objEstado = resEstado(0)

'            sideCUF = objEstado.CodIBGE

'            Select Case objEstado.CodIBGE

'                Case 11
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item11
'                Case 12
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item12
'                Case 13
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item13
'                Case 14
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item14
'                Case 15
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item15
'                Case 16
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item16
'                Case 17
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item17
'                Case 21
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item21
'                Case 22
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item22
'                Case 23
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item23
'                Case 24
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item24
'                Case 25
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item25
'                Case 26
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item26
'                Case 27
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item27
'                Case 28
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item28
'                Case 29
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item29
'                Case 31
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item31
'                Case 32
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item32
'                Case 33
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item33
'                Case 35
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item35
'                Case 41
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item41
'                Case 42
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item42
'                Case 43
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item43
'                Case 50
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item50
'                Case 51
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item51
'                Case 52
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item52
'                Case 53
'                    a5.infNFe.ide.cUF = TCodUfIBGE.Item53

'            End Select


'            a5.infNFe.ide.tpAmb = TAmb.Item2

'            resCidade = db1.ExecuteQuery(Of Cidade) _
'            ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndereco.Cidade)

'            objCidade = resCidade(0)

'            resPais = db1.ExecuteQuery(Of Paise) _
'            ("SELECT * FROM Paises WHERE Codigo = {0}", objEndereco.CodigoPais)

'            objPais = resPais(0)

'            resFatConfig = db1.ExecuteQuery(Of FATConfig) _
'            ("SELECT * FROM FatConfig WHERE Codigo = {0} ", "NUM_PROX_NFE")

'            objFatConfig = resFatConfig(0)

'            iResult = db1.ExecuteCommand("UPDATE FatConfig Set Conteudo = {0} WHERE Codigo = {1}", CLng(objFatConfig.Conteudo + 1), "NUM_PROX_NFE")

'            a5.infNFe.ide.cNF = Format(CLng(objFatConfig.Conteudo), "000000000")
'            a5.infNFe.ide.natOp = objNFiscal.DescrNF

'            a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item2


'            a5.infNFe.ide.mod = TMod.Item55
'            a5.infNFe.ide.serie = objNFiscal.Serie
'            a5.infNFe.ide.nNF = objNFiscal.NumNotaFiscal
'            a5.infNFe.ide.tpImp = TNFeInfNFeIdeTpImp.Item1
'            a5.infNFe.ide.dEmi = Format(objNFiscal.DataEmissao, "yyyy-MM-dd")

'            'se for nota de entrada
'            If objNFiscal.Tipo = 1 Then
'                a5.infNFe.ide.dSaiEnt = Format(objNFiscal.DataEntrada, "yyyy-MM-dd")
'                a5.infNFe.ide.tpNF = TNFeInfNFeIdeTpNF.Item0
'            Else
'                a5.infNFe.ide.dSaiEnt = Format(objNFiscal.DataSaida, "yyyy-MM-dd")
'                a5.infNFe.ide.tpNF = TNFeInfNFeIdeTpNF.Item1
'            End If

'            a5.infNFe.ide.procEmi = TProcEmi.Item0
'            a5.infNFe.ide.verProc = "Corporator"

'            Dim infNFeEmit As TNFeInfNFeEmit = New TNFeInfNFeEmit
'            a5.infNFe.emit = infNFeEmit

'            a5.infNFe.emit.ItemElementName = ItemChoiceType.CNPJ
'            a5.infNFe.emit.Item = objFilialEmpresa.CGC

'            'iEmpresaParam
'            resdicEmpresa = dic.ExecuteQuery(Of Empresa) _
'            ("SELECT * FROM Empresas WHERE Codigo = {0}", 1)

'            For Each objEmpresa In resdicEmpresa
'                a5.infNFe.emit.xNome = objEmpresa.Nome
'                Exit For
'            Next

'            Dim enderEmit As TEnderEmi = New TEnderEmi
'            a5.infNFe.emit.enderEmit = enderEmit

'            a5.infNFe.emit.enderEmit.xLgr = objEndereco.Endereco
'            a5.infNFe.emit.enderEmit.nro = "0"
'            a5.infNFe.emit.enderEmit.xBairro = objEndereco.Bairro

'            'se for Brasil 
'            If objPais.CodBacen = 1058 Then
'                a5.infNFe.emit.enderEmit.cMun = objCidade.CodIBGE
'                a5.infNFe.emit.enderEmit.xMun = objCidade.Descricao
'                '                a5.infNFe.emit.enderEmit.UF = objEndereco.SiglaEstado
'                Select Case objEndereco.SiglaEstado

'                    Case "AC"
'                        a5.infNFe.emit.enderEmit.UF = TUf.AC

'                    Case "AL"
'                        a5.infNFe.emit.enderEmit.UF = TUf.AL

'                    Case "AM"
'                        a5.infNFe.emit.enderEmit.UF = TUf.AM

'                    Case "AP"
'                        a5.infNFe.emit.enderEmit.UF = TUf.AP

'                    Case "BA"
'                        a5.infNFe.emit.enderEmit.UF = TUf.BA

'                    Case "CE"
'                        a5.infNFe.emit.enderEmit.UF = TUf.CE

'                    Case "DF"
'                        a5.infNFe.emit.enderEmit.UF = TUf.DF

'                    Case "ES"
'                        a5.infNFe.emit.enderEmit.UF = TUf.ES

'                    Case "GO"
'                        a5.infNFe.emit.enderEmit.UF = TUf.GO

'                    Case "MA"
'                        a5.infNFe.emit.enderEmit.UF = TUf.MA

'                    Case "MG"
'                        a5.infNFe.emit.enderEmit.UF = TUf.MG

'                    Case "MS"
'                        a5.infNFe.emit.enderEmit.UF = TUf.MS

'                    Case "MT"
'                        a5.infNFe.emit.enderEmit.UF = TUf.MT

'                    Case "PA"
'                        a5.infNFe.emit.enderEmit.UF = TUf.PA

'                    Case "PB"
'                        a5.infNFe.emit.enderEmit.UF = TUf.PB

'                    Case "PE"
'                        a5.infNFe.emit.enderEmit.UF = TUf.PE

'                    Case "PI"
'                        a5.infNFe.emit.enderEmit.UF = TUf.PI

'                    Case "PR"
'                        a5.infNFe.emit.enderEmit.UF = TUf.PR

'                    Case "RJ"
'                        a5.infNFe.emit.enderEmit.UF = TUf.RJ

'                    Case "RN"
'                        a5.infNFe.emit.enderEmit.UF = TUf.RN

'                    Case "RO"
'                        a5.infNFe.emit.enderEmit.UF = TUf.RO

'                    Case "RR"
'                        a5.infNFe.emit.enderEmit.UF = TUf.RR

'                    Case "RS"
'                        a5.infNFe.emit.enderEmit.UF = TUf.RS

'                    Case "SC"
'                        a5.infNFe.emit.enderEmit.UF = TUf.SC

'                    Case "SE"
'                        a5.infNFe.emit.enderEmit.UF = TUf.SE

'                    Case "SP"
'                        a5.infNFe.emit.enderEmit.UF = TUf.SP

'                    Case "TO"
'                        a5.infNFe.emit.enderEmit.UF = TUf.TO

'                End Select
'                If Len(objEndereco.CEP) > 0 Then
'                    a5.infNFe.emit.enderEmit.CEP = objEndereco.CEP
'                End If
'            Else
'                a5.infNFe.emit.enderEmit.cMun = "9999999"
'                a5.infNFe.emit.enderEmit.xMun = "EXTERIOR"
'                a5.infNFe.emit.enderEmit.UF = TUf.EX
'            End If

'            a5.infNFe.emit.enderEmit.cPais = objPais.CodBacen
'            a5.infNFe.emit.enderEmit.xPais = objPais.Nome
'            a5.infNFe.emit.IE = objFilialEmpresa.InscricaoEstadual


'            Dim infNFeDest As TNFeInfNFeDest = New TNFeInfNFeDest
'            a5.infNFe.dest = infNFeDest

'            'se o destinatrio for o cliente
'            If objNFiscal.Destinatario = 1 Then

'                resFiliaisClientes = db1.ExecuteQuery(Of FiliaisCliente) _
'                ("SELECT * FROM FiliaisClientes WHERE CodCliente = {0} AND CodFilial = {1}", objNFiscal.Cliente, objNFiscal.FilialCli)

'                For Each objFiliaisClientes In resFiliaisClientes

'                    lEndDest = objFiliaisClientes.Endereco

'                    If Len(objFiliaisClientes.CGC) = 11 Then
'                        a5.infNFe.dest.ItemElementName = ItemChoiceType1.CPF
'                    Else
'                        a5.infNFe.dest.ItemElementName = ItemChoiceType1.CNPJ
'                    End If

'                    a5.infNFe.dest.Item = objFiliaisClientes.CGC
'                    a5.infNFe.dest.IE = objFiliaisClientes.InscricaoEstadual
'                    If Len(Trim(objFiliaisClientes.InscricaoSuframa)) > 0 Then
'                        a5.infNFe.dest.ISUF = objFiliaisClientes.InscricaoSuframa
'                    End If
'                    Exit For
'                Next

'                resCliente = db1.ExecuteQuery(Of Cliente) _
'                ("SELECT * FROM Clientes WHERE Codigo = {0}", objNFiscal.Cliente)

'                For Each objCliente In resCliente
'                    a5.infNFe.dest.xNome = objCliente.RazaoSocial
'                    Exit For
'                Next


'                'se o desti
'            ElseIf objNFiscal.Destinatario = 2 Then

'                resFiliaisFornecedores = db1.ExecuteQuery(Of FiliaisFornecedore) _
'                ("SELECT * FROM FiliaisFornecedores WHERE CodFornecedor = {0} AND CodFilial = {1}", objNFiscal.Fornecedor, objNFiscal.FilialForn)

'                For Each objFiliaisFornecedores In resFiliaisFornecedores

'                    lEndDest = objFiliaisFornecedores.Endereco

'                    If Len(objFiliaisFornecedores.CGC) = 11 Then
'                        a5.infNFe.dest.ItemElementName = ItemChoiceType1.CPF
'                    Else
'                        a5.infNFe.dest.ItemElementName = ItemChoiceType1.CNPJ
'                    End If

'                    a5.infNFe.dest.Item = objFiliaisFornecedores.CGC
'                    a5.infNFe.dest.IE = objFiliaisFornecedores.InscricaoEstadual
'                    Exit For
'                Next

'                resFornecedor = db1.ExecuteQuery(Of Fornecedore) _
'                ("SELECT * FROM Fornecedores WHERE Codigo = {0}", objNFiscal.Fornecedor)

'                For Each objFornecedor In resFornecedor
'                    a5.infNFe.dest.xNome = objFornecedor.RazaoSocial
'                    Exit For
'                Next

'            Else
'                lEndDest = lEndereco
'                a5.infNFe.dest.xNome = objEmpresa.Nome
'                a5.infNFe.dest.IE = objFilialEmpresa.InscricaoEstadual

'            End If

'            resEndDest = db1.ExecuteQuery(Of Endereco) _
'            ("SELECT * FROM Enderecos WHERE Codigo = {0}", lEndDest)

'            For Each objEndDest In resEndDest

'                resEstado = db1.ExecuteQuery(Of Estado) _
'                ("SELECT * FROM Estados WHERE Sigla = {0}", objEndDest.SiglaEstado)

'                For Each objEstado In resEstado

'                    Exit For
'                Next

'                resCidade = db1.ExecuteQuery(Of Cidade) _
'                ("SELECT * FROM Cidades WHERE Descricao = {0}", objEndDest.Cidade)

'                For Each objCidade In resCidade
'                    a5.infNFe.ide.cMunFG = objCidade.CodIBGE
'                    Exit For
'                Next

'                Dim enderDest As TEndereco = New TEndereco
'                a5.infNFe.dest.enderDest = enderDest


'                resPais = db1.ExecuteQuery(Of Paise) _
'                ("SELECT * FROM Paises WHERE Codigo = {0}", objEndDest.CodigoPais)

'                For Each objPais In resPais
'                    a5.infNFe.dest.enderDest.cPais = objPais.CodBacen
'                    Exit For
'                Next

'                a5.infNFe.dest.enderDest.xLgr = objEndDest.Endereco
'                a5.infNFe.dest.enderDest.nro = "0"
'                a5.infNFe.dest.enderDest.xBairro = objEndDest.Bairro

'                'se for Brasil 
'                If objPais.CodBacen = 1058 Then
'                    a5.infNFe.dest.enderDest.cMun = objCidade.CodIBGE
'                    a5.infNFe.ide.cMunFG = objCidade.CodIBGE

'                    a5.infNFe.dest.enderDest.xMun = objCidade.Descricao
'                    '                    a5.infNFe.dest.enderDest.UF = objEndDest.SiglaEstado
'                    Select Case objEndDest.SiglaEstado

'                        Case "AC"
'                            a5.infNFe.dest.enderDest.UF = TUf.AC

'                        Case "AL"
'                            a5.infNFe.dest.enderDest.UF = TUf.AL

'                        Case "AM"
'                            a5.infNFe.dest.enderDest.UF = TUf.AM

'                        Case "AP"
'                            a5.infNFe.dest.enderDest.UF = TUf.AP

'                        Case "BA"
'                            a5.infNFe.dest.enderDest.UF = TUf.BA

'                        Case "CE"
'                            a5.infNFe.dest.enderDest.UF = TUf.CE

'                        Case "DF"
'                            a5.infNFe.dest.enderDest.UF = TUf.DF

'                        Case "ES"
'                            a5.infNFe.dest.enderDest.UF = TUf.ES

'                        Case "GO"
'                            a5.infNFe.dest.enderDest.UF = TUf.GO

'                        Case "MA"
'                            a5.infNFe.dest.enderDest.UF = TUf.MA

'                        Case "MG"
'                            a5.infNFe.dest.enderDest.UF = TUf.MG

'                        Case "MS"
'                            a5.infNFe.dest.enderDest.UF = TUf.MS

'                        Case "MT"
'                            a5.infNFe.dest.enderDest.UF = TUf.MT

'                        Case "PA"
'                            a5.infNFe.dest.enderDest.UF = TUf.PA

'                        Case "PB"
'                            a5.infNFe.dest.enderDest.UF = TUf.PB

'                        Case "PE"
'                            a5.infNFe.dest.enderDest.UF = TUf.PE

'                        Case "PI"
'                            a5.infNFe.dest.enderDest.UF = TUf.PI

'                        Case "PR"
'                            a5.infNFe.dest.enderDest.UF = TUf.PR

'                        Case "RJ"
'                            a5.infNFe.dest.enderDest.UF = TUf.RJ

'                        Case "RN"
'                            a5.infNFe.dest.enderDest.UF = TUf.RN

'                        Case "RO"
'                            a5.infNFe.dest.enderDest.UF = TUf.RO

'                        Case "RR"
'                            a5.infNFe.dest.enderDest.UF = TUf.RR

'                        Case "RS"
'                            a5.infNFe.dest.enderDest.UF = TUf.RS

'                        Case "SC"
'                            a5.infNFe.dest.enderDest.UF = TUf.SC

'                        Case "SE"
'                            a5.infNFe.dest.enderDest.UF = TUf.SE

'                        Case "SP"
'                            a5.infNFe.dest.enderDest.UF = TUf.SP

'                        Case "TO"
'                            a5.infNFe.dest.enderDest.UF = TUf.TO

'                    End Select

'                    If Len(objEndDest.CEP) > 0 Then
'                        a5.infNFe.dest.enderDest.CEP = objEndDest.CEP
'                    End If
'                Else
'                    a5.infNFe.dest.enderDest.cMun = "9999999"
'                    a5.infNFe.ide.cMunFG = "9999999"
'                    a5.infNFe.dest.enderDest.xMun = "EXTERIOR"
'                    a5.infNFe.dest.enderDest.UF = TUf.EX

'                End If

'                Exit For
'            Next



'            'lNumIntNFiscalParam
'            resItemNF = db1.ExecuteQuery(Of ItensNFiscal) _
'            ("SELECT * FROM ItensNFiscal WHERE  NumIntNF = {0} ORDER BY Item", lNumIntNF)

'            Dim NFDet(2) As TNFeInfNFeDet
'            a5.infNFe.det() = NFDet

'            iIndice = -1
'            For Each objItemNF In resItemNF

'                iIndice = iIndice + 1

'                Dim infNFeDet As TNFeInfNFeDet = New TNFeInfNFeDet
'                a5.infNFe.det(iIndice) = infNFeDet

'                Dim infNFeDetProd As TNFeInfNFeDetProd = New TNFeInfNFeDetProd
'                a5.infNFe.det(iIndice).prod = infNFeDetProd

'                a5.infNFe.det(iIndice).nItem = objItemNF.Item
'                a5.infNFe.det(iIndice).prod.cProd = objItemNF.Produto
'                a5.infNFe.det(iIndice).prod.cEAN = ""

'                resProduto = db1.ExecuteQuery(Of Produto) _
'                ("SELECT * FROM Produtos WHERE  Codigo = {0}", objItemNF.Produto)

'                For Each objProduto In resProduto

'                    a5.infNFe.det(iIndice).prod.xProd = objProduto.Descricao
'                    If Len(Trim(a5.infNFe.det(iIndice).prod.EXTIPI)) > 0 Then
'                        a5.infNFe.det(iIndice).prod.EXTIPI = objProduto.IPICodigo
'                    End If
'                    Exit For

'                Next

'                'lNumIntNFiscalPar
'                resTribItemNF = db1.ExecuteQuery(Of TributacaoItemNF) _
'                ("SELECT * FROM TributacaoItemNF WHERE  NumIntNF = {0} AND Item = {1}", lNumIntNF, objItemNF.Item)

'                For Each objTribItemNF In resTribItemNF
'                    a5.infNFe.det(iIndice).prod.CFOP = objTribItemNF.NaturezaOp
'                    Exit For
'                Next

'                a5.infNFe.det(iIndice).prod.uCom = objItemNF.UnidadeMed
'                a5.infNFe.det(iIndice).prod.qCom = Replace(Format(objItemNF.Quantidade, "######0.0000"), ",", ".")
'                a5.infNFe.det(iIndice).prod.vUnCom = Replace(Format(objItemNF.PrecoUnitario, "##########0.0000"), ",", ".")
'                a5.infNFe.det(iIndice).prod.vProd = Replace(Format(objItemNF.PrecoUnitario * objItemNF.Quantidade, "fixed"), ",", ".")
'                a5.infNFe.det(iIndice).prod.cEANTrib = ""
'                a5.infNFe.det(iIndice).prod.uTrib = objItemNF.UnidadeMed
'                a5.infNFe.det(iIndice).prod.qTrib = Replace(Format(objItemNF.Quantidade, "######0.0000"), ",", ".")
'                a5.infNFe.det(iIndice).prod.vUnTrib = Replace(Format(objItemNF.PrecoUnitario, "##########0.0000"), ",", ".")
'                If objItemNF.ValorDesconto > 0 Then
'                    a5.infNFe.det(iIndice).prod.vDesc = Replace(Format(objItemNF.ValorDesconto, "fixed"), ",", ".")
'                End If

'                '*************** ICMS ***************************************

'                Dim infNFeDetImposto As TNFeInfNFeDetImposto = New TNFeInfNFeDetImposto
'                a5.infNFe.det(iIndice).imposto = infNFeDetImposto

'                Dim infNFeDetImpostoICMS As TNFeInfNFeDetImpostoICMS = New TNFeInfNFeDetImpostoICMS
'                a5.infNFe.det(iIndice).imposto.ICMS = infNFeDetImpostoICMS

'                resTipoTribICMS = db1.ExecuteQuery(Of TiposTribICM) _
'                ("SELECT * FROM TiposTribICMS WHERE  Tipo = {0}", objTribItemNF.ICMSTipo)

'                For Each objTipoTribICMS In resTipoTribICMS
'                    Exit For
'                Next


'                Select Case objTipoTribICMS.TipoTribCST

'                    'tributacao integral
'                    Case 0
'                        Dim ICMS00 As New TNFeInfNFeDetImpostoICMSICMS00
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS00
'                        ICMS00.orig = objProduto.OrigemMercadoria
'                        ICMS00.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        ICMS00.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
'                        ICMS00.vBC = Replace(Format(objTribItemNF.ICMSBase, "fixed"), ",", ".")
'                        ICMS00.pICMS = Replace(Format(objTribItemNF.ICMSAliquota * 100, "##0.00"), ",", ".")
'                        ICMS00.vICMS = Replace(Format(objTribItemNF.ICMSValor, "fixed"), ",", ".")

'                    Case 10
'                        Dim ICMS10 As New TNFeInfNFeDetImpostoICMSICMS10
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS10
'                        ICMS10.orig = objProduto.OrigemMercadoria
'                        ICMS10.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        ICMS10.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
'                        ICMS10.vBC = Replace(Format(objTribItemNF.ICMSBase, "fixed"), ",", ".")
'                        ICMS10.pICMS = Replace(Format(objTribItemNF.ICMSAliquota * 100, "##0.00"), ",", ".")
'                        ICMS10.vICMS = Replace(Format(objTribItemNF.ICMSValor, "fixed"), ",", ".")
'                        ICMS10.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
'                        ICMS10.pMVAST = Replace(Format((objTribItemNF.ICMSSubstBase / objTribItemNF.ICMSBase - 1) * 100, "##0.00"), ",", ".")
'                        ICMS10.vBCST = Replace(Format(objTribItemNF.ICMSSubstBase, "fixed"), ",", ".")
'                        ICMS10.pICMSST = Replace(Format(objTribItemNF.ICMSSubstAliquota * 100, "##0.00"), ",", ".")
'                        ICMS10.vICMSST = Replace(Format(objTribItemNF.ICMSSubstValor, "fixed"), ",", ".")

'                    Case 20
'                        Dim ICMS20 As New TNFeInfNFeDetImpostoICMSICMS20
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS20
'                        ICMS20.orig = objProduto.OrigemMercadoria
'                        ICMS20.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        ICMS20.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
'                        ICMS20.pRedBC = Replace(Format(objTribItemNF.ICMSPercRedBase * 100, "##0.00"), ",", ".")
'                        ICMS20.vBC = Replace(Format(objTribItemNF.ICMSBase, "fixed"), ",", ".")
'                        ICMS20.pICMS = Replace(Format(objTribItemNF.ICMSAliquota * 100, "##0.00"), ",", ".")
'                        ICMS20.vICMS = Replace(Format(objTribItemNF.ICMSValor, "fixed"), ",", ".")

'                    Case 30
'                        Dim ICMS30 As New TNFeInfNFeDetImpostoICMSICMS30
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS30
'                        ICMS30.orig = objProduto.OrigemMercadoria
'                        ICMS30.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        ICMS30.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
'                        ICMS30.pMVAST = Replace(Format((objTribItemNF.ICMSSubstBase / objTribItemNF.ICMSBase - 1) * 100, "##0.00"), ",", ".")
'                        ICMS30.vBCST = Replace(Format(objTribItemNF.ICMSSubstBase, "fixed"), ",", ".")
'                        ICMS30.pICMSST = Replace(Format(objTribItemNF.ICMSSubstAliquota * 100, "##0.00"), ",", ".")
'                        ICMS30.vICMSST = Replace(Format(objTribItemNF.ICMSSubstValor, "fixed"), ",", ".")

'                    Case 40, 41, 50
'                        Dim ICMS40 As New TNFeInfNFeDetImpostoICMSICMS40
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS40
'                        ICMS40.orig = objProduto.OrigemMercadoria
'                        ICMS40.CST = Format(objTipoTribICMS.TipoTribCST, "00")

'                    Case 51
'                        Dim ICMS51 As New TNFeInfNFeDetImpostoICMSICMS51
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS51
'                        ICMS51.orig = objProduto.OrigemMercadoria
'                        ICMS51.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        ICMS51.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
'                        ICMS51.pRedBC = Replace(Format(objTribItemNF.ICMSPercRedBase * 100, "##0.00"), ",", ".")
'                        ICMS51.vBC = Replace(Format(objTribItemNF.ICMSBase, "fixed"), ",", ".")
'                        ICMS51.pICMS = Replace(Format(objTribItemNF.ICMSAliquota * 100, "##0.00"), ",", ".")
'                        ICMS51.vICMS = Replace(Format(objTribItemNF.ICMSValor, "fixed"), ",", ".")

'                    Case 60
'                        Dim ICMS60 As New TNFeInfNFeDetImpostoICMSICMS60
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS60
'                        ICMS60.orig = objProduto.OrigemMercadoria
'                        ICMS60.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        '???????                        ICMS60.vBCST = Format(objTribItemNF.ICMSSubstBase, "fixed")
'                        '???????                        ICMS60.vICMSST = Format(objTribItemNF.ICMSSubstValor, "fixed")


'                    Case 70
'                        Dim ICMS70 As New TNFeInfNFeDetImpostoICMSICMS70
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS70
'                        ICMS70.orig = objProduto.OrigemMercadoria
'                        ICMS70.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        ICMS70.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
'                        ICMS70.pRedBC = Replace(Format(objTribItemNF.ICMSPercRedBase * 100, "##0.00"), ",", ".")
'                        ICMS70.vBC = Replace(Format(objTribItemNF.ICMSBase, "fixed"), ",", ".")
'                        ICMS70.pICMS = Replace(Format(objTribItemNF.ICMSAliquota * 100, "##0.00"), ",", ".")
'                        ICMS70.vICMS = Replace(Format(objTribItemNF.ICMSValor, "fixed"), ",", ".")
'                        ICMS70.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
'                        ICMS70.pMVAST = Replace(Format((objTribItemNF.ICMSSubstBase / objTribItemNF.ICMSBase - 1) * 100, "##0.00"), ",", ".")
'                        ICMS70.vBCST = Replace(Format(objTribItemNF.ICMSSubstBase, "fixed"), ",", ".")
'                        ICMS70.pICMSST = Replace(Format(objTribItemNF.ICMSSubstAliquota * 100, "##0.00"), ",", ".")
'                        ICMS70.vICMSST = Replace(Format(objTribItemNF.ICMSSubstValor, "fixed"), ",", ".")

'                    Case 90
'                        Dim ICMS90 As New TNFeInfNFeDetImpostoICMSICMS90
'                        a5.infNFe.det(iIndice).imposto.ICMS.Item = ICMS90
'                        ICMS90.orig = objProduto.OrigemMercadoria
'                        ICMS90.CST = Format(objTipoTribICMS.TipoTribCST, "00")
'                        ICMS90.modBC = TNFeInfNFeDetImpostoICMSICMS00ModBC.Item3
'                        ICMS90.pRedBC = Replace(Format(objTribItemNF.ICMSPercRedBase * 100, "##0.00"), ",", ".")
'                        ICMS90.vBC = Replace(Format(objTribItemNF.ICMSBase, "fixed"), ",", ".")
'                        ICMS90.pICMS = Replace(Format(objTribItemNF.ICMSAliquota * 100, "##0.00"), ",", ".")
'                        ICMS90.vICMS = Replace(Format(objTribItemNF.ICMSValor, "fixed"), ",", ".")
'                        ICMS90.modBCST = TNFeInfNFeDetImpostoICMSICMS10ModBCST.Item4
'                        ICMS90.pMVAST = Replace(Format((objTribItemNF.ICMSSubstBase / objTribItemNF.ICMSBase - 1) * 100, "##0.00"), ",", ".")
'                        ICMS90.vBCST = Replace(Format(objTribItemNF.ICMSSubstBase, "fixed"), ",", ".")
'                        ICMS90.pICMSST = Replace(Format(objTribItemNF.ICMSSubstAliquota * 100, "##0.00"), ",", ".")
'                        ICMS90.vICMSST = Replace(Format(objTribItemNF.ICMSSubstValor, "fixed"), ",", ".")

'                End Select

'                '********************  IPI *******************************************

'                Dim infNFeDetImpostoIPI As TNFeInfNFeDetImpostoIPI = New TNFeInfNFeDetImpostoIPI
'                a5.infNFe.det(iIndice).imposto.IPI = infNFeDetImpostoIPI

'                a5.infNFe.det(iIndice).imposto.IPI.cEnq = "999"

'                resTipoTribIPI = db1.ExecuteQuery(Of TiposTribIPI) _
'                ("SELECT * FROM TiposTribIPI WHERE  Tipo = {0}", objTribItemNF.IPITipo)

'                For Each objTipoTribIPI In resTipoTribIPI
'                    Exit For
'                Next

'                'se for uma nota de entrada                    
'                If objNFiscal.Tipo = 1 Then
'                    If objTipoTribIPI.CSTEntrada = 0 Or objTipoTribIPI.CSTEntrada = 49 Then
'                        '                        Dim IPITrib As TNFeInfNFeDetImpostoIPIIPITrib = New TNFeInfNFeDetImpostoIPIIPITrib
'                        a5.infNFe.det(iIndice).imposto.IPI.Item = IPITrib


'                        Select Case objTipoTribIPI.CSTEntrada
'                            Case 0
'                                IPITrib.CST = TNFeInfNFeDetImpostoIPIIPITribCST.Item00
'                            Case 49
'                                IPITrib.CST = TNFeInfNFeDetImpostoIPIIPITribCST.Item49
'                        End Select
'                        IPITrib.ItemsElementName(1) = ItemsChoiceType.vBC
'                        IPITrib.Items(1) = Replace(Format(objTribItemNF.IPIBaseCalculo, "fixed"), ",", ".")
'                        IPITrib.ItemsElementName(2) = ItemsChoiceType.pIPI
'                        IPITrib.Items(2) = Replace(Format(objTribItemNF.IPIAliquota * 100, "##0.00"), ",", ".")
'                        IPITrib.vIPI = Replace(Format(objTribItemNF.IPIValor, "fixed"), ",", ".")
'                    ElseIf objTipoTribIPI.CSTEntrada = 1 Or objTipoTribIPI.CSTEntrada = 2 Or objTipoTribIPI.CSTEntrada = 3 Or objTipoTribIPI.CSTEntrada = 4 Or objTipoTribIPI.CSTEntrada = 5 Then
'                        Dim IPINT As TNFeInfNFeDetImpostoIPIIPINT = New TNFeInfNFeDetImpostoIPIIPINT
'                        a5.infNFe.det(iIndice).imposto.IPI.Item = IPINT
'                        IPINT.CST = Format(objTipoTribIPI.CSTEntrada, "00")
'                    End If

'                Else
'                    'é uma nota de saida

'                    If objTipoTribIPI.CSTSaida = 50 Or objTipoTribIPI.CSTSaida = 99 Then
'                        '                        Dim IPITrib As TNFeInfNFeDetImpostoIPIIPITrib = New TNFeInfNFeDetImpostoIPIIPITrib
'                        a5.infNFe.det(iIndice).imposto.IPI.Item = IPITrib

'                        Select Case objTipoTribIPI.CSTSaida
'                            Case 50
'                                IPITrib.CST = TNFeInfNFeDetImpostoIPIIPITribCST.Item50
'                            Case 99
'                                IPITrib.CST = TNFeInfNFeDetImpostoIPIIPITribCST.Item99
'                        End Select


'                        Dim ItemsElementName(2) As ItemsChoiceType
'                        Dim ItemsString(2) As String

'                        IPITrib.ItemsElementName = ItemsElementName
'                        IPITrib.Items = ItemsString

'                        'Dim ItensChoiceType1 As ItemsChoiceType = New ItemsChoiceType
'                        'IPITrib.ItemsElementName(1) = ItensChoiceType1

'                        IPITrib.ItemsElementName(0) = ItemsChoiceType.vBC
'                        IPITrib.Items(0) = Replace(Format(objTribItemNF.IPIBaseCalculo, "fixed"), ",", ".")

'                        'Dim ItensChoiceType2 As ItemsChoiceType = New ItemsChoiceType
'                        'IPITrib.ItemsElementName(2) = ItensChoiceType2


'                        IPITrib.ItemsElementName(1) = ItemsChoiceType.pIPI
'                        IPITrib.Items(1) = Replace(Format(objTribItemNF.IPIAliquota * 100, "##0.00"), ",", ".")
'                        IPITrib.vIPI = Replace(Format(objTribItemNF.IPIValor, "fixed"), ",", ".")
'                    ElseIf objTipoTribIPI.CSTSaida = 51 Or objTipoTribIPI.CSTSaida = 52 Or objTipoTribIPI.CSTSaida = 53 Or objTipoTribIPI.CSTSaida = 54 Or objTipoTribIPI.CSTSaida = 55 Then
'                        Dim IPINT As TNFeInfNFeDetImpostoIPIIPINT = New TNFeInfNFeDetImpostoIPIIPINT
'                        a5.infNFe.det(iIndice).imposto.IPI.Item = IPINT
'                        IPINT.CST = Format(objTipoTribIPI.CSTSaida, "00")
'                    End If


'                End If

'                '***********  Imposto de IMPORTACAO ****************************


'                resItemAdicaoDIItemNF = db1.ExecuteQuery(Of ItemAdicaoDIItemNF) _
'                ("SELECT * FROM ItemAdicaoDIItemNF WHERE  NumIntItemNF = {0}", objItemNF.NumIntDoc)

'                For Each objItemAdicaoDiItemNF In resItemAdicaoDIItemNF

'                    Dim infNFeDetImpostoII As TNFeInfNFeDetImpostoII = New TNFeInfNFeDetImpostoII
'                    a5.infNFe.det(iIndice).imposto.II = infNFeDetImpostoII

'                    a5.infNFe.det(iIndice).imposto.II.vII = Replace(Format(objItemAdicaoDiItemNF.ValorII, "fixed"), ",", ".")
'                    a5.infNFe.det(iIndice).imposto.II.vBC = Replace(Format(objItemAdicaoDiItemNF.ValorAduaneiro, "fixed"), ",", ".")

'                    resItensAdicaoDI = db1.ExecuteQuery(Of ItensAdicaoDI) _
'                    ("SELECT * FROM ItensAdicaoDI WHERE  NumIntDoc = {0}", objItemAdicaoDiItemNF.NumIntItemAdicaoDI)

'                    For Each objItensAdicaoDI In resItensAdicaoDI
'                        Exit For
'                    Next

'                    a5.infNFe.det(iIndice).imposto.II.vDespAdu = Replace(Format(objItensAdicaoDI.ValorTotalFOBEmReal - objItensAdicaoDI.ValorTotalCIFEmReal, "fixed"), ",", ".")

'                    If lNumIntDIInfo = 0 Then

'                        resAdicaoDI = db1.ExecuteQuery(Of AdicaoDI) _
'                        ("SELECT * FROM ItensAdicaoDI WHERE  NumIntDoc = {0}", objItensAdicaoDI.NumIntAdicaoDI)

'                        For Each objAdicaoDI In resAdicaoDI
'                            lNumIntDIInfo = objAdicaoDI.NumIntDI
'                            Exit For
'                        Next
'                    End If

'                    Exit For

'                Next

'                '***********  PIS ****************************

'                Dim infNFeDetImpostoPIS As TNFeInfNFeDetImpostoPIS = New TNFeInfNFeDetImpostoPIS
'                a5.infNFe.det(iIndice).imposto.PIS = infNFeDetImpostoPIS

'                lErro = PIS_CST(iCST, objTribItemNF)
'                If lErro <> SUCESSO Then Error 10000

'                Select Case iCST

'                    Case 0, 1
'                        Dim PISAliq As New TNFeInfNFeDetImpostoPISPISAliq

'                        a5.infNFe.det(iIndice).imposto.PIS.Item = PISAliq


'                        PISAliq.CST = Format(iCST, "00")

'                        lErro = PIS_Aliquota(dPISAliquota, objFilialEmpresa)
'                        If lErro <> SUCESSO Then Error 10001

'                        PISAliq.pPIS = Replace(Format(dPISAliquota * 100, "##0.00"), ",", ".")
'                        PISAliq.vPIS = Replace(Format(objTribItemNF.PISCredito, "fixed"), ",", ".")
'                        If dPISAliquota <> 0 Then
'                            PISAliq.vBC = Replace(Format(objTribItemNF.PISCredito / dPISAliquota, "fixed"), ",", ".")
'                        End If

'                    Case 3
'                        Dim PISQtde As New TNFeInfNFeDetImpostoPISPISQtde
'                        a5.infNFe.det(iIndice).imposto.PIS.Item = PISQtde

'                        PISQtde.CST = Format(iCST, "00")

'                    Case 4, 6, 7, 8, 9
'                        Dim PISNT As New TNFeInfNFeDetImpostoPISPISNT
'                        a5.infNFe.det(iIndice).imposto.PIS.Item = PISNT

'                        PISNT.CST = Format(iCST, "00")

'                    Case 99
'                        Dim PISOutr As New TNFeInfNFeDetImpostoPISPISOutr
'                        a5.infNFe.det(iIndice).imposto.PIS.Item = PISOutr


'                        PISOutr.CST = Format(iCST, "00")

'                End Select

'                '***********  COFINS ****************************

'                Dim infNFeDetImpostoCOFINS As TNFeInfNFeDetImpostoCOFINS = New TNFeInfNFeDetImpostoCOFINS
'                a5.infNFe.det(iIndice).imposto.COFINS = infNFeDetImpostoCOFINS

'                lErro = COFINS_CST(iCST, objTribItemNF)
'                If lErro <> SUCESSO Then Error 10002

'                Select Case iCST

'                    Case 0, 1
'                        Dim COFINSAliq As New TNFeInfNFeDetImpostoCOFINSCOFINSAliq

'                        a5.infNFe.det(iIndice).imposto.COFINS.Item = COFINSAliq


'                        COFINSAliq.CST = Format(iCST, "00")

'                        lErro = COFINS_Aliquota(dCOFINSAliquota, objFilialEmpresa)
'                        If lErro <> SUCESSO Then Error 10003

'                        COFINSAliq.pCOFINS = Replace(Format(dCOFINSAliquota * 100, "##0.00"), ",", ".")
'                        COFINSAliq.vCOFINS = Replace(Format(objTribItemNF.COFINSCredito, "fixed"), ",", ".")
'                        If dCOFINSAliquota <> 0 Then
'                            COFINSAliq.vBC = Replace(Format(objTribItemNF.COFINSCredito / dCOFINSAliquota, "fixed"), ",", ".")
'                        End If

'                    Case 3
'                        Dim COFINSQtde As New TNFeInfNFeDetImpostoCOFINSCOFINSQtde
'                        a5.infNFe.det(iIndice).imposto.COFINS.Item = COFINSQtde

'                        COFINSQtde.CST = Format(iCST, "00")

'                    Case 4, 6, 7, 8, 9
'                        Dim COFINSNT As New TNFeInfNFeDetImpostoCOFINSCOFINSNT
'                        a5.infNFe.det(iIndice).imposto.COFINS.Item = COFINSNT

'                        COFINSNT.CST = Format(iCST, "00")

'                    Case 99
'                        Dim COFINSOutr As New TNFeInfNFeDetImpostoCOFINSCOFINSOutr
'                        a5.infNFe.det(iIndice).imposto.COFINS.Item = COFINSOutr


'                        COFINSOutr.CST = Format(iCST, "00")

'                End Select

'            Next

'            '***********  total ****************************

'            Dim infNFeTotal As TNFeInfNFeTotal = New TNFeInfNFeTotal
'            a5.infNFe.total = infNFeTotal


'            '***********  icms total ****************************

'            Dim infNFeTotalICMSTot As TNFeInfNFeTotalICMSTot = New TNFeInfNFeTotalICMSTot
'            a5.infNFe.total.ICMSTot = infNFeTotalICMSTot

'            resTributacaoNF = db1.ExecuteQuery(Of TributacaoNF) _
'            ("SELECT *  FROM TributacaoNF WHERE  NumIntDoc = {0}", objNFiscal.NumIntDoc)

'            For Each objTributacaoNF In resTributacaoNF
'                a5.infNFe.total.ICMSTot.vBC = Replace(Format(objTributacaoNF.ICMSBase, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vICMS = Replace(Format(objTributacaoNF.ICMSValor, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vBCST = Replace(Format(objTributacaoNF.ICMSSubstBase, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vST = Replace(Format(objTributacaoNF.ICMSSubstValor, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vProd = Replace(Format(objNFiscal.ValorProdutos, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vFrete = Replace(Format(objNFiscal.ValorFrete, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vSeg = Replace(Format(objNFiscal.ValorSeguro, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vDesc = Replace(Format(objNFiscal.ValorDesconto, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vIPI = Replace(Format(objTributacaoNF.IPIValor, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vPIS = Replace(Format(objTributacaoNF.PISCredito, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vCOFINS = Replace(Format(objTributacaoNF.COFINSCredito, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vOutro = Replace(Format(objNFiscal.ValorOutrasDespesas, "fixed"), ",", ".")
'                a5.infNFe.total.ICMSTot.vNF = Replace(Format(objNFiscal.ValorTotal, "fixed"), ",", ".")
'                Exit For

'            Next

'            a5.infNFe.total.ICMSTot.vII = Replace(Format(0, "fixed"), ",", ".")

'            If lNumIntDIInfo <> 0 Then
'                resDIInfo = db1.ExecuteQuery(Of DIInfo) _
'                ("SELECT * FROM DIInfo WHERE  NumIntDoc = {0}", lNumIntDIInfo)

'                For Each objDIInfo In resDIInfo
'                    a5.infNFe.total.ICMSTot.vII = Replace(Format(objDIInfo.IIValor, "fixed"), ",", ".")
'                    Exit For
'                Next

'            End If

'            '***********  retencao ****************************

'            Dim infNFeTotalRetTrib As TNFeInfNFeTotalRetTrib = New TNFeInfNFeTotalRetTrib
'            a5.infNFe.total.retTrib = infNFeTotalRetTrib

'            If objTributacaoNF.PISRetido > 0.0 Then
'                a5.infNFe.total.retTrib.vRetPIS = Replace(Format(objTributacaoNF.PISRetido, "fixed"), ",", ".")
'            End If
'            If objTributacaoNF.COFINSRetido > 0 Then
'                a5.infNFe.total.retTrib.vRetCOFINS = Replace(Format(objTributacaoNF.COFINSRetido, "fixed"), ",", ".")
'            End If
'            If objTributacaoNF.CSLLRetido > 0 Then
'                a5.infNFe.total.retTrib.vRetCSLL = Replace(Format(objTributacaoNF.CSLLRetido, "fixed"), ",", ".")
'            End If
'            If objTributacaoNF.IRRFBase > 0 Then
'                a5.infNFe.total.retTrib.vBCIRRF = Replace(Format(objTributacaoNF.IRRFBase, "fixed"), ",", ".")
'            End If
'            If objTributacaoNF.IRRFValor > 0 Then
'                a5.infNFe.total.retTrib.vIRRF = Replace(Format(objTributacaoNF.IRRFValor, "fixed"), ",", ".")
'            End If
'            If objTributacaoNF.INSSValorBase > 0 Then
'                a5.infNFe.total.retTrib.vBCRetPrev = Replace(Format(objTributacaoNF.INSSValorBase, "fixed"), ",", ".")
'            End If
'            If objTributacaoNF.INSSRetido > 0 Then
'                a5.infNFe.total.retTrib.vBCRetPrev = Replace(Format(objTributacaoNF.INSSRetido, "fixed"), ",", ".")
'            End If

'            '***********  transportadora ****************************

'            Dim infNFeTransp As TNFeInfNFeTransp = New TNFeInfNFeTransp
'            a5.infNFe.transp = infNFeTransp

'            Dim infNFeTranspTransporta As TNFeInfNFeTranspTransporta = New TNFeInfNFeTranspTransporta
'            a5.infNFe.transp.transporta = infNFeTranspTransporta

'            a5.infNFe.transp.modFrete = objNFiscal.FreteRespons - 1

'            resTransp = db1.ExecuteQuery(Of Transportadora) _
'                ("SELECT * FROM Transportadoras WHERE  Codigo = {0}", objNFiscal.CodTransportadora)

'            For Each objTransp In resTransp
'                If Len(objTransp.CGC) = 14 Then
'                    a5.infNFe.transp.transporta.ItemElementName = ItemChoiceType2.CNPJ
'                Else
'                    a5.infNFe.transp.transporta.ItemElementName = ItemChoiceType2.CPF
'                End If
'                a5.infNFe.transp.transporta.Item = objTransp.CGC
'                a5.infNFe.transp.transporta.xNome = objTransp.Nome
'                a5.infNFe.transp.transporta.IE = objTransp.InscricaoEstadual

'                resEndereco = db1.ExecuteQuery(Of Endereco) _
'                    ("SELECT * FROM Transportadoras WHERE  Codigo = {0}", objNFiscal.CodTransportadora)

'                For Each objEndereco In resEndereco
'                    a5.infNFe.transp.transporta.xEnder = objEndereco.Endereco
'                    a5.infNFe.transp.transporta.xMun = objEndereco.Cidade
'                    a5.infNFe.transp.transporta.UF = objEndereco.SiglaEstado
'                    Exit For
'                Next
'                Exit For

'            Next

'            '***********  veiculo ****************************

'            If Len(Trim(objNFiscal.Placa)) > 0 And Len(Trim(objNFiscal.PlacaUF)) > 0 Then

'                Dim veiculo As TVeiculo = New TVeiculo
'                a5.infNFe.transp.veicTransp = veiculo

'                a5.infNFe.transp.veicTransp.placa = objNFiscal.Placa
'                '            a5.infNFe.transp.veicTransp.UF = objNFiscal.PlacaUF
'                Select Case objNFiscal.PlacaUF

'                    Case "AC"
'                        a5.infNFe.transp.veicTransp.UF = TUf.AC

'                    Case "AL"
'                        a5.infNFe.transp.veicTransp.UF = TUf.AL

'                    Case "AM"
'                        a5.infNFe.transp.veicTransp.UF = TUf.AM

'                    Case "AP"
'                        a5.infNFe.transp.veicTransp.UF = TUf.AP

'                    Case "BA"
'                        a5.infNFe.transp.veicTransp.UF = TUf.BA

'                    Case "CE"
'                        a5.infNFe.transp.veicTransp.UF = TUf.CE

'                    Case "DF"
'                        a5.infNFe.transp.veicTransp.UF = TUf.DF

'                    Case "ES"
'                        a5.infNFe.transp.veicTransp.UF = TUf.ES

'                    Case "GO"
'                        a5.infNFe.transp.veicTransp.UF = TUf.GO

'                    Case "MA"
'                        a5.infNFe.transp.veicTransp.UF = TUf.MA

'                    Case "MG"
'                        a5.infNFe.transp.veicTransp.UF = TUf.MG

'                    Case "MS"
'                        a5.infNFe.transp.veicTransp.UF = TUf.MS

'                    Case "MT"
'                        a5.infNFe.transp.veicTransp.UF = TUf.MT

'                    Case "PA"
'                        a5.infNFe.transp.veicTransp.UF = TUf.PA

'                    Case "PB"
'                        a5.infNFe.transp.veicTransp.UF = TUf.PB

'                    Case "PE"
'                        a5.infNFe.transp.veicTransp.UF = TUf.PE

'                    Case "PI"
'                        a5.infNFe.transp.veicTransp.UF = TUf.PI

'                    Case "PR"
'                        a5.infNFe.transp.veicTransp.UF = TUf.PR

'                    Case "RJ"
'                        a5.infNFe.transp.veicTransp.UF = TUf.RJ

'                    Case "RN"
'                        a5.infNFe.transp.veicTransp.UF = TUf.RN

'                    Case "RO"
'                        a5.infNFe.transp.veicTransp.UF = TUf.RO

'                    Case "RR"
'                        a5.infNFe.transp.veicTransp.UF = TUf.RR

'                    Case "RS"
'                        a5.infNFe.transp.veicTransp.UF = TUf.RS

'                    Case "SC"
'                        a5.infNFe.transp.veicTransp.UF = TUf.SC

'                    Case "SE"
'                        a5.infNFe.transp.veicTransp.UF = TUf.SE

'                    Case "SP"
'                        a5.infNFe.transp.veicTransp.UF = TUf.SP

'                    Case "TO"
'                        a5.infNFe.transp.veicTransp.UF = TUf.TO

'                End Select

'            End If
'            '***********  volume ****************************

'            Dim infNFeTranspVol(1) As TNFeInfNFeTranspVol

'            a5.infNFe.transp.vol = infNFeTranspVol

'            Dim infNFeTranspVol1 As TNFeInfNFeTranspVol = New TNFeInfNFeTranspVol

'            a5.infNFe.transp.vol(0) = infNFeTranspVol1

'            a5.infNFe.transp.vol(0).qVol = objNFiscal.VolumeQuant
'            a5.infNFe.transp.vol(0).esp = objNFiscal.VolumeEspecie
'            a5.infNFe.transp.vol(0).marca = objNFiscal.VolumeMarca
'            If Len(Trim(objNFiscal.VolumeNumero)) > 0 Then a5.infNFe.transp.vol(0).nVol = objNFiscal.VolumeNumero
'            a5.infNFe.transp.vol(0).pesoL = Replace(Format(objNFiscal.PesoLiq, "##########0.000"), ",", ".")
'            a5.infNFe.transp.vol(0).pesoB = Replace(Format(objNFiscal.PesoBruto, "##########0.000"), ",", ".")

'            '***********  cobranca ****************************

'            Dim infNFeCobr As TNFeInfNFeCobr = New TNFeInfNFeCobr
'            a5.infNFe.cobr = infNFeCobr

'            Dim infNFeCobrFat As TNFeInfNFeCobrFat = New TNFeInfNFeCobrFat
'            a5.infNFe.cobr.fat = infNFeCobrFat

'            'se fir um titulo a pagar
'            If objNFiscal.ClasseDocCPR = 1 Then

'                resTitPag = db1.ExecuteQuery(Of TitulosPagTodo) _
'                ("SELECT * FROM TitulosPagTodos WHERE NumIntDoc = {0}", objNFiscal.NumIntDocCPR)

'                For Each objTitPag In resTitPag
'                    If objTitPag.CondicaoPagto = 1 Then
'                        a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item0
'                    Else
'                        a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item1
'                    End If
'                    a5.infNFe.cobr.fat.nFat = objNFiscal.NumNotaFiscal
'                    a5.infNFe.cobr.fat.vOrig = Replace(Format(objTitPag.ValorTotal, "fixed"), ",", ".")
'                    a5.infNFe.cobr.fat.vLiq = Replace(Format(objTitPag.ValorTotal, "fixed"), ",", ".")

'                    resParcPag = db1.ExecuteQuery(Of ParcelasPagToda) _
'                    ("SELECT * FROM ParcelasPagTodas WHERE NumIntTitulo = {0}", objNFiscal.NumIntDocCPR)

'                    iIndice = -1

'                    Dim Dup(50) As TNFeInfNFeCobrDup

'                    a5.infNFe.cobr.dup = Dup

'                    For Each objParcPag In resParcPag
'                        iIndice = iIndice + 1


'                        Dim infNFeCobrDup As TNFeInfNFeCobrDup = New TNFeInfNFeCobrDup
'                        a5.infNFe.cobr.dup(iIndice) = infNFeCobrDup

'                        a5.infNFe.cobr.dup(iIndice).nDup = objNFiscal.NumNotaFiscal & "/" & objParcPag.NumParcela
'                        a5.infNFe.cobr.dup(iIndice).dVenc = Format(objParcPag.DataVencimento, "yyyy-MM-dd")
'                        a5.infNFe.cobr.dup(iIndice).vDup = Replace(Format(objParcPag.Valor, "fixed"), ",", ".")
'                    Next

'                    Exit For
'                Next


'            ElseIf objNFiscal.ClasseDocCPR = 2 Then

'                resTitRec = db1.ExecuteQuery(Of TitulosRecTodo) _
'                ("SELECT * FROM TitulosRecTodos WHERE NumIntDoc = {0}", objNFiscal.NumIntDocCPR)

'                For Each objTitRec In resTitRec
'                    If objTitRec.CondicaoPagto = 1 Then
'                        a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item0
'                    Else
'                        a5.infNFe.ide.indPag = TNFeInfNFeIdeIndPag.Item1
'                    End If

'                    a5.infNFe.cobr.fat.nFat = objNFiscal.NumNotaFiscal
'                    a5.infNFe.cobr.fat.vOrig = Replace(Format(objTitRec.Valor, "fixed"), ",", ".")
'                    a5.infNFe.cobr.fat.vLiq = Replace(Format(objTitRec.Valor, "fixed"), ",", ".")

'                    resParcRec = db1.ExecuteQuery(Of ParcelasRecToda) _
'                    ("SELECT * FROM ParcelasRecTodas WHERE NumIntTitulo = {0}", objNFiscal.NumIntDocCPR)

'                    iIndice = -1

'                    Dim Dup(50) As TNFeInfNFeCobrDup

'                    a5.infNFe.cobr.dup = Dup

'                    For Each objParcRec In resParcRec
'                        iIndice = iIndice + 1

'                        Dim infNFeCobrDup As TNFeInfNFeCobrDup = New TNFeInfNFeCobrDup
'                        a5.infNFe.cobr.dup(iIndice) = infNFeCobrDup

'                        a5.infNFe.cobr.dup(iIndice).nDup = objNFiscal.NumNotaFiscal & "/" & objParcRec.NumParcela
'                        a5.infNFe.cobr.dup(iIndice).dVenc = Format(objParcRec.DataVencimento, "yyyy-MM-dd")
'                        a5.infNFe.cobr.dup(iIndice).vDup = Replace(Format(objParcRec.Valor, "fixed"), ",", ".")
'                    Next

'                    Exit For
'                Next
'            End If


'            a5.infNFe.Id = sideCUF & Format(objNFiscal.DataEmissao, "yyMM") & a5.infNFe.emit.Item
'            a5.infNFe.Id = a5.infNFe.Id & "55" & Format(CInt(a5.infNFe.ide.serie), "000")
'            a5.infNFe.Id = a5.infNFe.Id & Format(CLng(a5.infNFe.ide.nNF), "000000000")
'            a5.infNFe.Id = a5.infNFe.Id & Format(CLng(a5.infNFe.ide.cNF), "000000000")

'            Dim iDigito As Integer

'            CalculaDV_Modulo11(a5.infNFe.Id, iDigito)

'            a5.infNFe.Id = "NFe" & a5.infNFe.Id & iDigito
'            a5.infNFe.ide.cDV = iDigito



'            Dim AD As AssinaturaDigital = New AssinaturaDigital


'            Dim mySerializer As New XmlSerializer(GetType(TNFe))

'            XMLStream = New MemoryStream(10000)

'            mySerializer.Serialize(XMLStream, a5)

'            Dim xm As Byte()
'            xm = XMLStream.ToArray

'            XMLString = System.Text.Encoding.UTF8.GetString(xm)

'            AD.Assinar(XMLString, "infNFe", cert)

'            Dim xMlD As XmlDocument

'            xMlD = AD.XMLDocAssinado()

'            Dim xString As String
'            xString = AD.XMLStringAssinado



'            XMLStringNFes = XMLStringNFes & Mid(xString, 22) & " "

'            '****************  salva o arquivo 

'            XMLStreamDados = New MemoryStream(10000)

'            Dim xDados1 As Byte()

'            xDados1 = System.Text.Encoding.UTF8.GetBytes(Mid(xString, 22))

'            XMLStreamDados.Write(xDados1, 0, xDados1.Length)

'            Dim DocDados1 As XmlDocument = New XmlDocument

'            XMLStreamDados.Position = 0
'            DocDados1.Load(XMLStreamDados)
'            sArquivo = "c:\nfe\" & a5.infNFe.Id & ".xml"
'            DocDados1.Save(sArquivo)


'        Next


'        envioNFe.versao = "1.10"


'        Dim mySerializerw As New XmlSerializer(GetType(TEnviNFe))

'        XMLStream1 = New MemoryStream(10000)

'        mySerializerw.Serialize(XMLStream1, envioNFe)

'        Dim xmw As Byte()
'        xmw = XMLStream1.ToArray

'        XMLString1 = System.Text.Encoding.UTF8.GetString(xmw)

'        XMLString2 = Mid(XMLString1, 1, Len(XMLString1) - 10) & XMLStringNFes & Mid(XMLString1, Len(XMLString1) - 10)

'        XMLString2 = Mid(XMLString2, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString2, 20)



'        Dim XMLStringRetEnvNFE As String

'        'Load the client certificate from a file.
'        'Dim x509 As X509Certificate = X509Certificate.CreateFromSignedFile("c:\nfe\ecnpj.cer")

'        '************* valida dados antes do envio **********************
'        Dim xDados As Byte()

'        xDados = System.Text.Encoding.UTF8.GetBytes(XMLString2)

'        XMLStreamDados = New MemoryStream(10000)

'        XMLStreamDados.Write(xDados, 0, xDados.Length)


'        Dim DocDados As XmlDocument = New XmlDocument
'        XMLStreamDados.Position = 0
'        DocDados.Load(XMLStreamDados)
'        sArquivo = "c:\nfe\Lote" & envioNFe.idLote & ".xml"
'        DocDados.Save(sArquivo)

'        Call objValidaXML.validaXML(sArquivo, "c:\nfe\enviNFe_v1.10.xsd")

'        NfeRecepcao.ClientCertificates.Add(cert)

'        XMLStringRetEnvNFE = NfeRecepcao.nfeRecepcaoLote(XMLStringCabec, XMLString2)


'        Dim xRet As Byte()

'        xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetEnvNFE)

'        XMLStreamRet = New MemoryStream(10000)

'        XMLStreamRet.Write(xRet, 0, xRet.Length)

'        Dim mySerializerRetEnvNFe As New XmlSerializer(GetType(TRetEnviNFe))

'        Dim objRetEnviNFE As TRetEnviNFe = New TRetEnviNFe

'        XMLStreamRet.Position = 0

'        objRetEnviNFE = mySerializerRetEnvNFe.Deserialize(XMLStreamRet)

'        iResult = db1.ExecuteCommand("INSERT INTO NFeFedRetEnvi ( Lote, tpAmb, verAplic, versao, cStat, xMotivo, cUF, nRec, dhRecbto, tMed) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9} )", _
'        envioNFe.idLote, objRetEnviNFE.tpAmb, objRetEnviNFE.verAplic, objRetEnviNFE.versao, objRetEnviNFE.cStat, objRetEnviNFE.xMotivo, objRetEnviNFE.cUF, objRetEnviNFE.infRec.nRec, objRetEnviNFE.infRec.dhRecbto, objRetEnviNFE.infRec.tMed)

'        Dim objconsReciNFe As TConsReciNFe = New TConsReciNFe

'        objconsReciNFe.tpAmb = TAmb.Item2
'        objconsReciNFe.versao = "1.10"
'        objconsReciNFe.nRec = objRetEnviNFE.infRec.nRec

'        Dim mySerializerx As New XmlSerializer(GetType(TConsReciNFe))

'        XMLStream1 = New MemoryStream(10000)
'        mySerializerx.Serialize(XMLStream1, objconsReciNFe)

'        Dim xm1 As Byte()
'        xm1 = XMLStream1.ToArray

'        XMLString1 = System.Text.Encoding.UTF8.GetString(xm1)

'        XMLString1 = Mid(XMLString1, 1, 19) & " encoding=""utf-8"" " & Mid(XMLString1, 20)

'        Dim XMLStringRetConsReciNFE As String

'        NFeRetRecepCao.ClientCertificates.Add(cert)

'        XMLStringRetConsReciNFE = NFeRetRecepCao.nfeRetRecepcao(XMLStringCabec, XMLString1)

'        xRet = System.Text.Encoding.UTF8.GetBytes(XMLStringRetConsReciNFE)

'        XMLStreamRet = New MemoryStream(10000)
'        XMLStreamRet.Write(xRet, 0, xRet.Length)

'        Dim mySerializerRetConsReciNFe As New XmlSerializer(GetType(TRetConsReciNFe))

'        Dim objconsRetReciNFe As TRetConsReciNFe = New TRetConsReciNFe

'        XMLStreamRet.Position = 0

'        objconsRetReciNFe = mySerializerRetConsReciNFe.Deserialize(XMLStreamRet)

'        iResult = db1.ExecuteCommand("INSERT INTO NFeFedRetConsReci ( versao, tpAmb, verAplic, nRec, cStat, xMotivo, cUF) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6} )", _
'        objconsRetReciNFe.versao, objconsRetReciNFe.tpAmb, objconsRetReciNFe.verAplic, objconsRetReciNFe.nRec, objconsRetReciNFe.cStat, objconsRetReciNFe.xMotivo, objconsRetReciNFe.cUF)


'        For i = 0 To objconsRetReciNFe.protNFe.Count - 1

'            If String.IsNullOrEmpty(objconsRetReciNFe.protNFe(i).infProt.nProt) Then
'                objconsRetReciNFe.protNFe(i).infProt.nProt = ""
'            End If

'            objNFiscal = colNFiscal(i + 1)

'            iResult = db1.ExecuteCommand("INSERT INTO NFeFedProtNFe ( NumIntNF, versao, nRec, tpAmb, verAplic, chNFe, dHRecbto, nProt, cStat, xMotivo) VALUES ( {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9} )", _
'            objNFiscal.NumIntDoc, objconsRetReciNFe.protNFe(i).versao, objconsRetReciNFe.nRec, objconsRetReciNFe.protNFe(i).infProt.tpAmb, objconsRetReciNFe.protNFe(i).infProt.verAplic, objconsRetReciNFe.protNFe(i).infProt.chNFe, objconsRetReciNFe.protNFe(i).infProt.dhRecbto, objconsRetReciNFe.protNFe(i).infProt.nProt, objconsRetReciNFe.protNFe(i).infProt.cStat, objconsRetReciNFe.protNFe(i).infProt.xMotivo)

'        Next


'        'a5.infNFe.ide.NFref
'        'a5.infNFe.versao = "1.02"
'        'a5.infNFe.Id = "1"
'        'a5.infNFe.ide.cUF = TCodUfIBGE.Item33 'Rio de Janeiro
'        'a5.infNFe.ide.cNF = "000000001"
'        'a4.versao = "1.02"
'        'a4.versaoDados = "1.07"


'        'mySerializer.Serialize(instance, a4)
'        'i1 = instance.Position
'        'instance.Position = 0

'        'i1 = instance.Read(b1, 0, instance.Length)

'        'b2 = Encoding.Convert(System.Text.Encoding.ASCII, System.Text.Encoding.UTF8, b1)

'        's1 = System.Text.Encoding.ASCII.GetString(b2)


'        ''z.Body = x

'        ''x.nfeCabecMsg = "xyz"
'        ''x.nfeDadosMsg = "abc"
'        ''arr1 = coll.AllKeys
'        ''For loop1 = 0 To arr1.GetUpperBound(0)
'        ''    Response.Write("Key: " & Server.HtmlEncode(arr1(loop1)) & "<br>")
'        ''     Get all values under this key.
'        ''    arr2 = coll.GetValues(loop1)
'        ''    For loop2 = 0 To arr2.GetUpperBound(0)
'        ''        Response.Write("Value " & CStr(loop2) & ": " & Server.HtmlEncode(arr2(loop2)) & "<br><br>")
'        ''    Next loop2
'        ''Next loop1


'        ''col1.Add("m1", "1")
'        ''col1.Add("t2", "2")

'        ''HyperLink1.NavigateUrl = "~/Default2.aspx?m1=1&m2=2"
'        ''Request.QueryString
'        ''Request.QueryString.Add(col1)

'        db1.Transaction.Commit()



'    Catch ex As Exception
'        db1.Transaction.Rollback()

'    End Try


'    'Erro_Button1_Click:


'End Sub

''Public Sub validaXML(ByVal _arquivo As String, ByVal _schema As String)

''    ' Create a new validating reader

''    Dim reader As XmlValidatingReader = New XmlValidatingReader(New XmlTextReader(New StreamReader(_arquivo)))
''    'Dim reader As XmlValidatingReader = New XmlValidatingReader()
''    '        Dim reader1 As XmlWriter
''    '       reader1.
''    Dim schema(1) As System.Xml.Schema.XmlSchema

''    '// Create a schema collection, add the xsd to it

''    Dim schemaCollection As XmlSchemaSet = New XmlSchemaSet()

''    schemaCollection.Add("http://www.portalfiscal.inf.br/nfe", _schema)

''    schemaCollection.CopyTo(schema, 0)

''    '// Add the schema collection to the XmlValidatingReader

''    reader.Schemas.Add(schema(0))

''    '       Console.Write("Início da validação...\n")

''    '    // Wire up the call back.  The ValidationEvent is fired when the
''    '    // XmlValidatingReader hits an issue validating a section of the xml

''    '            reader. += new ValidationEventHandler(reader_ValidationEventHandler);
''    AddHandler reader.ValidationEventHandler, AddressOf reader_ValidationEventHandler

''    '            // Iterate through the xml document



''    '            while (reader.Read()) {}
''    While reader.Read()
''    End While


''    '          Console.WriteLine("\rFim de validação\n");
''    'Console.ReadLine();
''End Sub

''Sub reader_ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)

''    '            // Report back error information to the console...
''    MessageBox.Show(e.Exception.Message)
''    '        Console.WriteLine("\rLinha:{0} Coluna:{1} Erro:{2} Name:[3} Valor:{4}\r", e.Exception.LinePosition, e.Exception.LineNumber, e.Exception.Message, sender.Name, sender.Value)


''End Sub
'Sub CalculaDV_Modulo11(ByVal sString As String, ByRef iDigito As Integer)
'    Dim iIndice As Integer
'    Dim iMult As Integer
'    Dim iTotal As Integer

'    iMult = 2

'    For iIndice = Len(sString) To 1 Step -1

'        iTotal = iTotal + (Mid(sString, iIndice, 1) * iMult)

'        If iMult = 9 Then
'            iMult = 2
'        Else
'            iMult = iMult + 1
'        End If

'    Next

'    iDigito = iTotal Mod 11

'    iDigito = 11 - iDigito

'    If iDigito > 9 Then iDigito = 0

'End Sub

'Function PIS_CST(ByRef iCST As Integer, ByVal objTributacaoItemNF As TributacaoItemNF) As Long
'    If objTributacaoItemNF.PISCredito > 0 Then
'        iCST = 1
'    Else
'        iCST = 4
'    End If
'    PIS_CST = SUCESSO
'End Function

'Function PIS_Aliquota(ByRef dAliquota As Double, ByVal objFilialEmpresa As FiliaisEmpresa) As Long
'    If objFilialEmpresa.PISNaoCumulativo = 1 Then
'        dAliquota = 0.0165
'    Else
'        dAliquota = 0.0065
'    End If
'    PIS_Aliquota = SUCESSO
'End Function

'Function COFINS_CST(ByRef iCST As Integer, ByVal objTributacaoItemNF As TributacaoItemNF) As Long
'    If objTributacaoItemNF.COFINSCredito > 0 Then
'        iCST = 1
'    Else
'        iCST = 4
'    End If
'    COFINS_CST = SUCESSO
'End Function

'Function COFINS_Aliquota(ByRef dAliquota As Double, ByVal objFilialEmpresa As FiliaisEmpresa) As Long
'    If objFilialEmpresa.COFINSNaoCumulativo = 1 Then
'        dAliquota = 0.076
'    Else
'        dAliquota = 0.03
'    End If
'    COFINS_Aliquota = SUCESSO
'End Function

'Sub Armazena_Estado(ByVal SiglaEstado As Integer, ByRef UF As TUf)

'    Select Case SiglaEstado

'        Case "AC"
'            UF = TUf.AC

'        Case "AL"
'            UF = TUf.AL

'        Case "AM"
'            UF = TUf.AM

'        Case "AP"
'            UF = TUf.AP

'        Case "BA"
'            UF = TUf.BA

'        Case "CE"
'            UF = TUf.CE

'        Case "DF"
'            UF = TUf.DF

'        Case "ES"
'            UF = TUf.ES

'        Case "GO"
'            UF = TUf.GO

'        Case "MA"
'            UF = TUf.MA

'        Case "MG"
'            UF = TUf.MG

'        Case "MS"
'            UF = TUf.MS

'        Case "MT"
'            UF = TUf.MT

'        Case "PA"
'            UF = TUf.PA

'        Case "PB"
'            UF = TUf.PB

'        Case "PE"
'            UF = TUf.PE

'        Case "PI"
'            UF = TUf.PI

'        Case "PR"
'            UF = TUf.PR

'        Case "RJ"
'            UF = TUf.RJ

'        Case "RN"
'            UF = TUf.RN

'        Case "RO"
'            UF = TUf.RO

'        Case "RR"
'            UF = TUf.RR

'        Case "RS"
'            UF = TUf.RS

'        Case "SC"
'            UF = TUf.SC

'        Case "SE"
'            UF = TUf.SE

'        Case "SP"
'            UF = TUf.SP

'        Case "TO"
'            UF = TUf.TO

'    End Select
'End Sub







'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
'    Dim x As ClassCancelaNFe = New ClassCancelaNFe
'    Dim lErro As Long
'    Dim lNumIntNF As Long

'    Try

'        lNumIntNF = NumIntNF.Text
'        lErro = x.Cancela_NFe(lNumIntNF, "cancelado pelo usuario")

'    Catch ex As Exception
'        MessageBox.Show(ex.Message, "Erro")

'    End Try

'End Sub
