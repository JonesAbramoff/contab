﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:2.0.50727.1434
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System.Xml.Serialization

'
'This source code was auto-generated by xsd, Version=2.0.50727.1432.
'

'''<remarks/>
<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"),  _
 System.SerializableAttribute(),  _
 System.Diagnostics.DebuggerStepThroughAttribute(),  _
 System.ComponentModel.DesignerCategoryAttribute("code"),  _
 System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.portalfiscal.inf.br/nfe"),  _
 System.Xml.Serialization.XmlRootAttribute("cancNFe", [Namespace]:="http://www.portalfiscal.inf.br/nfe", IsNullable:=false)>  _
Partial Public Class TCancNFe
    
    Private infCancField As TCancNFeInfCanc
    
    Private versaoField As String
    
    '''<remarks/>
    Public Property infCanc() As TCancNFeInfCanc
        Get
            Return Me.infCancField
        End Get
        Set
            Me.infCancField = value
        End Set
    End Property
    
    '''<remarks/>
    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="token")>  _
    Public Property versao() As String
        Get
            Return Me.versaoField
        End Get
        Set
            Me.versaoField = value
        End Set
    End Property
End Class

'''<remarks/>
<System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.1432"),  _
 System.SerializableAttribute(),  _
 System.Diagnostics.DebuggerStepThroughAttribute(),  _
 System.ComponentModel.DesignerCategoryAttribute("code"),  _
 System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=true, [Namespace]:="http://www.portalfiscal.inf.br/nfe")>  _
Partial Public Class TCancNFeInfCanc
    
    Private tpAmbField As TAmb
    
    Private xServField As String
    
    Private chNFeField As String
    
    Private nProtField As String
    
    Private xJustField As String
    
    Private idField As String
    
    Public Sub New()
        MyBase.New
        Me.xServField = "CANCELAR"
    End Sub
    
    '''<remarks/>
    Public Property tpAmb() As TAmb
        Get
            Return Me.tpAmbField
        End Get
        Set
            Me.tpAmbField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property xServ() As String
        Get
            Return Me.xServField
        End Get
        Set
            Me.xServField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property chNFe() As String
        Get
            Return Me.chNFeField
        End Get
        Set
            Me.chNFeField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property nProt() As String
        Get
            Return Me.nProtField
        End Get
        Set
            Me.nProtField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property xJust() As String
        Get
            Return Me.xJustField
        End Get
        Set
            Me.xJustField = value
        End Set
    End Property
    
    '''<remarks/>
    <System.Xml.Serialization.XmlAttributeAttribute(DataType:="ID")>  _
    Public Property Id() As String
        Get
            Return Me.idField
        End Get
        Set
            Me.idField = value
        End Set
    End Property
End Class

