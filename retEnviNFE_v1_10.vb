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
 System.Xml.Serialization.XmlRootAttribute("retEnviNFe", [Namespace]:="http://www.portalfiscal.inf.br/nfe", IsNullable:=false)>  _
Partial Public Class TRetEnviNFe
    
    Private tpAmbField As TAmb
    
    Private verAplicField As String
    
    Private cStatField As String
    
    Private xMotivoField As String
    
    Private cUFField As String
    
    Private infRecField As TRetEnviNFeInfRec
    
    Private versaoField As String
    
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
    Public Property verAplic() As String
        Get
            Return Me.verAplicField
        End Get
        Set
            Me.verAplicField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property cStat() As String
        Get
            Return Me.cStatField
        End Get
        Set
            Me.cStatField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property xMotivo() As String
        Get
            Return Me.xMotivoField
        End Get
        Set
            Me.xMotivoField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property cUF() As String
        Get
            Return Me.cUFField
        End Get
        Set(ByVal value As String)
            Me.cUFField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property infRec() As TRetEnviNFeInfRec
        Get
            Return Me.infRecField
        End Get
        Set
            Me.infRecField = value
        End Set
    End Property
    
    '''<remarks/>
    <System.Xml.Serialization.XmlAttributeAttribute()>  _
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
Partial Public Class TRetEnviNFeInfRec
    
    Private nRecField As String
    
    Private dhRecbtoField As Date
    
    Private tMedField As String
    
    '''<remarks/>
    Public Property nRec() As String
        Get
            Return Me.nRecField
        End Get
        Set
            Me.nRecField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property dhRecbto() As Date
        Get
            Return Me.dhRecbtoField
        End Get
        Set
            Me.dhRecbtoField = value
        End Set
    End Property
    
    '''<remarks/>
    Public Property tMed() As String
        Get
            Return Me.tMedField
        End Get
        Set
            Me.tMedField = value
        End Set
    End Property
End Class
