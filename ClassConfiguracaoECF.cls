VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConfiguracaoECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsDirPAFECF As String
Private mvarsPrincExec As String
Private mvarsVersaoEspReqPAFECF As String
Private mvarsNumLaudo As String
Private mvarsContatoDesenv As String
Private mvarsTelDesenv As String
Private mvarsEndDesenv As String
Private mvarsVersaoPAFECF As String
Private mvarsNomePAFECF As String
Private mvarsRazaoSocialDesenv As String
Private mvarsIMDesenv As String
Private mvarsIEDesenv As String
Private mvarsCNPJDesenv As String


Public Property Get sCNPJDesenv() As String
    sCNPJDesenv = mvarsCNPJDesenv
End Property

Public Property Let sCNPJDesenv(ByVal vData As String)
    mvarsCNPJDesenv = vData
End Property

Public Property Get sIEDesenv() As String
    sIEDesenv = mvarsIEDesenv
End Property

Public Property Let sIEDesenv(ByVal vData As String)
    mvarsIEDesenv = vData
End Property

Public Property Get sIMDesenv() As String
    sIMDesenv = mvarsIMDesenv
End Property

Public Property Let sIMDesenv(ByVal vData As String)
    mvarsIMDesenv = vData
End Property

Public Property Get sRazaoSocialDesenv() As String
    sRazaoSocialDesenv = mvarsRazaoSocialDesenv
End Property

Public Property Let sRazaoSocialDesenv(ByVal vData As String)
    mvarsRazaoSocialDesenv = vData
End Property

Public Property Get sNomePAFECF() As String
    sNomePAFECF = mvarsNomePAFECF
End Property

Public Property Let sNomePAFECF(ByVal vData As String)
    mvarsNomePAFECF = vData
End Property

Public Property Get sVersaoPAFECF() As String
    sVersaoPAFECF = mvarsVersaoPAFECF
End Property

Public Property Let sVersaoPAFECF(ByVal vData As String)
    mvarsVersaoPAFECF = vData
End Property

Public Property Get sEndDesenv() As String
    sEndDesenv = mvarsEndDesenv
End Property

Public Property Let sEndDesenv(ByVal vData As String)
    mvarsEndDesenv = vData
End Property

Public Property Get sTelDesenv() As String
    sTelDesenv = mvarsTelDesenv
End Property

Public Property Let sTelDesenv(ByVal vData As String)
    mvarsTelDesenv = vData
End Property

Public Property Get sContatoDesenv() As String
    sContatoDesenv = mvarsContatoDesenv
End Property

Public Property Let sContatoDesenv(ByVal vData As String)
    mvarsContatoDesenv = vData
End Property

Public Property Get sNumLaudo() As String
    sNumLaudo = mvarsNumLaudo
End Property

Public Property Let sNumLaudo(ByVal vData As String)
    mvarsNumLaudo = vData
End Property

Public Property Get sVersaoEspReqPAFECF() As String
    sVersaoEspReqPAFECF = mvarsVersaoEspReqPAFECF
End Property

Public Property Let sVersaoEspReqPAFECF(ByVal vData As String)
    mvarsVersaoEspReqPAFECF = vData
End Property

Public Property Get sPrincExec() As String
    sPrincExec = mvarsPrincExec
End Property

Public Property Let sPrincExec(ByVal vData As String)
    mvarsPrincExec = vData
End Property

Public Property Get sDirPAFECF() As String
    sDirPAFECF = mvarsDirPAFECF
End Property

Public Property Let sDirPAFECF(ByVal vData As String)
    mvarsDirPAFECF = vData
End Property



