VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsTitulo As String
Private mvariTabelaDinamica As Integer
Private mvariIncluirGrafico As Integer
Private mvariTipoGrafico As Integer
Private mvarcolCampos As New Collection
Private mvarcolFormulas As New Collection
Private mvariFormato As Integer
Private mvarsLocalizacaoCsv As String
Private mvarsArquivo As String

Public Property Let sArquivo(ByVal vData As String)
    mvarsArquivo = vData
End Property

Public Property Get sArquivo() As String
    sArquivo = mvarsArquivo
End Property

Public Property Let sLocalizacaoCsv(ByVal vData As String)
    mvarsLocalizacaoCsv = vData
End Property

Public Property Get sLocalizacaoCsv() As String
    sLocalizacaoCsv = mvarsLocalizacaoCsv
End Property

Public Property Let iFormato(ByVal vData As Integer)
    mvariFormato = vData
End Property

Public Property Get iFormato() As Integer
    iFormato = mvariFormato
End Property

Public Property Set colFormulas(ByVal vData As Object)
    Set mvarcolFormulas = vData
End Property

Public Property Get colFormulas() As Collection
    Set colFormulas = mvarcolFormulas
End Property

Public Property Set colCampos(ByVal vData As Object)
    Set mvarcolCampos = vData
End Property

Public Property Get colCampos() As Collection
    Set colCampos = mvarcolCampos
End Property

Public Property Let iTipoGrafico(ByVal vData As Integer)
    mvariTipoGrafico = vData
End Property

Public Property Get iTipoGrafico() As Integer
    iTipoGrafico = mvariTipoGrafico
End Property

Public Property Let iIncluirGrafico(ByVal vData As Integer)
    mvariIncluirGrafico = vData
End Property

Public Property Get iIncluirGrafico() As Integer
    iIncluirGrafico = mvariIncluirGrafico
End Property

Public Property Let iTabelaDinamica(ByVal vData As Integer)
    mvariTabelaDinamica = vData
End Property

Public Property Get iTabelaDinamica() As Integer
    iTabelaDinamica = mvariTabelaDinamica
End Property

Public Property Let sTitulo(ByVal vData As String)
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
    sTitulo = mvarsTitulo
End Property




