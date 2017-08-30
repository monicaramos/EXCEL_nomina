Attribute VB_Name = "libGestoria"
Option Explicit


Public mConfig As CFGControl
Public Conn As Connection


Public MiXL As Object  ' Variable que contiene la referencia
    ' de Microsoft Excel.
Public ExcelNoSeEjecutaba As Boolean   ' Indicador para liberación final .
Public ExcelSheet As Object
Public wrk As Excel.Workbook

Public BaseDatos As String

Public EsImportaci As Byte
Public NombreHoja As String
Dim Rc As Byte


Public Usuario As Long


Public Sub Main()
Dim I As Integer
'Vemos si ya se esta ejecutando
If App.PrevInstance Then
    MsgBox "Ya se está ejecutando el programa de traspaso a Excel (Tenga paciencia).", vbCritical
    Screen.MousePointer = vbDefault
    Exit Sub
End If


Set mConfig = New CFGControl
If mConfig.Leer = 1 Then
    MsgBox "No configurado"
    End
End If

'Si es importacion o creacion
NombreHoja = Command '"/I ariagro1"
'NombreHoja = "/I|ariagro4|22000|" '"/P|ariagro4|22000|"


I = InStr(1, NombreHoja, "/")
If I = 0 Then
    MsgBox "Mal lanzado el programa", vbExclamation
    End
End If

NombreHoja = Mid(NombreHoja, I + 1)

Select Case Mid(NombreHoja, 1, 1)
    Case "C"
        EsImportaci = 0
    
    Case "I"
        EsImportaci = 1
    
    Case "D"
        EsImportaci = 3 ' exportacion de las tablas de horas y horas destajo
    
    Case "P"
        EsImportaci = 4 ' exportacion de horas trabajadas Picassent
    
    Case Else
        EsImportaci = 2 ' exportacion de horas trabajadas por trabajador en un mes
                        ' Informe de asesoria
End Select

'BaseDatos = Mid(NombreHoja, 3, Len(NombreHoja))
BaseDatos = RecuperaValor(NombreHoja, 2)
If BaseDatos = "" Then
    MsgBox "Falta la base de datos", vbCritical
    End
End If

If EsImportaci = 2 Then
    If Dir(App.Path & "\" & mConfig.Plantilla, vbArchive) <> mConfig.Plantilla Then
        MsgBox "Falta la plantilla, para realizar la exportacion."
        End
    End If
End If

If EsImportaci = 1 Or EsImportaci = 2 Or EsImportaci = 4 Then
    Usuario = RecuperaValor(NombreHoja, 3)
End If

NombreHoja = ""

frmClasifica.Show


End Sub

Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim cont As Integer
Dim cad As String

    I = 0
    cont = 1
    cad = ""
    Do
        J = I + 1
        I = InStr(J, Cadena, "|")
        If I > 0 Then
            If cont = Orden Then
                cad = Mid(Cadena, J, I - J)
                I = Len(Cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function

Public Function AbrirConexion(BaseDatos As String) As Boolean
Dim cad As String

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
                       
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & Trim(BaseDatos) & ";SERVER=" & mConfig.SERVER & ";"
    cad = cad & ";UID=" & mConfig.User
    cad = cad & ";PWD=" & mConfig.password
'++monica: tema del vista
    cad = cad & ";Persist Security Info=true"
    
    Conn.ConnectionString = cad
    Conn.Open
    If Err.Number <> 0 Then
        MsgBox "Error en la cadena de conexion" & vbCrLf & BaseDatos, vbCritical
        End
    Else
        AbrirConexion = True
    End If
End Function



Public Function DevuelveValor(vSQL As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not Rs.EOF Then
        ' antes RS.Fields(0).Value > 0
        If Not IsNull(Rs.Fields(0).Value) Then DevuelveValor = Rs.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
    End If
End Function

'''
'''Mezcala una linea del fichero de texto sobre el fichero de EXCEL
'''Public Function MezclaFicheros(ValorLinea As String, NumeroLinea As Integer) As Byte
'''Dim Col As Integer
'''Dim aux As String
'''Dim j As Integer
'''Dim inicio As Integer
'''Dim Cadena As String
'''
'''
'''On Error GoTo ErrorMezcla
'''MezclaFicheros = 1
'''inicio = 1
'''Col = 2  'Porque empieza en la 2
'''Do
'''    j = InStr(inicio, ValorLinea, "|")
'''    If j > 0 Then
'''        Cadena = Mid(ValorLinea, inicio, j - inicio)
'''        aux = ComasAPuntos(Cadena)
'''        ExcelSheet.Cells(NumeroLinea, Col) = aux
'''        inicio = j + 1
'''        Col = ColumnaProxima(Col)
'''    End If
'''Loop Until j = 0
'''
'''MezclaFicheros = 0
'''Exit Function
'''ErrorMezcla:
'''
'''End Function
