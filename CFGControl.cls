VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFGControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const NumeroVariables = 6


'Declaracion de las variables
'---- Servidor donde esta MYSQL
Private mvarSERVER As String 'Variable local
'---- Numero de contabilidades maximo

Private mvarUser As String 'Variable local
'---- General
Private mvarpassword As String 'Variable local


Private mvarColImporte As Integer 'Variable local
Private mvarPlantilla As String 'plantilla
Private mvarPlantillaHoras As String 'plantilla horas


Dim Mens As String
Dim NF As Integer


Public Property Let SERVER(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarSERVER = vData
End Property

Public Property Get SERVER() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     SERVER = mvarSERVER
End Property




'----
'---- Usuario general BD
'----
Public Property Let User(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.User = 5
     mvarUser = vData
End Property

Public Property Get User() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.User
     User = mvarUser
End Property




'----
'---- General
'----
Public Property Let password(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarpassword = vData
End Property

Public Property Get password() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     password = mvarpassword
End Property


Public Property Let ColImporte(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.SERVER = 5
     mvarColImporte = vData
End Property

Public Property Get ColImporte() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.SERVER
     ColImporte = mvarColImporte
End Property


Public Property Let Plantilla(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarPlantilla = vData
End Property

Public Property Get Plantilla() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     Plantilla = mvarPlantilla
End Property

Public Property Let PlantillaHoras(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarPlantillaHoras = vData
End Property

Public Property Get PlantillaHoras() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     PlantillaHoras = mvarPlantillaHoras
End Property


Public Function Guardar() As Byte
    Dim NF As Integer
    Dim F As Integer
    Dim Leido As String
    Dim Linea As String
    Dim Param As String
    Dim Valor As String
    Dim I As Integer
    Dim vec(NumeroVariables) As Boolean
        'Para saber si algun parametro no se ha puesto
        'El v(0) sera la BD y el v(4) el de NomFich
        'Cuando acabemos de pasar el fichero comprobaremos si hay que poner todos los valores
          
    For I = 0 To NumeroVariables
        vec(I) = False
    Next I
    
    Guardar = 1
    On Error GoTo ErrGuardar
    F = FreeFile
    '-- Control de existencia del fichero
    If Dir(App.Path & "\ControlNomi.cfg") = "" Then
        Open App.Path & "\ControlNomi.cfg" For Output As #F   ' Abre el archivo.
        Linea = " Server = " & mvarSERVER
        Print #F, Linea
        Linea = " User = " & Encripta(mvarUser)
        Print #F, Linea
        Linea = " Password = " & Encripta(mvarpassword)
        Print #F, Linea
        Linea = " ColImporte = " & mvarColImporte
        Print #F, Linea
        Linea = " Plantilla = " & mvarPlantilla
        Print #F, Linea
        Linea = " Plantilla = " & mvarPlantillaHoras
        Print #F, Linea
        
        Close #NF
    Else
        NF = F + 1
        Open App.Path & "\ControlNomi.cfg" For Input As #NF   ' Abre el archivo.
        Open App.Path & "\Control2.cfg" For Output As #F   ' Abre el archivo.
        Do While Not EOF(NF)   ' Repite el bucle hasta el final del archivo.
            Line Input #NF, Leido   ' Lee el registro en la variable.
            Leido = Trim(Leido)
            If Leido = "" Or Left(Leido, 1) = "#" Then ' Se desprecian las líneas en blanco y comentarios
                'Tiene comentarios o es linea en blanco
                Linea = Leido
                'ELSE
                Else
                    I = InStr(1, Leido, "=")
                    If I > 0 Then
                        Param = Trim(Mid(Leido, 1, I - 1))
                        Valor = Trim(Mid(Leido, I + 1, Len(Leido) - I))
                        Select Case Param
                            Case "ColImporte"
                                Linea = " ColImporte = " & mvarColImporte
                                vec(0) = True
                            Case "Server"
                                Linea = " Server = " & mvarSERVER
                                vec(1) = True
                            Case "User"
                                Linea = " User = " & Encripta(mvarUser)
                                vec(2) = True
                            Case "Password"
                                Linea = " Password = " & Encripta(mvarpassword)
                                vec(3) = True
                            Case "Plantilla"
                                Linea = " Plantilla = " & mvarPlantilla
                                vec(4) = True
                            Case "PlantillaHoras"
                                Linea = " PlantillaHoras = " & mvarPlantillaHoras
                                vec(5) = True
                            Case Else
                                '-- No se reconoce el parámetro
                                Linea = "# " & Leido
                          End Select
                        'ELSE i>0
                        Else
                            '-- falta el signo de igualdad en el parámetro
                            'Comentamos la linea
                            Linea = "# " & Leido
                    End If
            End If
            'Escribimos en F
            Print #F, Linea
        Loop
        
        
        'Ahora comprobamos si todos los valores han sido colocados
        For I = 0 To NumeroVariables
        
            If vec(I) = False Then
                Select Case I
                    Case 0     '"ColImporte"
                        If mvarColImporte <> "" Then _
                        Linea = " ColImporte = " & mvarColImporte
                    Case 1   '"Server"
                        If mvarSERVER <> "" Then _
                        Linea = " Server = " & mvarSERVER
                    Case 2       '"User"
                        If mvarUser <> "" Then _
                        Linea = " User = " & Encripta(mvarUser)
                    Case 3      '"NomFich"
                        If mvarpassword <> "" Then _
                        Linea = " Password = " & Encripta(mvarpassword)
                    Case 4   '"Plantilla"
                        If mvarPlantilla <> "" Then _
                        Linea = " Plantilla = " & mvarPlantilla
                    Case 5   '"PlantillaHoras"
                        If mvarPlantillaHoras <> "" Then _
                        Linea = " PlantillaHoras = " & mvarPlantillaHoras
                End Select
                Print #F, Linea
            End If
        Next I
        Close #NF   ' Cierra el archivo.
        Close #F
        'Ahora, con los archivo cerrados procedemos a cambiar los valores
        FileCopy App.Path & "\ControlNomi.cfg", App.Path & "\ControlNomi.BAK"
        FileCopy App.Path & "\Control2.cfg", App.Path & "\ControlNomi.cfg"
        Kill App.Path & "\Control2.cfg"
    End If
    Guardar = 0
Exit Function
ErrGuardar:
    MsgBox "No se ha podido guardar la configuración. " & vbCrLf & "Número: " & Err.Number & vbCrLf & _
        "Descripción: " & Err.Description, vbExclamation
End Function



Public Function Leer() As Byte
  '-- Al inicializarse la clase se cargan los valores de parámetros correspon
    '   dientes, se comprueba igualmente que el fichero de configuración exista
    Dim Leido As String
    Dim Param As String
    Dim Valor As String
    Dim I As Integer
    
    
    '-- Control de existencia del fichero
    If Dir(App.Path & "\ControlNomi.cfg") = "" Then
        ' El fichero no existe
        Leer = 1
        MsgBox "Falta el fichero de configuración " & App.Path & "\ControlNomi.cfg", vbCritical ', "Club Carrefour"
    Else
        NF = FreeFile
        Open App.Path & "\ControlNomi.cfg" For Input As #NF   ' Abre el archivo.
        Do While Not EOF(NF)   ' Repite el bucle hasta el final del archivo.
            Line Input #NF, Leido   ' Lee el registro en la variable.
            If Leido <> "" Or Left(Leido, 1) <> "#" Then ' Se desprecian las líneas en blanco y comentarios
                I = InStr(1, Leido, "=")
                If I > 0 Then
                    Param = Trim(Mid(Leido, 1, I - 1))
                    Valor = Trim(Mid(Leido, I + 1, Len(Leido) - I))
                    Select Case Param
                        Case "ColImporte"
                            mvarColImporte = Valor
                        Case "Server"
                            mvarSERVER = Valor
                        Case "User"
                            mvarUser = Valor
                        Case "Password"
                            mvarpassword = Valor
                        Case "Plantilla"
                            mvarPlantilla = Valor
                        Case "PlantillaHoras"
                            mvarPlantillaHoras = Valor
                        Case Else
                            '-- No se reconoce el parámetro
                    End Select
                Else
                    '-- falta el signo de igualdad en el parámetro
                End If
                                     
            End If
        Loop
        Close #NF   ' Cierra el archivo.
        'Desencriptamos password
        Desencripta (True)
        Desencripta (False)
        
        Leer = 0
    End If

End Function

Private Sub Desencripta(ElPassword As Boolean)
Dim C As String
Dim I As Integer

If ElPassword Then
    C = mvarpassword
    Else
    C = mvarUser
End If
Mens = ""
For I = 1 To Len(C)
    If (I Mod 2) <> 0 Then
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, I, 1))
        NF = NF + 1
        Mens = Mens & Chr(NF)
    End If
Next I

If ElPassword Then
    mvarpassword = Mens
    Else
    mvarUser = Mens
End If

End Sub



Private Function Encripta(ByRef C As String) As String
Dim I As Integer


Mens = ""
Randomize
For I = 1 To Len(C)
        'Solo valen los caracteres impares
        NF = Asc(Mid(C, I, 1))
        NF = NF - 1
        Mens = Mens & Chr(NF)
        'Y el aleatorio , entre 48 y 122  <- 74 ->
        NF = Int((74 * Rnd)) ' Genera valores aleatorios entre 0y74.
        NF = NF + 48
        Mens = Mens & Chr(NF)
Next I

Encripta = Mens

End Function

