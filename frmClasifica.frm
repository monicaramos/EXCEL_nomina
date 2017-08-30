VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmClasifica 
   Caption         =   "Generacion archivo clasificacion"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmClasifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCompletar 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7395
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   90
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   360
         Width           =   7155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   32
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   31
         Top             =   1800
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   90
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1050
         Width           =   7155
      End
      Begin ComctlLib.ProgressBar Pb2 
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   1530
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1350
         Picture         =   "frmClasifica.frx":1782
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero nominas"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1530
         Picture         =   "frmClasifica.frx":1884
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fichero generado"
         Height          =   195
         Left            =   90
         TabIndex        =   33
         Top             =   810
         Width           =   1395
      End
   End
   Begin VB.Frame FrameConfig 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text8 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2790
         PasswordChar    =   "*"
         TabIndex        =   15
         Text            =   "Text8"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   2
         Left            =   2790
         TabIndex        =   14
         Text            =   "Text8"
         Top             =   930
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   1
         Left            =   2790
         TabIndex        =   13
         Text            =   "Text8"
         Top             =   570
         Width           =   1515
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   0
         Left            =   2790
         TabIndex        =   17
         Text            =   "Text8"
         Top             =   1650
         Width           =   1485
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   21
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Máximo de Calidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   1710
         Width           =   2145
      End
      Begin VB.Label Label7 
         Caption         =   "CLASIFICACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   240
         TabIndex        =   16
         Top             =   1350
         Width           =   825
      End
   End
   Begin VB.Frame FrameImportar 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1380
         Width           =   5385
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Importar"
         Height          =   375
         Left            =   4500
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   690
         Width           =   6735
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   960
         Picture         =   "frmClasifica.frx":1986
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Concepto"
         Height          =   195
         Left            =   1530
         TabIndex        =   25
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   450
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   900
         Picture         =   "frmClasifica.frx":1A11
         Top             =   420
         Width           =   240
      End
   End
   Begin VB.Frame FrameEscribir 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin ComctlLib.ProgressBar Pb1 
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   1170
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   690
         Width           =   7155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   1
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Fichero generado"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   450
         Width           =   1395
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1530
         Picture         =   "frmClasifica.frx":1B13
         Top             =   450
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7260
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClasifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private WithEvents frmC As frmCal
Private NoEncontrados As String




Dim SQL As String
Dim VariasEntradas As String


Dim Albaran As Long
Dim FechaHora As String
Dim Socio As String
Dim Campo As String
Dim Variedad As String
Dim TipoEntr As String
Dim KilosNet As String
Dim Calidad(20) As String
Dim Importe As Currency
Dim Trabajador As String
Dim Concepto As String
Dim CodigoAsesoria As String

Private WithEvents frmMens As frmMensajes 'Registros que no ha entrado con error
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1


Private Sub cmdConfig_Click(Index As Integer)
Dim I As Integer

    If Index = 1 Then
        Unload Me
    Else
        SQL = ""
        For I = 0 To Text8.Count - 1
            If Text8(I).Text = "" Then SQL = SQL & "Campo: " & I & vbCrLf
        Next I
        If SQL <> "" Then
            SQL = "No pueden haber campos vacios: " & vbCrLf & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            Text8(0).SetFocus
            Exit Sub
        End If
        
        'mConfig.MaxCalidades = Text8(0).Text
        mConfig.SERVER = Text8(1).Text
        mConfig.User = Text8(2).Text
        mConfig.password = Text8(3).Text
        
        mConfig.Guardar
        
        vConfiguracion False
'        If varConfig.Grabar = 0 Then End
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Rc As Byte
Dim Mens As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
        
    
    If Text2.Text <> "" Then
        If Dir(Text2.Text) <> "" Then
            MsgBox "Fichero ya existe", vbExclamation
            Exit Sub
        Else
            Select Case EsImportaci
                Case 2 ' exportacion del listado de asesoria
                    FileCopy App.Path & "\" & mConfig.Plantilla, Text2.Text
                Case 3 ' exportacion del fichero de horas
                    FileCopy App.Path & "\" & mConfig.PlantillaHoras, Text2.Text
            End Select
            NombreHoja = Text2.Text
        End If
    End If
    
    'Abrimos excel
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        'Si queremos que se vea descomentamos  esto
        MiXL.Application.visible = False
'        MiXL.Parent.Windows(1).Visible = False
    
        'Realizamos todos los datos
        'abrimos conexion
        If AbrirConexion(BaseDatos) Then
        
            Screen.MousePointer = vbHourglass
            
            'Vamos linea a linea
            Mens = "Error insertando en Excel"
            Select Case EsImportaci
                Case 2
                    If Not RecorremosLineas(Mens) Then
                        MsgBox Mens, vbExclamation
                    End If
                Case 3
                    If Not RecorremosLineasHoras(Mens) Then
                        MsgBox Mens, vbExclamation
                    End If
            End Select
            
            Screen.MousePointer = vbDefault
            
        End If
    
        'Cerramos el excel
        CerrarExcel
                
        MsgBox "Proceso finalizado", vbExclamation


    End If

    
End Sub

Private Sub Command2_Click()
Dim Rc As Byte
Dim I As Integer
Dim Rs1 As ADODB.Recordset
Dim KilosI As Long
Dim vSQL As String


    'IMPORTAR


    If Text5.Text = "" Then
        MsgBox "Escriba el nombre del fichero excel", vbExclamation
        Exit Sub
    End If
        
        
    If Dir(Text5.Text) = "" Then
        MsgBox "Fichero no existe"
        Exit Sub
    End If
    
    If Text4.Text = "" Then
        MsgBox "Escriba la fecha", vbExclamation
        Exit Sub
    End If
    
    If Text7.Text = "" Then
        MsgBox "Ponga la descripción del concepto", vbExclamation
        Exit Sub
    End If
    
    NombreHoja = Text5.Text
    
    'Abrimos excel
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        'Realizamos todos los datos
        'abrimos conexion
        If AbrirConexion(BaseDatos) Then
        
            
            'Vamos linea a linea, buscamos su trabajador
            RecorremosLineasLiquidacion
            
        End If
    
        'Cerramos el excel
        CerrarExcel
      


        Dim Rs As ADODB.Recordset
        Dim C As Long
        Dim cad As String
        SQL = "Select * from tmpinformes WHERE campo1 <> 0 and codusu = " & Usuario


        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        C = 0
        While Not Rs.EOF
            SQL = SQL & (Rs!codigo1) & "        "
            If (C Mod 6) = 5 Then SQL = SQL & vbCrLf
            C = C + 1
            Rs.MoveNext
        Wend
        Rs.Close
        If C > 0 Then
            Set frmMens = New frmMensajes
            
            frmMens.Cadena = "select * from tmpinformes where campo1 <> 0 and codusu = " & Usuario
            frmMens.OpcionMensaje = 1
            frmMens.Show vbModal
            
'            SQL = "Se han encontrado " & C & " registros con datos incorrectos en la BD: " & vbCrLf & SQL
'            SQL = SQL & " ¿Desea continuar ?"
'            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbNo Then Exit Sub
        End If

        'Abrimos los registros =0 k son los OK'
        SQL = "¿ Desea importar los trabajadores correctos ?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub


        SQL = "Select * from tmpinformes WHERE campo1 = 0 and codusu = " & Usuario & " and importe1 > 0"
        
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        C = 0
        While Not Rs.EOF
            C = C + 1
                        
            Trabajador = DevuelveValor("select codtraba from straba where codasesoria = " & Rs!codigo1)
                        
            vSQL = "insert into rrecasesoria (codtraba, fechahora, concepto, importe) values  "
            vSQL = vSQL & "( " & Trabajador & ",'" & Format(CDate(Rs!fecha1), "yyyy-mm-dd") & "',"
            vSQL = vSQL & "'" & Text7.Text & "'," & TransformaComasPuntos(Rs!importe1) & ") "
            
            Conn.Execute vSQL
                    
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    MsgBox "FIN", vbInformation
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click(Index As Integer)
Dim Rc As Byte
Dim Mens As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
        

    If Text1.Text = "" Then
        MsgBox "Escriba el nombre del fichero excel", vbExclamation
        Exit Sub
    End If
        
        
    If Dir(Text1.Text) = "" Then
        MsgBox "Fichero no existe"
        Exit Sub
    End If
    NombreHoja = Text1.Text
    
    If Text3.Text <> "" Then
        If Dir(Text3.Text) <> "" Then
            MsgBox "Fichero ya existe", vbExclamation
            Exit Sub
        Else
            FileCopy Text1.Text, Text3.Text
            NombreHoja = Text3.Text
        End If
    End If
    
    'Abrimos excel
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        'Realizamos todos los datos
        'abrimos conexion
        If AbrirConexion(BaseDatos) Then
        
            
            'Vamos linea a linea, buscamos su trabajador
            Mens = "Error insertando en Excel"
            RecorremosLineasPicassent Mens
            
        End If
    
        'Cerramos el excel
        CerrarExcel

        MsgBox "Proceso finalizado", vbExclamation

    End If
    
    
End Sub


Private Sub Form_Load()
    
'    Combo1.ListIndex = Month(Now) - 1
'    Text3.Text = Year(Now)
    FrameEscribir.visible = False
    FrameImportar.visible = False
    Me.FrameConfig.visible = False
    Me.FrameCompletar.visible = False
    Limpiar
    Select Case EsImportaci
    Case 0
        Caption = "CONFIGURACION"
        FrameConfig.visible = True
'        vConfiguracion True
    Case 1
        Caption = "Cargar Diferencias desde fichero excel"
        FrameImportar.visible = True
    Case 2
        Caption = "Crear fichero Nóminas"
        
        FrameEscribir.visible = True
    Case 3
        Caption = "Crear fichero Horas"
        FrameEscribir.visible = True
        
    Case 4
        Caption = "Crear fichero Nóminas"
        
        FrameCompletar.visible = True
    
        
    End Select
    
    
 
End Sub

Private Sub Limpiar()
Dim T As Control
    For Each T In Me.Controls
        If TypeOf T Is TextBox Then
            T.Text = ""
        End If
    Next
        
End Sub
Private Function TransformaComasPuntos(Cadena) As String
Dim cad As String
Dim J As Integer
    
    J = InStr(1, Cadena, ",")
    If J > 0 Then
        cad = Mid(Cadena, 1, J - 1) & "." & Mid(Cadena, J + 1)
    Else
        cad = Cadena
    End If
    TransformaComasPuntos = cad
End Function

Private Sub frmC_Selec(vFecha As Date)
   Text4.Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub

Private Sub Image1_Click()
    AbrirDialogo 1
End Sub

Private Sub Image2_Click()
    AbrirDialogo 1
End Sub


Private Sub AbrirDialogo(Opcion As Integer)

    On Error GoTo EA
    
    With Me.CommonDialog1
        Select Case Opcion
        Case 0, 2
            .DialogTitle = "Archivo origen de datos"
        Case 1
            .DialogTitle = "Archivo destino de datos"
        End Select
        '[Monica]07/11/2013: modificado añdo el tipo de datos XLSX
        .Filter = "EXCEL (*.xls)|*.xls|EXCEL (*.xlsx)|*.xlsx"
        .CancelError = True
        If Opcion <> 1 Then
            .ShowOpen
            If Opcion = 0 Then
                Text2.Text = .FileName
                
            Else
                If Opcion = 3 Then
                    Text1.Text = .FileName
                Else
                    Text5.Text = .FileName
                End If
            End If
        Else
            .ShowSave
            Text2.Text = .FileName
            Text3.Text = .FileName
        End If
        
    End With
EA:
End Sub

Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function RecorremosLineas(Mens As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim NFile As Integer
Dim Cooperativa As Integer

    On Error GoTo eRecorremosLineas

    RecorremosLineas = False


    SQL = "select * from tmpinformes where codusu = " & Usuario & " order by codigo1 "
    Sql1 = "select count(*) from tmpinformes where codusu = " & Usuario
    

    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Cooperativa = DevuelveValor("select cooperativa from rparam")
    If Cooperativa = 2 Then ExcelSheet.Cells(4, 3).Value = Format(RT!fecha1, "dd/mm/yyyy")  ' fecha
    
    I = 1
    While Not RT.EOF
        I = I + 1
    
        IncrementarProgresNew Pb1, 1
    
        If Cooperativa = 4 Then
            ExcelSheet.Cells(I, 1).Value = 3
            ExcelSheet.Cells(I, 2).Value = DBLet(RT!codigo1, "N")
            ExcelSheet.Cells(I, 3).Value = ""
            
            SQL = "select niftraba, nomtraba from straba where codtraba = " & RT!codigo1
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs.EOF Then
                ExcelSheet.Cells(I, 4).Value = Rs.Fields(0).Value ' nif del trabajador
                ExcelSheet.Cells(I, 5).Value = Rs.Fields(1).Value ' nombre del trabajador
            Else
                ExcelSheet.Cells(I, 4).Value = "" ' nif del trabajador
                ExcelSheet.Cells(I, 5).Value = "" ' nombre del trabajador
            End If
            
            Set Rs = Nothing
            
            ExcelSheet.Cells(I, 6).Value = Format(RT!fecha1, "dd/mm/yyyy") ' fecha
            
            
            ExcelSheet.Cells(I, 7).Value = DBLet(RT!importe2, "N") 'numero de dias trabajados
            ExcelSheet.Cells(I, 8).Value = 0
            ExcelSheet.Cells(I, 9).Value = 0
            ExcelSheet.Cells(I, 10).Value = DBLet(RT!importe3, "N") ' Importe bruto
            ExcelSheet.Cells(I, 11).Value = DBLet(RT!importe1, "N") ' importe anticipado
            ExcelSheet.Cells(I, 12).Value = DBLet(RT!nombre1, "N") ' cadena de dias trabajados
            
    '[Monica]04/11/2010: las columnas de 1 al 31 del mes no hay que ponerlas
    '        ExcelSheet.Cells(I, 13).Value = "" ' columna M
    '
    '        For J = 1 To 31
    '            ExcelSheet.Cells(I, 13 + J).Value = Mid(RT!nombre1, J, 1) ' el dia i ha trabajado
    '        Next J
    
        Else ' para Picassent
            ExcelSheet.Cells(I + 6, 1).Value = DBLet(RT!codigo1, "N")
            
            SQL = "select niftraba, nomtraba from straba where codtraba = " & RT!codigo1
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs.EOF Then
                ExcelSheet.Cells(I + 6, 2).Value = Rs.Fields(1).Value ' nombre del trabajador
            Else
                ExcelSheet.Cells(I + 6, 2).Value = "" ' nombre del trabajador
            End If
            
            Set Rs = Nothing
            
            ExcelSheet.Cells(I + 6, 3).Value = DBLet(RT!importe3, "N") ' Importe bruto
            ExcelSheet.Cells(I + 6, 4).Value = DBLet(RT!importe1, "N") ' importe anticipado
            ExcelSheet.Cells(I + 6, 5).Value = DBLet(RT!importe2, "N") ' numero de dias trabajados
        
        
        End If
    
    
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineas = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function


Private Function RecorremosLineasPicassent(Mens As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim NFile As Integer
Dim Cooperativa As Integer

    On Error GoTo eRecorremosLineas

    RecorremosLineasPicassent = False


    
    I = 1
    Cod = Trim(ExcelSheet.Cells(I + 8, 2).Value)
    While Cod <> ""
        I = I + 1
        Cod = Trim(ExcelSheet.Cells(I + 8, 2).Value)
    Wend
    
    Me.Pb2.visible = True
    Me.Pb2.Max = I
    Me.Pb2.Value = 0
    Me.Refresh
    
    Set RT = Nothing
    
    
    I = 9
    Cod = Trim(ExcelSheet.Cells(I, 2).Value)
    While Cod <> ""
        SQL = "select codtraba, niftraba, nomtraba from straba where codasesoria = " & Cod
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            Sql1 = "select * from tmpinformes where codusu = " & Usuario
            Sql1 = Sql1 & " and codigo1 = " & Rs!codtraba
            
            Set RT = New ADODB.Recordset
            RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RT.EOF Then
                ExcelSheet.Cells(I, 4).Value = DBLet(RT!importe3, "N")  ' Importe bruto
                ExcelSheet.Cells(I, 5).Value = DBLet(RT!importe1, "N")  ' importe anticipado
                ExcelSheet.Cells(I, 6).Value = DBLet(RT!importe2, "N") ' numero de dias trabajados
            End If
            Set RT = Nothing
        End If
        
        I = I + 1
    
        IncrementarProgresNew Pb2, 1
    
        Cod = Trim(ExcelSheet.Cells(I, 2).Value)
        
    Wend
        
    
    RecorremosLineasPicassent = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function




Private Function RecorremosLineasHoras(Mens As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim NFile As Integer

    On Error GoTo eRecorremosLineas

    RecorremosLineasHoras = False


    '[Monica] 19/04/2010: añadida la condicion del sql en el fichero condicionsql.txt
    If Dir(App.Path & "\condicionsql.txt", vbArchive) <> "" Then
    
        NFile = FreeFile
    
        Open App.Path & "\condicionsql.txt" For Input As #NFile
 
        If Not EOF(NFile) Then
            Line Input #NFile, Lin
    
            SQL = Lin
        End If
    End If
    '[Monica] 19/04/2010

    Sql1 = "select count(*) from (" & SQL & ") condicion "


    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    I = 0
    While Not RT.EOF
        I = I + 1
    
        IncrementarProgresNew Pb1, 1
        
        'caso de que exportemos las horas
        If InStr(1, SQL, "horas.") Then
            'SELECT horas.codtraba, straba.nomtraba, horas.fechahora, horas.codalmac, salmpr.nomalmac, horas.horasdia, horas.horasproduc, horas.compleme, horas.horasext, horas.fecharec,  horas.pasaridoc,  IF(pasaridoc=1,'*','') as pasari, horas.intconta,  IF(intconta=1,'*','') as intcon,  horas.nroparte  FROM horas, straba, salmpr  WHERE horas.codtraba = straba.codtraba and  horas.codalmac = salmpr.codalmac  ORDER BY  horas.fechahora desc, horas.codtraba
            For J = 0 To 13
                If I = 13 Then
                    ExcelSheet.Cells(I, J + 1).Value = DBLet(RT.Fields(J).Value, "F")
                Else
                    ExcelSheet.Cells(I, J + 1).Value = RT.Fields(J).Value
                End If
            Next J
        Else
        'caso de que exportemos las horas de destajo
            'SELECT horasdestajo.codtraba, straba.nomtraba, horasdestajo.fechahora, horasdestajo.codvarie, variedades.nomvarie, "
            'horasdestajo.codforfait, forfaits.nomconfe, horasdestajo.numcajon, horasdestajo.kilos, horasdestajo.importe, "
            'horasdestajo.horas "
            'FROM horasdestajo, straba, variedades, forfaits "
            ' WHERE horasdestajo.codtraba = straba.codtraba and "
            ' horasdestajo.codvarie = variedades.codvarie And ""
            ' horasdestajo.codforfait = forfaits.codforfait "
            For J = 0 To 10
                If I = 2 Then
                    ExcelSheet.Cells(I, J + 1).Value = DBLet(RT.Fields(J).Value, "F")
                Else
                    ExcelSheet.Cells(I, J + 1).Value = RT.Fields(J).Value
                End If
            Next J
        End If
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineasHoras = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function


Private Sub Image3_Click()
 AbrirDialogo 2
End Sub

Private Sub Image4_Click()
    AbrirDialogo 3
End Sub

Private Sub Image5_Click()
    MsgBox "Formato importe:   SOLO el punto decimal: 1.49", vbExclamation
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
      While imgFecha(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    imgFecha(Index).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text4.Text <> "" Then frmC.NovaData = Text4
    
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    Text4.SetFocus  '<===
    ' ********************************************
     

End Sub

Private Sub Text4_LostFocus()
    PonerFormatoFecha Text4
'    Text4.Text = Trim(Text4.Text)
'    If Text4.Text <> "" Then
'        If IsDate(Text4.Text) Then
'            Text4.Text = Format(Text4.Text, "dd/mm/yyyy")
'        Else
'            MsgBox "Fecha incorrecta", vbExclamation
'            Text4.Text = ""
'        End If
'    End If
End Sub



'-------------------------------------
Private Function RecorremosLineasLiquidacion()
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer
Dim Codigo As Long

    'Desde la fila donde empieza los trabajadores
    'Hasta k este vacio
    'Iremos insertando en tmpHoras
    ' Con trbajador, importe, 0 , 1 ,2
    '             Existe, No existe, IMPORTE negativo
    '
    
    SQL = "DELETE FROM tmpinformes where codusu = " & Usuario
    Conn.Execute SQL
    
    FIN = False
    I = 10 '2
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(I, 2).Value)) <> "" Then
            LineasEnBlanco = 0
            If IsNumeric((ExcelSheet.Cells(I, 2).Value)) Then
                If Val(ExcelSheet.Cells(I, 2).Value) > 0 Then
                        FechaHora = Text4.Text
                        Concepto = Text7.Text
                        'codigo asesoria
                        CodigoAsesoria = Val(ExcelSheet.Cells(I, 2).Value)
                        'Importe
                        Importe = CCur(ExcelSheet.Cells(I, 9).Value)
                        
'[Monica]14/12/2010: siempre es la misma columna
'                        If Importe = 0 Then
'                            Importe = CCur(ExcelSheet.Cells(I, mConfig.ColImporte + 1).Value)
'                        Else
'                            Importe = CCur(ExcelSheet.Cells(I, mConfig.ColImporte).Value)
'                        End If
                        
'                        Trabajador = DevuelveValor("select codtraba from straba where codasesoria = " & Codigo)
                        
                        'InsertartmpLiquida
                        InsertaTmpExcel
                    End If
            End If
        Else
            LineasEnBlanco = LineasEnBlanco + 1
            If LineasEnBlanco < 30 Then
               ' FIN = False
            Else
                FIN = True
               
            End If
        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
End Function


Private Sub InsertaTmpExcel()
Dim vSQL As String
Dim vSql2 As String
Dim RT As ADODB.Recordset
Dim RT1 As ADODB.Recordset
Dim RT2 As ADODB.Recordset
Dim Existe As Boolean
Dim ExisteCalidad As Boolean
Dim ExisteEnTemporal As Boolean
Dim TotalKilos As Long
Dim Cuadra As Boolean
Dim JJ As Integer
Dim NRegs As Integer

    On Error GoTo EInsertaTmpExcel
    
    vSQL = "Select * from straba where codasesoria = " & CodigoAsesoria
    Set RT = New ADODB.Recordset
    RT.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RT.EOF Then
        Existe = False
    Else
        Existe = True
    End If
    
    ' si existe el trabajador
    If Existe Then
    
        ExisteEnTemporal = False
        vSQL = "select * from tmpinformes where codigo1 = " & CodigoAsesoria & " and codusu = " & Usuario
    
        Set RT2 = New ADODB.Recordset
        RT2.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        If Not RT2.EOF Then
            ExisteEnTemporal = True
        End If
    
        SQL = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, nombre1, campo1) values ("
        SQL = SQL & Usuario & ","
        SQL = SQL & CodigoAsesoria & ","
        SQL = SQL & "'" & Format(CDate(FechaHora), "yyyy-mm-dd") & "',"
        SQL = SQL & TransformaComasPuntos(Importe) & ","
        SQL = SQL & "'" & Concepto & "',"
    
        If ExisteEnTemporal Then
            SQL = SQL & "2)"
        Else
            NRegs = DevuelveValor("select count(*) from rrecasesoria where codtraba = " & RT!codtraba & " and fechahora = " & "'" & Format(CDate(FechaHora), "yyyy-mm-dd") & "'")
            If NRegs = 0 Then
                SQL = SQL & "0)"
            Else
                SQL = SQL & "3)"
            End If
        End If
        
    Else
        SQL = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, nombre1, campo1) values ("
        SQL = SQL & Usuario & ","
        SQL = SQL & CodigoAsesoria & ","
        SQL = SQL & "'" & Format(CDate(FechaHora), "yyyy-mm-dd") & "',"
        SQL = SQL & TransformaComasPuntos(Importe) & ","
        SQL = SQL & "'" & Concepto & "',"
        SQL = SQL & "1)" ' no existe el trabajador
        
    End If
    
    
    If SQL <> "" Then Conn.Execute SQL
        
    RT.Close
    
    Exit Sub
EInsertaTmpExcel:
    MsgBox Err.Description
End Sub



Private Sub InsertaReciboAsesoria()
Dim vSQL As String
Dim vSql2 As String
Dim RT As ADODB.Recordset
Dim RT1 As ADODB.Recordset
Dim RT2 As ADODB.Recordset
Dim Existe As Boolean
Dim ExisteCalidad As Boolean
Dim ExisteEnTemporal As Boolean
Dim TotalKilos As Long
Dim Cuadra As Boolean
Dim JJ As Integer

    On Error GoTo EInsertaTmpExcel
    
    vSQL = "Select * from rrecasesoria "
    vSQL = vSQL & " WHERE codtraba = " & Trabajador
    vSQL = vSQL & " and fechahora = '" & Format(CDate(FechaHora), "yyyy-mm-dd") & "'"

    Set RT = New ADODB.Recordset
    RT.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RT.EOF Then
        Existe = False
    Else
        Existe = True
    End If
    
    ' si no existe el registro lo insertamos
    If Existe Then
        vSql2 = "insert into rrecasesoria (codtraba, fechahora, concepto, importe) values  "
        vSql2 = vSql2 & "( " & Trabajador & ",'" & Format(CDate(FechaHora), "yyyy-mm-dd") & ","
        vSql2 = vSql2 & "'" & Concepto & "'," & TransformaComasPuntos(Importe) & ") "
        
        Conn.Execute vSql2
    End If
    
    RT.Close
    
    Exit Sub
EInsertaTmpExcel:
    MsgBox Err.Description
End Sub



Private Sub vConfiguracion(Leer As Boolean)

'    With varConfig
'        If Leer Then
'            Text8(0).Text = .IniLinNomina
'            Text8(1).Text = .FinLinNominas
'            Text8(2).Text = .ColTrabajadorNom
'            Text8(3).Text = .hc
'            Text8(4).Text = .HPLUS
'            Text8(5).Text = .DIAST
'            Text8(6).Text = .Anticipos
'            Text8(7).Text = .ColTrabajadoresLIQ
'            Text8(8).Text = .ColumnaLiquidacion
'            Text8(9).Text = .FilaLIQ
'            Text8(10).Text = .HN
'        Else
'            .IniLinNomina = Val(Text8(0).Text)
'            .FinLinNominas = Val(Text8(1).Text)
'            .ColTrabajadorNom = Val(Text8(2).Text)
'            .hc = Val(Text8(3).Text)
'            .HPLUS = Val(Text8(4).Text)
'            .DIAST = Val(Text8(5).Text)
'            .Anticipos = Val(Text8(6).Text)
'            .ColTrabajadoresLIQ = Val(Text8(7).Text)
'            .ColumnaLiquidacion = Val(Text8(8).Text)
'            .FilaLIQ = Val(Text8(9).Text)
'            .HN = Val(Text8(10).Text)
'        End If
'    End With
End Sub

Private Sub Text8_GotFocus(Index As Integer)
    With Text8(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text8_LostFocus(Index As Integer)
    With Text8(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        Select Case Index
            Case 0 ' numero de calidades
                If Not IsNumeric(.Text) Then
                    MsgBox "Campo debe ser numérico", vbExclamation
                    .Text = ""
                    .SetFocus
                    Exit Sub
                End If
                .Text = Val(.Text)
            
            Case 2, 3 ' usuario y password deben de estar encriptados
            
            
        End Select
            
            
    End With
End Sub

Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function


Public Function PonerFormatoFecha(ByRef T As TextBox) As Boolean
Dim cad As String

    cad = T.Text
    If cad <> "" Then
        If Not EsFechaOK(cad) Then
            MsgBox "Fecha incorrecta. (dd/MM/yyyy)", vbExclamation
            cad = "mal"
        End If
        If cad <> "" And cad <> "mal" Then
            T.Text = cad
            PonerFormatoFecha = True
        Else
'                T.Text = ""
            PonerFoco T
        End If
    End If
End Function


Public Function EsFechaOK(T As String) As Boolean
Dim cad As String
Dim mes As String, dia As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        
      '==== Anade: Laura 04/02/2005 =============
        If Len(cad) < 6 Then
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el dia es correcto, valores entre 1-31
        dia = Mid(cad, 1, 2)
        If IsNumeric(dia) Then
            If dia < 1 Or dia > 31 Then
                EsFechaOK = False
                Exit Function
            End If
        Else
            EsFechaOK = False
            Exit Function
        End If
        
        'Comprobar que el mes es correcto, valores entre 1-12
        mes = Mid(cad, 3, 2)
        If IsNumeric(mes) Then
            If mes < 1 Or mes > 12 Then
                EsFechaOK = False
                Exit Function
            End If
        Else
            EsFechaOK = False
            Exit Function
        End If
      '============================================
        
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    Else
        dia = Mid(cad, 1, 2)
        mes = Mid(cad, 4, 2)
    End If
    
    If IsDate(cad) Then
        EsFechaOK = True
        T = Format(cad, "dd/MM/yyyy")
      '==== Añade: Laura 08/02/2005
        If Month(T) <> Val(mes) Then EsFechaOK = False
        If Day(T) <> Val(dia) Then EsFechaOK = False
      '====
    Else
        EsFechaOK = False
    End If
End Function

Private Sub text4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub text7_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me 'ESC
    End If
End Sub

Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

