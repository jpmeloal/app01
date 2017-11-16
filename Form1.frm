VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "CONTROL ASISTENCIA"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   5040
   End
   Begin VB.CommandButton btnSalida 
      BackColor       =   &H008080FF&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton btnIngreso 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Nombre PC:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Colaborador:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label txtpc 
      BackColor       =   &H00800000&
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label txtusuario 
      BackColor       =   &H00800000&
      Caption         =   "usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "hh:mm:ss"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "dd/mm/aaaa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Line linSegundos 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   2160
      X2              =   2160
      Y1              =   2280
      Y2              =   3480
   End
   Begin VB.Line linMinutos 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2160
      X2              =   3120
      Y1              =   2280
      Y2              =   1920
   End
   Begin VB.Line linHoras 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   1440
      X2              =   2160
      Y1              =   2160
      Y2              =   2280
   End
   Begin VB.Shape shpCirculo 
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape shpCirculo2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   960
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnIngreso_Click()
    
    'ENTRADA_B18938_24072017_061255.csv
    Dim obj As FileSystemObject
    Dim tx As Scripting.TextStream
    Dim NombreArchivo, RutaArchivo As String
    Dim PcName As String
    Dim UserName As String
    Dim sArchivo As String
    Dim msj As String
    Dim ins As Date
    Dim fec As String
    Dim val As String
    Dim hora, min, seg As Integer
    
    PcName = Environ("COMPUTERNAME")
    UserName = UCase(Environ("USERNAME"))
    RutaArchivo = "\\grupoib.local\dfs3\Asistencia_GTP"
    ins = Now
    hora = Hour(ins)
    min = Minute(ins)
    seg = Second(ins)
    fec = Right("0" & Day(Date), 2) & Right("0" & Month(Date), 2) & Year(Date)
    val = "no"
    
    On Error GoTo ErrorHandler
    sArchivo = Dir(RutaArchivo & "\")
    Do While sArchivo <> ""
        If InStr(1, sArchivo, "ENTRADA") > 0 And InStr(1, sArchivo, UserName) > 0 And InStr(1, sArchivo, fec) > 0 Then
            val = "ok"
        End If
    sArchivo = Dir()
    Loop
    
    If val = "no" Then
        If min >= 10 Then
            min = min - 10
        Else
            min = min + 60 - 10
            hora = hora - 1
        End If
    
        NombreArchivo = "ENTRADA_" & UserName & "_" & fec & "_" & Right("0" & hora, 2) & Right("0" & min, 2) & Right("0" & seg, 2) & ".csv"
        msj = Hour(ins) & Right("0" & Minute(ins), 2) & Right("0" & Second(ins), 2)
        RutaArchivo = RutaArchivo & "\" & NombreArchivo
        'Generamos el archivo de entrada
        Set obj = New FileSystemObject
        Set tx = obj.CreateTextFile(RutaArchivo)
        tx.Close

        MsgBox "HORA ENTRADA REGISTRADA: " & TimeValue(Now)
    Else
        MsgBox "LA HORA DE INGRESO YA HA SIDO REGISTRADA", vbExclamation, "Alerta"
    End If
    Exit Sub
ErrorHandler:
    MsgBox "NO CUENTA CON PERMISO AL SERVIDOR DE ENVIO", vbExclamation, "Alerta"
    
End Sub

Private Sub btnSalida_Click()
    
    'SALIDA_B18938_24072017_061255.csv
    Dim obj As FileSystemObject
    Dim tx As Scripting.TextStream
    Dim NombreArchivo, RutaArchivo As String
    Dim PcName As String
    Dim UserName As String
    Dim sArchivo As String
    Dim msj As String
    Dim ins As Date
    Dim fec As String
    
    PcName = Environ("COMPUTERNAME")
    UserName = UCase(Environ("USERNAME"))
    RutaArchivo = "\\grupoib.local\dfs3\Asistencia_GTP"
    ins = Now
    fec = Right("0" & Day(Date), 2) & Right("0" & Month(Date), 2) & Year(Date)
    
    On Error GoTo ErrorHandler
    sArchivo = Dir(RutaArchivo & "\")
    Do While sArchivo <> ""
        If InStr(1, sArchivo, "SALIDA") > 0 And InStr(1, sArchivo, UserName) > 0 And InStr(1, sArchivo, fec) > 0 Then
            Kill (RutaArchivo & "\" & sArchivo)
        End If
    sArchivo = Dir()
    Loop
    
    If True Then
        NombreArchivo = "SALIDA_" & UserName & "_" & fec & "_" & Right("0" & Hour(ins), 2) & Right("0" & Minute(ins), 2) & Right("0" & Second(ins), 2) & ".csv"
        msj = Hour(ins) & Right("0" & Minute(ins), 2) & Right("0" & Second(ins), 2)
        RutaArchivo = RutaArchivo & "\" & NombreArchivo
        'Generamos el archivo de entrada
        Set obj = New FileSystemObject
        Set tx = obj.CreateTextFile(RutaArchivo)
        tx.Close
        MsgBox "HORA SALIDA REGISTRADA: " & TimeValue(Now)
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "NO CUENTA CON PERMISO AL SERVIDOR DE ENVIO", vbExclamation, "Alerta"
    
End Sub

Private Sub Form_Load()
    
    Dim PcName As String
    Dim UserName As String
    PcName = Environ("COMPUTERNAME")
    UserName = Environ("USERNAME")
    
    Label3.Caption = Date   ' Fecha Actual
    txtpc.Caption = PcName
    txtusuario.Caption = UserName
    'Ubicar agujas en el centro del circulo
    linHoras.X1 = shpCirculo.Left + shpCirculo.Width \ 2
    linHoras.Y1 = shpCirculo.Top + shpCirculo.Height \ 2
    linMinutos.X1 = shpCirculo.Left + shpCirculo.Width \ 2
    linMinutos.Y1 = shpCirculo.Top + shpCirculo.Height \ 2
    linSegundos.X1 = shpCirculo.Left + shpCirculo.Width \ 2
    linSegundos.Y1 = shpCirculo.Top + shpCirculo.Height \ 2
End Sub

Private Sub Timer1_Timer()
    
    Label4.Caption = Time

    Const R = 3.1415927 / 180
    Dim CentroX As Integer
    Dim CentroY As Integer
        
    'Centro de las agujas
    CentroX = shpCirculo.Left + shpCirculo.Width \ 2
    CentroY = shpCirculo.Top + shpCirculo.Height \ 2
        
    'Mover agujas cada segundo
    linHoras.X2 = CentroX + Sin(Hour(Now) * 30 * R) * (shpCirculo2.Width - shpCirculo.Width) * 0.34
    linHoras.Y2 = CentroY - Cos(Hour(Now) * 30 * R) * (shpCirculo2.Height - shpCirculo.Height) * 0.34
    linMinutos.X2 = CentroX + Sin(Minute(Now) * 6 * R) * (shpCirculo2.Width - shpCirculo.Width) * 0.48
    linMinutos.Y2 = CentroY - Cos(Minute(Now) * 6 * R) * (shpCirculo2.Height - shpCirculo.Height) * 0.48
    linSegundos.X2 = CentroX + Sin(Second(Now) * 6 * R) * (shpCirculo2.Width - shpCirculo.Width) * 0.64
    linSegundos.Y2 = CentroY - Cos(Second(Now) * 6 * R) * (shpCirculo2.Height - shpCirculo.Height) * 0.64
End Sub

