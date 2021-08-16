VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Año en Factura"
   ClientHeight    =   3420
   ClientLeft      =   7425
   ClientTop       =   6150
   ClientWidth     =   5040
   ControlBox      =   0   'False
   Icon            =   "Año en factura.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5040
   Begin VB.Data Data7 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data6 
      Caption         =   "ClientesAntiguos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ClientesAntiguos"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data5 
      Caption         =   "Totales"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\MTL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   6000
      Width           =   3855
   End
   Begin VB.Data Data4 
      Caption         =   "TotalesAntiguos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\MTL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TotalesAntiguos"
      Top             =   5640
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Data Data3 
      Caption         =   "ANOACTUAL"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\MTL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AnoActual"
      Top             =   5280
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data2 
      Caption         =   "facturas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\MTL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Factura"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "FacturasAntiguas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\MTL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FacturasAntiguas"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Archivar  >>"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Ultimo año de facturacion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Introduce nuevo año de facturación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Se va a proceder a archivar las facturas del último año. Pulse Archivar para continuar..."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean
Private Sub Command1_Click()
    Dim opcion As Integer
    bandera = False
    GuardarFacturasAntiguas
    If bandera = True Then
        opcion = MsgBox("¿Deseas cambiar el año de facturacion?", vbOKCancel)
    End If
    If opcion = 1 Or bandera = False Then
        Form15.Height = 3690
        Text1.SetFocus
        Command1.Enabled = False
        Command2.Enabled = False
        If Not Data3.Recordset.BOF Then
            Label5.Caption = Data3.Recordset.Fields("AnoFacturacionActual")
            Text1.Text = Val(Label5.Caption + 1)
            Data3.Recordset.MoveFirst
            Data3.Recordset.Delete
            Data3.Refresh
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If Data3.Recordset.BOF Then
        If Not IsNumeric(Text1.Text) Or Text1.Text = " " Then
        'Or Text1.Text < Label1.Caption Esto lo he quitado de la anterior linea de codigo
        'para poder meter el año de facturacion que yo quiera
            MsgBox "Los valores introducidos no son correctos"
        Else
            Data3.Recordset.AddNew
            Data3.Recordset.Fields("AnoFacturacionActual") = Val(Text1.Text)
            Data3.UpdateRecord
            Data3.Refresh
            MsgBox "Debe de reiniciar el programa para que los cambios surtan efecto", vbInformation
            Command3.Enabled = False
            Command2.Enabled = True
            Command2.Caption = "Cerrar"
        End If
    End If
End Sub

Private Sub Form_Load()
    Form15.Height = 1920
End Sub
Private Sub GuardarFacturasAntiguas()
    If Data2.Recordset.BOF Then
        MsgBox "No hay facturas para archivar, no puedo continuar", vbCritical, "Error de usuario"
        bandera = True
    Else
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data1.Recordset.AddNew
            Data1.Recordset.Fields("CodigoEmpresa") = Data2.Recordset.Fields("CodigoEmpresa")
            Data1.Recordset.Fields("NombreEmpresa") = Data2.Recordset.Fields("NombreEmpresa")
            Data1.Recordset.Fields("codfacAnt") = Data2.Recordset.Fields("CodFac")
            Data1.Recordset.Fields("fecha") = Data2.Recordset.Fields("fecha")
            Data1.Recordset.Fields("concepto") = Data2.Recordset.Fields("concepto")
            Data1.Recordset.Fields("TotalIVA") = Data2.Recordset.Fields("TotalIVA")
            Data1.Recordset.Fields("iva") = Data2.Recordset.Fields("iva")
            Data1.Recordset.Fields("total") = Data2.Recordset.Fields("total")
            Data1.Recordset.Fields("total+iva") = Data2.Recordset.Fields("total+iva")
            Data1.Recordset.Fields("Anno") = Data2.Recordset.Fields("Ano")
            Data1.UpdateRecord
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
        BorrarFacturasUltimoAno
    End If
End Sub
Private Sub BorrarFacturasUltimoAno()
    If Not Data2.Recordset.BOF Then
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.Delete
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
    End If
    GuardarTotalesAntiguos
End Sub
Private Sub GuardarTotalesAntiguos()
    Data5.Recordset.MoveFirst
    Do While Not Data5.Recordset.EOF
        Data4.Recordset.AddNew
        Data4.Recordset.Fields("CodigoEmpresa") = Data5.Recordset.Fields("CodigoEmpresa")
        Data4.Recordset.Fields("NombreEmpresa") = Data5.Recordset.Fields("NombreEmpresa")
        Data4.Recordset.Fields("codfacAnt") = Data5.Recordset.Fields("CodFac")
        Data4.Recordset.Fields("total") = Data5.Recordset.Fields("total")
        Data4.Recordset.Fields("total+iva") = Data5.Recordset.Fields("total+iva")
        Data4.Recordset.Fields("fecha") = Data5.Recordset.Fields("fecha")
        Data4.Recordset.Fields("sumatotaliva") = Data5.Recordset.Fields("sumatotaliva")
        Data4.Recordset.Fields("Anno") = Data5.Recordset.Fields("Ano")
        Data4.UpdateRecord
        Data5.Recordset.MoveNext
    Loop
    Data5.Refresh
    BorrarTotalesUltimoAno
    MsgBox "Factura archivada sin problemas", vbInformation
End Sub
Private Sub BorrarTotalesUltimoAno()
     If Not Data5.Recordset.BOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
            Data5.Recordset.Delete
            Data5.Recordset.MoveNext
        Loop
        Data5.Refresh
    End If
    GuardarClientesAntiguos
End Sub
Private Sub GuardarClientesAntiguos()
    Data7.Recordset.MoveFirst
    Do While Not Data7.Recordset.EOF
        Data6.Recordset.AddNew
        Data6.Recordset.Fields("CodigoEmpresa") = Data7.Recordset.Fields("CodigoEmpresa")
        Data6.Recordset.Fields("NombreEmpresa") = Data7.Recordset.Fields("NombreEmpresa")
        Data6.Recordset.Fields("codfac") = Data7.Recordset.Fields("CodFac")
        Data6.Recordset.Fields("Nombre") = Data7.Recordset.Fields("Nombre")
        Data6.Recordset.Fields("Direccion") = Data7.Recordset.Fields("Direccion")
        Data6.Recordset.Fields("fecha") = Data7.Recordset.Fields("fecha")
        Data6.Recordset.Fields("Ano") = Data7.Recordset.Fields("Ano")
        Data6.UpdateRecord
        Data7.Recordset.MoveNext
    Loop
    Data7.Refresh
    BorrarClientesAntiguos
End Sub
Private Sub BorrarClientesAntiguos()
    If Not Data7.Recordset.BOF Then
        Data7.Recordset.MoveFirst
        Do While Not Data7.Recordset.EOF
            Data7.Recordset.Delete
            Data7.Recordset.MoveNext
        Loop
        Data7.Refresh
    End If
    MsgBox "Clientes del año anterior archivados correctamente", vbInformation, "Información del sistema"
End Sub
