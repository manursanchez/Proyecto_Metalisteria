VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperar factura"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   Icon            =   "Recuperar factura.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
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
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Introduce el número de factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   120
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Data Data6 
      Caption         =   "Año de facturacion"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\MTL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AnoActual"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Data Data5 
      Caption         =   "Totales"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Data Data4 
      Caption         =   "Auxcli"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Data Data3 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Factura"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Empresa:"
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
      Width           =   1095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim bandera As Boolean
    bandera = False
    recuperada = True
    Data1.RecordSource = "select * from factura where codfac=" & Text1.Text & " AND CodigoEmpresa=" & Empresa
    Data1.Refresh
    If Data1.Recordset.RecordCount = 1 Then
        'Metemos en la tabla auxiliar de facturas la factura consultada
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
                Data2.Recordset.AddNew
                Data2.Recordset.Fields("numerofactura") = Data1.Recordset.Fields("codfac")
                Data2.Recordset.Fields("concepto") = Data1.Recordset.Fields("concepto")
                Data2.Recordset.Fields("TotalIVA") = Data1.Recordset.Fields("TotalIVA")
                ivarecuperado = Data1.Recordset.Fields("iva")
                Data2.Recordset.Fields("iva") = Data1.Recordset.Fields("iva")
                Data2.Recordset.Fields("total") = Data1.Recordset.Fields("total")
                Data2.Recordset.Fields("total+iva") = Data1.Recordset.Fields("total+iva")
                Data2.Recordset.Update
                Data2.Refresh
                Data1.Recordset.MoveNext
                bandera = True
        Loop
        Data3.RecordSource = "select * from clientes where codfac=" & Text1.Text & " AND CodigoEmpresa=" & Empresa
        '"select * from clientes where codfac=" + Text1.Text
        Data3.Refresh
        
        'Metemos en la tabla auxiliar de clientes el cliente consultado
        
        Data4.Recordset.AddNew
        Data4.Recordset.Fields("codfac") = Data3.Recordset.Fields("codfac")
        Data4.Recordset.Fields("nombre") = Data3.Recordset.Fields("nombre")
        Data4.Recordset.Fields("direccion") = Data3.Recordset.Fields("direccion")
        Data4.Recordset.Fields("fecha") = Data3.Recordset.Fields("fecha")
        Data4.Recordset.Update
        Data4.Refresh
        
        'Recuperamos los totales de esta factura
        Data5.RecordSource = "select * from totales where codfac=" & Text1.Text & " AND CodigoEmpresa=" & Empresa
        '"Select * from totales where codfac=" + Text1.Text
        Data5.Refresh
        t = Data5.Recordset.Fields("total")
        tiva = Data5.Recordset.Fields("total+iva")
        STI = Data5.Recordset.Fields("sumatotaliva")
        
        'Borramos la factura de la tabla principal
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
           Data1.Recordset.Delete
           Data1.Recordset.MoveNext
        Loop
        
        'Borramos el cliente de la tabla principal
        Data3.Recordset.MoveFirst
        Data3.Recordset.Delete
        
        'Borramos los totales
        Data5.Recordset.MoveFirst
        Data5.Recordset.Delete
        
        'Metemos en la variable global Anno el año de facturacion
        Anno = Data6.Recordset.Fields("AnoFacturacionActual")
    End If
    If bandera = False Then
            MsgBox "Esa factura no existe, Asegurese que introduce un numero de factura correcto", vbInformation, "Error de datos"
            Text1.SetFocus
    Else
        Load Form4
        Form4.Show vbModal
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text3.Text = NombreEmpresa
End Sub
