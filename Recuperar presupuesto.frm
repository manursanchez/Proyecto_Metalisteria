VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECUPERAR PRESUPUESTO"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "Recuperar presupuesto.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Introduce número de presupuesto"
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
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
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
         Height          =   615
         Left            =   120
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Presupuesto"
      Top             =   2160
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   2640
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clientesdepresupuesto"
      Top             =   3120
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   3600
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totalesdepresupuesto"
      Top             =   4080
      Width           =   2775
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim bandera As Boolean
    bandera = False
    recuperada = True
    Data1.RecordSource = "select * from presupuesto where codfac=" + Text1.Text
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
        Data3.RecordSource = "select * from clientesdepresupuesto where codfac=" + Text1.Text
        Data3.Refresh
        
        'Metemos en la tabla auxiliar de clientes el cliente consultado
        
        Data4.Recordset.AddNew
        Data4.Recordset.Fields("codfac") = Data3.Recordset.Fields("codfac")
        Data4.Recordset.Fields("nombre") = Data3.Recordset.Fields("nombre")
        Data4.Recordset.Fields("direccion") = Data3.Recordset.Fields("direccion")
        Data4.Recordset.Fields("fecha") = Data3.Recordset.Fields("fecha")
        Data4.Recordset.Update
        Data4.Refresh
        
        'Recuperamos los totales de este presupuesto
        Data5.RecordSource = "Select * from totalesdepresupuesto where codfac=" + Text1.Text
        Data5.Refresh
        t = Data5.Recordset.Fields("total")
        tiva = Data5.Recordset.Fields("total+iva")
        STI = Data5.Recordset.Fields("sumatotaliva")
        
        'Borramos el presupuesto de la tabla principal
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
           Data1.Recordset.Delete
           Data1.Recordset.MoveNext
        Loop
        
        'Borramos el cliente de presupuesto de la tabla principal
        Data3.Recordset.MoveFirst
        Data3.Recordset.Delete
        
        'Borramos los totales de presupuestos
        Data5.Recordset.MoveFirst
        Data5.Recordset.Delete
    
    End If
    If bandera = False Then
            MsgBox "Este presupuesto no existe, Asegurese que introduce un numero de presupuesto correcto", vbInformation, "Error de datos"
    Else
        Load Form11
        Form11.Show vbModal
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

