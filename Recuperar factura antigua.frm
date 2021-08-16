VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperar factura antigua"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "Recuperar factura antigua.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Introducir datos de factura a recuperar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5175
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   3375
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
         Height          =   375
         Left            =   3840
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
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
         Height          =   405
         Left            =   3840
         MaxLength       =   4
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
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
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Introduce el codigo de la factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Introduce el año de la factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.Data Data5 
      Caption         =   "ClientesAntiguos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ClientesAntiguos"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data4 
      Caption         =   "FacturasAntiguas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FacturasAntiguas"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data3 
      Caption         =   "TotalesAntiguos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TotalesAntiguos"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data2 
      Caption         =   "AuxFacturas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Auxclientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean
Private Sub Command1_Click()
    If Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Then
        MsgBox "Introduzca correctamente los datos", vbInformation, "Error de datos"
        Text1.SetFocus
    Else
            
            Data4.RecordSource = "select * from FacturasAntiguas where CodFacAnt=" & Text1.Text & " AND CodigoEmpresa=" & Empresa & "and Anno=" + Text2.Text
            '"Select * from FacturasAntiguas where CodFacAnt=" + Text1.Text & "and Anno=" + Text2.Text
            Data4.Refresh
            If Data4.Recordset.RecordCount = 1 Then
                
                'Metemos en la tabla auxiliar de facturas la factura consultada
                Data4.Recordset.MoveFirst
                Do While Not Data4.Recordset.EOF
                        Data2.Recordset.AddNew
                        Data2.Recordset.Fields("numerofactura") = Data4.Recordset.Fields("codfacAnt")
                        Data2.Recordset.Fields("concepto") = Data4.Recordset.Fields("concepto")
                        Data2.Recordset.Fields("TotalIVA") = Data4.Recordset.Fields("TotalIVA")
                        ivarecuperado = Data4.Recordset.Fields("iva")
                        Data2.Recordset.Fields("iva") = Data4.Recordset.Fields("iva")
                        Data2.Recordset.Fields("total") = Data4.Recordset.Fields("total")
                        Data2.Recordset.Fields("total+iva") = Data4.Recordset.Fields("total+iva")
                        Data2.Recordset.Update
                        Data2.Refresh
                        Data4.Recordset.MoveNext
                        bandera = True
                Loop
                Data5.RecordSource = "select * from ClientesAntiguos where codfac=" & Text1.Text & " AND CodigoEmpresa=" & Empresa & "and ano=" + Text2.Text
                '"select * from ClientesAntiguos where codfac=" + Text1.Text & "and ano=" + Text2.Text
                Data5.Refresh
                
                'Metemos en la tabla auxiliar de clientes el cliente consultado
                
                Data1.Recordset.AddNew
                Data1.Recordset.Fields("codfac") = Data5.Recordset.Fields("codfac")
                Data1.Recordset.Fields("nombre") = Data5.Recordset.Fields("nombre")
                Data1.Recordset.Fields("direccion") = Data5.Recordset.Fields("direccion")
                Data1.Recordset.Fields("fecha") = Data5.Recordset.Fields("fecha")
                Data1.Recordset.Update
                Data1.Refresh
                        
                'Recuperamos los totales de esta factura
                
                Data3.RecordSource = "select * from totalesAntiguos where CodFacAnt=" & Text1.Text & " AND CodigoEmpresa=" & Empresa & "and Anno=" + Text2.Text
                '"Select * from totalesAntiguos where codfacAnt=" + Text1.Text & "and Anno=" + Text2.Text
                Data3.Refresh
                
                t = Data3.Recordset.Fields("total")
                tiva = Data3.Recordset.Fields("total+iva")
                STI = Data3.Recordset.Fields("sumatotaliva")
                
            End If
            If bandera = False Then
                    MsgBox "Esa factura no existe, Asegurese que introduce un numero de factura correcto", vbInformation, "Error de datos"
                    Text1.SetFocus
            Else
                Load Form17
                Form17.Show vbModal
                Unload Me
            End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text3.Text = NombreEmpresa
    bandera = False
End Sub

