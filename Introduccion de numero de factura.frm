VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NUEVA FACTURA"
   ClientHeight    =   2085
   ClientLeft      =   7590
   ClientTop       =   6960
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Introduccion de numero de factura.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4215
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
   Begin VB.Frame Frame2 
      Caption         =   "Introduce código de factura"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   645
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      RecordSource    =   "AnoActual"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Factura"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim consulta As String
    comprobar
    If Text1.Text = "" Or Not IsNumeric(Text1.Text) Then
        MsgBox "Debes introducir un numero en el cuadro negro", vbInformation, "Informacion"
    Else
        consulta = "select * from factura where codfac=" & Text1.Text & " AND CodigoEmpresa=" & Empresa
        Data1.RecordSource = consulta
        Data1.Refresh
        If Data1.Recordset.RecordCount < 1 Then
            numfactura = Text1.Text
            Anno = Data4.Recordset.Fields("AnoFacturacionActual")
            Unload Form2
            Load Form3
            Form3.Show vbModal
        Else
            MsgBox "El número de factura ya existe,eso significa que ya existe una factura con ese numero y no puede haber dos facturas con numeros iguales", vbExclamation, "Error de datos"
            Text1.SetFocus
        End If
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub comprobar()
     If Not Data2.Recordset.BOF Then
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.Delete
            Data2.Recordset.MoveNext
        Loop
        MsgBox "Error de datos. Pulse aceptar para continuar.", vbInformation, "Informacion del sistema"
        'Si da este error, es que ha ocurrido algo en tablas auxilares
     End If
     If Not Data3.Recordset.BOF Then
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
            Data3.Recordset.Delete
            Data3.Recordset.MoveNext
        Loop
        MsgBox "He encontrado errores en la base de datos.Estos errores son debidos al mal apagado del ordenador. He procedido a su reparación. No obstante, revise las últimas FACTURAS, por si encuentra algún error. Si encontrara algun error mas, póngase en contacto con el programador.", vbInformation, "Informacion del sistema"
    End If
End Sub

Private Sub Form_Load()
    Text3.Text = NombreEmpresa
    Anno = 0
End Sub
