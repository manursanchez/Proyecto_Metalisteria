VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reparacion automatizada del programa"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "Reparacion automatizada del programa.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Progreso"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6375
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
         Max             =   60
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobando:..."
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
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comenzar la operación"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Data Data8 
      Caption         =   "Totales de presupuesto"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totalesdepresupuesto"
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Data Data7 
      Caption         =   "Clientes de presupuesto"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clientesdepresupuesto"
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Data Data6 
      Caption         =   "Presupuesto"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Presupuesto"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Data Data5 
      Caption         =   "Totales"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Data Data4 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Data Data3 
      Caption         =   "facturas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Factura"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Data Data2 
      Caption         =   "Auxfacturas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   600
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Auxcli"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Dim bandera As Boolean
    'Dim codigo As Double
    Dim contador As Integer
    Dim valor As Integer
    'bandera = False
    ' Borrado de tablas auxiliares
    If Not Data2.Recordset.BOF Then
        Do While Not Data2.Recordset.EOF
            contador = contador + 1
            Data2.Recordset.MoveNext
        Loop
    End If
    ProgressBar1.Min = 0
    ProgressBar1.Max = contador + 2
    valor = 0
    If Not Data1.Recordset.BOF Then
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
            valor = valor + 1
            ProgressBar1.Value = valor
        Loop
    End If
    If Not Data2.Recordset.BOF Then
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.Delete
            Data2.Recordset.MoveNext
            valor = valor + 1
            ProgressBar1.Value = valor
        Loop
    End If
    MsgBox "Comprobación finalizada", vbInformation, "Informacion de sistema"
    
    'Comprobacion de FACTURAS
    'Data3.Refresh
    'Data4.Refresh
    'If Data3.Recordset.BOF Then
    '    MsgBox "No hay registros"
    'Else
    '    Data3.Recordset.MoveFirst
    '    Do While Not Data3.Recordset.EOF
    '        codigo = Data3.Recordset.Fields("codfac")
    '        If Data4.Recordset.BOF Then
    '            MsgBox "No hay registros en Data4"
    '        Else
    '            Do While Not Data4.Recordset.EOF
    '                If Data4.Recordset.Fields("codfac") = codigo Then
    '                    MsgBox "regisxtro encontrado"
    '                    Data4.Recordset.MoveNext
    '                Else
    '                    MsgBox "registro no encontrado"
    '                    Data4.Recordset.MoveNext
    '                End If
    '            Loop
    '
    '        End If
    '      Data3.Recordset.MoveNext
    '    Loop
    'End If
    '
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
