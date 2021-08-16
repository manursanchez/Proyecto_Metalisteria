VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form17 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualización de facturas antiguas"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "Visualizacion facturas antiguas.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   6000
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle de Factura"
      Height          =   3495
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   7455
      Begin VB.Frame Frame4 
         Caption         =   "Totales"
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   7215
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            TabIndex        =   16
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   15
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "SUMA TOTAL"
            Height          =   255
            Left            =   4920
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "IVA"
            Height          =   255
            Left            =   2520
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "TOTAL"
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "Visualizacion facturas antiguas.frx":0442
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   "          |                                                   |                  |                  |             "
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de identificación de la factura"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         DataField       =   "fecha"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COD"
         DataField       =   "NumeroFactura"
         DataSource      =   "Data2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente de la factura"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   7455
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dirección"
         DataField       =   "direccion"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         DataField       =   "nombre"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Totales Antiguos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TotalesAntiguos"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data2 
      Caption         =   "AuxFacturas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data1 
      Caption         =   "Auxcli"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\Mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        DataEnvironment1.Connection1.Open
        DataEnvironment1.auxcli
        DataReport1.Show vbModal
        DataEnvironment1.Connection1.Close
End Sub

Private Sub Command2_Click()
        
        Unload Me
End Sub

Private Sub Form_Load()
    StatusBar1.SimpleText = "Trabajando con la empresa: " & NombreEmpresa
    Text1.Text = t
    Text2.Text = STI
    Text3.Text = tiva
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
        Loop
        Data1.Refresh
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.Delete
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
End Sub
