VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORMULARIO DE ACCESO"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8970
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Año Actual de facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   12
      Top             =   3120
      Width           =   2535
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         DataField       =   "AnoFacturacionActual"
         DataSource      =   "Data1"
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Año de facturacion"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\MTL.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AnoActual"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4935
      Width           =   8970
      _ExtentX        =   15822
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
   Begin VB.Frame Frame3 
      Caption         =   "Gestión de facturas antiguas"
      Height          =   2775
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command8 
         Caption         =   "RECUPERAR FACTURA ANTIGUA"
         Height          =   975
         Left            =   240
         Picture         =   "Principal.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "NUEVA FACTURACIÓN"
         Height          =   975
         Left            =   240
         Picture         =   "Principal.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CAMBIAR DE EMPRESA"
      Height          =   975
      Left            =   3240
      Picture         =   "Principal.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Presupuestos"
      Height          =   2775
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command5 
         Caption         =   "RECUPERAR"
         Height          =   975
         Left            =   240
         Picture         =   "Principal.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "NUEVO"
         Height          =   975
         Left            =   240
         Picture         =   "Principal.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Facturas"
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command3 
         Caption         =   "VENTAS"
         Height          =   975
         Left            =   240
         Picture         =   "Principal.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "RECUPERAR"
         Height          =   975
         Left            =   240
         Picture         =   "Principal.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "NUEVA"
         Height          =   975
         Left            =   240
         Picture         =   "Principal.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menu Principal"
      Begin VB.Menu acerca 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu reparacion 
         Caption         =   "Reparación manual del programa"
      End
      Begin VB.Menu guion 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu facturas 
      Caption         =   "Facturas"
      Begin VB.Menu nfactura 
         Caption         =   "Nueva factura"
      End
      Begin VB.Menu rfactura 
         Caption         =   "Recuperar factura"
      End
      Begin VB.Menu lfacturas 
         Caption         =   "Informe de ventas"
      End
   End
   Begin VB.Menu presupuestos 
      Caption         =   "Presupuestos"
      Begin VB.Menu npresupuesto 
         Caption         =   "Nuevo Presupuesto"
      End
      Begin VB.Menu rpresupuesto 
         Caption         =   "Recuperar Presupuesto"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acerca_Click()
    Load Form9
    Form9.Show vbModal
End Sub

Private Sub Command1_Click()
    Load Form2
    Form2.Show vbModal
End Sub

Private Sub Command2_Click()
    Load Form5
    Form5.Show vbModal
End Sub

Private Sub Command3_Click()
    Load Form7
    Form7.Show vbModal
End Sub

Private Sub Command4_Click()
    Load Form6
    Form6.Show vbModal
End Sub

Private Sub Command5_Click()
    Load Form13
    Form13.Show vbModal
End Sub

Private Sub Command6_Click()
    NombreEmpresa = ""
    Empresa = 0
    Unload Me
End Sub

Private Sub Command7_Click()
    Dim opc As Integer
    opc = MsgBox("¡¡ADVERTENCIA!!, Si continúa las facturas del ultimo año se archivarán y comenzará un nuevo año de facturación,¿Desea continuar?", vbOKCancel)
    If opc = 1 Then
        Load Form15
        Form15.Show vbModal
    Else
        MsgBox "Proceso detenido por el usuario", vbInformation, "Información del sistema"
    End If
End Sub

Private Sub Command8_Click()
    Load Form16
    Form16.Show vbModal
End Sub

Private Sub Form_Load()
    StatusBar1.SimpleText = "Trabajando con la empresa: " & NombreEmpresa
    recuperada = False
End Sub


Private Sub lfacturas_Click()
    Load Form7
    Form7.Show vbModal
End Sub

Private Sub nfactura_Click()
    Load Form2
    Form2.Show vbModal
End Sub

Private Sub npresupuesto_Click()
    Load Form6
    Form6.Show vbModal
End Sub

Private Sub reparacion_Click()
    Load Form14
    Form14.Show vbModal
End Sub

Private Sub rfactura_Click()
    Load Form5
    Form5.Show vbModal
End Sub


Private Sub rpresupuesto_Click()
    Load Form13
    Form13.Show vbModal
End Sub

Private Sub salir_Click()
    Unload Me
End Sub


