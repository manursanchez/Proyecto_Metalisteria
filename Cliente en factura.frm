VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introduce los datos del cliente"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "Cliente en factura.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      MaxLength       =   150
      TabIndex        =   2
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      MaxLength       =   150
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "No ha introducido el NOMBRE a quien va dirigida la factura.Debe introducirlo para continuar", vbInformation, "Informacion del sistema"
    Else
        If Text2.Text = "" Then
            MsgBox "No ha introducido la DIRECCION a donde va dirigida la factura.Debe introducirla para continuar", vbInformation, "Informacion del sistema"
        Else
            Data1.Recordset.AddNew
            Data1.Recordset.Fields("codfac") = numfactura
            Data1.Recordset.Fields("nombre") = Text1.Text
            Data1.Recordset.Fields("direccion") = Text2.Text
            Data1.Recordset.Fields("fecha") = Date
            Data1.Recordset.Update
            Data1.Refresh
            nombre = Text1.Text
            direccion = Text2.Text
            Unload Form3
            Load Form4
            Form4.Show vbModal
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

