VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminacion de elementos en la factura"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "Eliminacion de elementos.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "Total+iva"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Text            =   "total+iva"
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar elemento"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "Total"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataField       =   "TotalIVA"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Pulse las flechitas de los laterales para recorrer los elementos de la factura"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   1800
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      DataField       =   "Concepto"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "IVA:"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Total + IVA:"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Total Factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Total"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Precio"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Data1.Recordset.BOF Then
        MsgBox "No hay registros para borrar", vbInformation, "Información de sistema"
    Else
        Form8.Text4.Text = Form8.Text4.Text - Form8.Text3.Text
        Form8.Text5.Text = Form8.Text5.Text - Form8.Text6.Text
        Form8.Text7.Text = Form8.Text7.Text - Form8.Text2.Text
        Form8.Text4.Text = Round(Form8.Text4.Text, 2)
        Form8.Text5.Text = Round(Form8.Text5.Text, 2)
        Form8.Text7.Text = Round(Form8.Text7.Text, 2)
        Data1.Recordset.Delete
        Data1.Refresh
    End If
End Sub

Private Sub Command2_Click()
    Form4.Text9.Text = Form8.Text4.Text
    Form4.Text7.Text = Form8.Text5.Text
    Form4.Text8.Text = Form8.Text7.Text
    Form4.MSFlexGrid1.Refresh
    Form4.Data1.Refresh
    If Data1.Recordset.BOF Then
        Form4.Text6.Enabled = True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Form8.Text4.Text = Form4.Text9.Text
    Form8.Text5.Text = Form4.Text7.Text
    Form8.Text7.Text = Form4.Text8.Text
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Form4.Text9.Text = Form8.Text4.Text
    Form4.Text7.Text = Form8.Text5.Text
    Form4.Text8.Text = Form8.Text7.Text
    Form4.MSFlexGrid1.Refresh
    Form4.Data1.Refresh
    If Data1.Recordset.BOF Then
        Form4.Text6.Enabled = True
    End If
End Sub
