VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar elementos del presupuesto"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "Eliminar elementos presupuesto.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Concepto"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8055
   End
   Begin VB.Data Data1 
      Caption         =   "Pulse las flechitas de los laterales para recorrer los elementos del preupuesto"
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
      Width           =   7815
   End
   Begin VB.TextBox Text2 
      DataField       =   "TotalIVA"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      DataField       =   "Total"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar elemento"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "Total+iva"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4200
      TabIndex        =   0
      Text            =   "total+iva"
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "IVA:"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Total IVA:"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Total"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Total Pres.:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Total+IVA:"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Data1.Recordset.BOF Then
        MsgBox "No hay registros para borrar", vbInformation, "Información de sistema"
    Else
        Form12.Text4.Text = Form12.Text4.Text - Form12.Text3.Text
        Form12.Text5.Text = Form12.Text5.Text - Form12.Text6.Text
        Form12.Text7.Text = Form12.Text7.Text - Form12.Text2.Text
        Form12.Text4.Text = Round(Form12.Text4.Text, 2)
        Form12.Text5.Text = Round(Form12.Text5.Text, 2)
        Form12.Text7.Text = Round(Form12.Text7.Text, 2)
        Data1.Recordset.Delete
        Data1.Refresh
    End If
End Sub

Private Sub Command2_Click()
    Form11.Text9.Text = Form12.Text4.Text
    Form11.Text7.Text = Form12.Text5.Text
    Form11.Text8.Text = Form12.Text7.Text
    Form11.MSFlexGrid1.Refresh
    Form11.Data1.Refresh
    If Data1.Recordset.BOF Then
        Form11.Text6.Enabled = True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Form12.Text4.Text = Form11.Text9.Text
    Form12.Text5.Text = Form11.Text7.Text
    Form12.Text7.Text = Form11.Text8.Text
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Form11.Text9.Text = Form12.Text4.Text
    Form11.Text7.Text = Form12.Text5.Text
    Form11.Text8.Text = Form12.Text7.Text
    Form11.MSFlexGrid1.Refresh
    Form11.Data1.Refresh
    If Data1.Recordset.BOF Then
        Form11.Text6.Enabled = True
    End If
End Sub
