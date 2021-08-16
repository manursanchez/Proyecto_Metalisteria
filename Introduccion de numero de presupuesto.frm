VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NUEVO PRESUPUESTO"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "Introduccion de numero de presupuesto.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Introduce el número de presupuesto"
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
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxcli"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Presupuesto"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "Form6"
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
        consulta = "select * from presupuesto where codfac=" + Text1.Text
        Data1.RecordSource = consulta
        Data1.Refresh
        If Data1.Recordset.RecordCount < 1 Then
            numfactura = Val(Text1.Text)
            Load Form10
            Form10.Show vbModal
            Unload Form6
        Else
            MsgBox "El número de presupuesto ya existe,eso significa que ya existe un presupuesto con ese numero y no puede haber dos presupuestos con numeros iguales", vbExclamation, "Error de datos"
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
        MsgBox "Error de datos en tablas auxiliares. Pulse aceptar para continuar.", vbInformation, "Informacion del sistema"
     End If
     If Not Data3.Recordset.BOF Then
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
            Data3.Recordset.Delete
            Data3.Recordset.MoveNext
        Loop
        MsgBox "He encontrado errores en la base de datos.Estos errores son debidos al mal apagado del ordenador. He procedido a su reparación. No obstante, revise los últimos PRESUPUESTOS, por si encuentra algún error. Si encontrara algun error mas, póngase en contacto con el programador.", vbInformation, "Informacion del sistema"
    End If
End Sub
