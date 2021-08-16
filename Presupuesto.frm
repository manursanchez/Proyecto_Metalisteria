VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "Presupuesto.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   1335
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   8055
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         DataField       =   "nombre"
         DataSource      =   "Data5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   23
         Top             =   360
         Width           =   7815
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFFF&
         DataField       =   "direccion"
         DataSource      =   "Data5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   22
         Top             =   840
         Width           =   7815
      End
      Begin VB.Data Data5 
         Caption         =   "Aux clientes"
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
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Presupuesto"
      Height          =   3975
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   8055
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4440
         TabIndex        =   28
         Text            =   "0"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   0
         Top             =   480
         Width           =   7575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   15
         Text            =   "0"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Añadir a Presupuesto"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar un elemento del presupuesto"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   3255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Eliminar todos los elementos del presupuesto"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6240
         TabIndex        =   13
         Text            =   "0"
         Top             =   1440
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "Presupuesto.frx":0442
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2355
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   "          |                                                   |                  |                  |             "
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "IVA"
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
         Left            =   4560
         TabIndex        =   29
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "CONCEPTO"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7800
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Iva:"
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
         Left            =   4320
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "IVA a aplicar:"
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
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Total:"
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
         Left            =   2280
         TabIndex        =   18
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Total Pres."
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
         Left            =   2520
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Total+IVA"
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Panel de control de Presupuestos"
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   6360
      Width           =   8055
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar y salir"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Guardar e Imprimir"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "fecha"
      DataSource      =   "Data5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6600
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Auxiliar de facturas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "auxfacturas"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Data Data2 
      Caption         =   "PRESUPUESTO"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Presupuesto"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data3 
      Caption         =   "CLIENTES PRE"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clientesdepresupuesto"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data4 
      Caption         =   "TOTALES PRE"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totalesdepresupuesto"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "FECHA:"
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
      Left            =   5160
      TabIndex        =   26
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Nº PRESUPUESTO:"
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
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Numero de factura"
      DataField       =   "codfac"
      DataSource      =   "Data5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   24
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'OJO ESTO ES EL CODIGO DE LOS PRESUPUESTOS
Private Sub Command1_Click()
    If Data1.Recordset.BOF = True Then
            MsgBox "No hay presupuesto para guardar ni para imprimir.Añada elementos a presupuesto y guarde o imprima despues", vbInformation, "Error de datos"
            bandera = True
    Else
        guardarfactura
        borrarauxcli
        borrarauxfac
        recuperada = False
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    If Data1.Recordset.BOF = True Then
            MsgBox "No hay presupuesto para guardar ni para imprimir.Añada elementos a presupuesto y guarde o imprima despues", vbInformation, "Error de datos"
    Else
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        guardarfactura
        nfactura = Label3.Caption
        DataEnvironment1.Connection1.Open
        DataEnvironment1.auxcli
        DataReport2.Show vbModal
        DataEnvironment1.Connection1.Close
        Unload Me
    End If
End Sub

Private Sub Command3_Click()
    Load Form12
    Form12.Show vbModal
End Sub

Private Sub Command4_Click()
    Dim opc As Integer
    If Data1.Recordset.BOF = False Then
        opc = MsgBox("¡¡¡¡ATENCION!!!!, Si sale no se guardará el presupuesto, ¿Desea salir sin guardar los cambios, ni el presupuesto?", vbOKCancel, "Informacion de sistema")
        If opc = 1 Then
            Unload Me
        Else
            MsgBox "Para guardar el presupuesto debe pulsar el boton Guardar y salir o Guardar e imprimir", vbInformation, "Informacion de sistema"
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Command5_Click()
    Dim totaliva As Double
    If Text1.Text = "" Or Text2.Text = "" Or Text6.Text = "" Or Text4.Text = "" Then
        MsgBox "Algunas de las casillas CONCEPTO, TotalIVA, IVA o TOTAL estan vacías. Asegurese que introduce informacion en dichas casillas", vbInformation, "Informacion de sistema"
    Else
        totaliva = Text4.Text * Text6.Text / 100
        totaliva = Text4.Text + totaliva
        Data1.Recordset.AddNew
        Data1.Recordset.Fields("numerofactura") = Label3.Caption
        Data1.Recordset.Fields("concepto") = Text1.Text
        Data1.Recordset.Fields("TotalIVA") = Round(Text2.Text, 2)
        Data1.Recordset.Fields("iva") = Text6.Text
        Data1.Recordset.Fields("total") = Round(Text4.Text, 2)
        Data1.Recordset.Fields("total+iva") = Round(totaliva, 2)
        Data1.Recordset.Update
        Data1.Refresh
        Text9.Text = Round(Text9.Text, 2) + Round(Text4.Text, 2)
        Text9.Text = Round(Text9.Text, 2)
        Text7.Text = (Text9.Text * Text6.Text / 100) + Text9.Text
        Text7.Text = Round(Text7.Text, 2)
        Text8.Text = Round(Text8.Text, 2) + Round(Text2.Text, 2)
        Text8.Text = Round(Text8.Text, 2)
        Command6.Enabled = True
        Text1.Text = ""
        Text2.Text = 0
        Text6.Enabled = False
        Text4.Text = 0
    End If
    MSFlexGrid1.Refresh
End Sub

Private Sub Command6_Click()
    If Data1.Recordset.BOF Then
        MsgBox "No hay nada para borrar", vbInformation, "Información de sistema"
    Else
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
        Loop
        Data1.Refresh
        Command6.Enabled = False
        Text7.Text = 0
        Text6.Text = 0
        Text6.Enabled = True
        Text8.Text = 0
        Text9.Text = 0
    End If
End Sub

Private Sub Form_Load()
    If recuperada = True Then
        Form11.Caption = "Presupuesto recuperado"
        Text6.Text = ivarecuperado
        Text6.Enabled = False
        Text9.Text = t
        Text7.Text = tiva
        Text8.Text = STI
    Else
        Label3.Caption = numfactura
        Text3.Text = nombre
        Text5.Text = direccion
        Text10.Text = Date
        bandera = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    borrarauxcli
    borrarauxfac
    recuperada = False
End Sub

'Este validate controla que se introduzca bien la fecha
Private Sub Text10_Validate(Cancel As Boolean)
    If Not IsDate(Text10.Text) Then
        MsgBox "No reconozco este formato de fecha. Utilice el formato dia/mes/año. Por ejemplo: 10/05/2002 sería 10 de mayo del año 2002", vbExclamation, "Error de tipos de datos"
        Cancel = True
    End If
End Sub

'Estos tres validates son para controlar que solo se metan valores numericos
Private Sub Text2_Validate(Cancel As Boolean)
    If Not IsNumeric(Text2.Text) Then
        MsgBox "Solo se aceptan valores numéricos o tipo moneda. Por favor NO introduzca caracteres", vbInformation, "Error de datos"
        Cancel = True
    End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
    If Not IsNumeric(Text4.Text) Then
        MsgBox "Solo se aceptan valores numéricos o tipo moneda. Por favor NO introduzca caracteres", vbInformation, "Error de datos"
        Cancel = True
    End If
End Sub
Private Sub Text4_LostFocus()
    Text2.Text = Round((Text4.Text * Text6.Text) / 100, 2)
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
    If Not IsNumeric(Text6.Text) Then
        MsgBox "Solo se aceptan valores numéricos o tipo moneda. Por favor NO introduzca caracteres", vbInformation, "Error de datos"
        Cancel = True
    End If
End Sub

Private Sub guardarfactura()
        
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
                Data2.Recordset.AddNew
                Data2.Recordset.Fields("codfac") = Data1.Recordset.Fields("numerofactura")
                Data2.Recordset.Fields("fecha") = Val(Text10.Text)
                Data2.Recordset.Fields("concepto") = Data1.Recordset.Fields("concepto")
                Data2.Recordset.Fields("TotalIVA") = Data1.Recordset.Fields("TotalIVA")
                Data2.Recordset.Fields("iva") = Data1.Recordset.Fields("iva")
                Data2.Recordset.Fields("total") = Data1.Recordset.Fields("total")
                Data2.Recordset.Fields("total+iva") = Data1.Recordset.Fields("total+iva")
                Data2.UpdateRecord
                Data1.Recordset.MoveNext
            Loop
            Data2.Refresh
            guardarcliente
      
End Sub
Private Sub guardarcliente()
        
            Data3.Recordset.AddNew
            Data3.Recordset.Fields("codfac") = Val(Label3.Caption)
            Data3.Recordset.Fields("nombre") = Text3.Text
            Data3.Recordset.Fields("direccion") = Text5.Text
            Data3.Recordset.Fields("fecha") = Text10.Text
            Data3.UpdateRecord
            Data3.Refresh
            guardartotales
End Sub
Private Sub guardartotales()
    
        Data4.Recordset.AddNew
        Data4.Recordset.Fields("codfac") = Val(Label3.Caption)
        Data4.Recordset.Fields("total") = Text9.Text
        Data4.Recordset.Fields("total+iva") = Text7.Text
        Data4.Recordset.Fields("sumatotaliva") = Val(Text8.Text)
        Data4.UpdateRecord
        Data4.Refresh
        MsgBox "Presupuesto guardado sin problemas", vbInformation
End Sub
Private Sub borrarauxcli()
    If Not Data5.Recordset.BOF Then
        Data5.Recordset.MoveFirst
        Do While Not Data5.Recordset.EOF
            Data5.Recordset.Delete
            Data5.Recordset.MoveNext
        Loop
        Data5.Refresh
    End If
End Sub
Private Sub borrarauxfac()
    If Not Data1.Recordset.BOF Then
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
        Loop
        Data1.Refresh
    End If
End Sub

