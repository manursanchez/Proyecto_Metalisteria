VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de ventas"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "Listado de facturas.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Metalisteria M&B\mtl.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informes de ventas"
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame3 
         Caption         =   "Suma total de todas las facturas"
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   5295
         Begin VB.CommandButton Command3 
            Caption         =   "Ver suma total"
            Height          =   615
            Left            =   3120
            TabIndex        =   3
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "Total+IVA:"
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
            TabIndex        =   17
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
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
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   1320
            TabIndex        =   15
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Introduce el intervalo "
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5295
         Begin VB.CommandButton Command2 
            Caption         =   "Calcular"
            Height          =   615
            Left            =   4080
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   1
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   3000
            TabIndex        =   14
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Total ventas + IVA"
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
            Left            =   2760
            TabIndex        =   12
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Total Ventas"
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
            Left            =   360
            TabIndex        =   11
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Introduce la segunda fecha:"
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
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Introduce la primera fecha:"
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
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim f As Date
    Dim f2 As Date
    Dim sumatotal As Double
    Dim sumatotaliva As Double
    If Text1.Text = "" Or Text2.Text = "" Or Not IsDate(Text1.Text) Or Not IsDate(Text2.Text) Then
        MsgBox "El formato de fecha no es correcto por favor intentelo de nuevo. El formato de fecha correcto es dd/mm/aaaa. Ejemplo: 31/12/2002", vbCritical, "Error de datos"
    Else
        f = Text1.Text
        f2 = Text2.Text
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            If Data1.Recordset.Fields("fecha") >= f And Data1.Recordset.Fields("fecha") <= f2 Then
                sumatotal = sumatotal + Data1.Recordset.Fields("total")
                sumatotaliva = sumatotaliva + Data1.Recordset.Fields("total+iva")
                Data1.Recordset.MoveNext
            Else
                Data1.Recordset.MoveNext
            End If
        Loop
        Label6.Caption = Round(sumatotaliva, 2)
        Label5.Caption = Round(sumatotal, 2)
    End If
End Sub


Private Sub Command3_Click()
    Dim sumatotal As Double
    Dim sumatotaliva As Double
    If Data1.Recordset.BOF Then
        MsgBox "No hay registros para sumar", vbInformation
    Else
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            sumatotal = sumatotal + Data1.Recordset.Fields("total")
            sumatotaliva = sumatotaliva + Data1.Recordset.Fields("total+iva")
            Data1.Recordset.MoveNext
        Loop
        Label7.Caption = Round(sumatotal, 2)
        Label8.Caption = Round(sumatotaliva, 2)
    End If
End Sub
