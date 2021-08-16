VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de..."
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "Acerca de....frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Telefono: 607544555     Correo Electronico: mrodriguezs@eresmas.com"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Técnico especialista en Informática de Gestión"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "MANUEL RODRIGUEZ SANCHEZ"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "APLICACION PERSONALIZADA PARA LA REALIZACION DE FACTURAS Y PRESUPUESTOS PARA LA EMPRESA ""METALISTERIA MYB"" "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Programación de la aplicación a cargo de:"
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
      TabIndex        =   0
      Top             =   720
      Width           =   6735
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
