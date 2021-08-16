VERSION 5.00
Begin VB.Form Form18 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCION DE EMPRESA"
   ClientHeight    =   4260
   ClientLeft      =   6285
   ClientTop       =   5175
   ClientWidth     =   6540
   Icon            =   "Empresa.frx":0000
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6540
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command4 
         Caption         =   "SALIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   4815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "METALISTERIA BONILLA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   4815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MARTIN-METAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "METALISTERIA M Y B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    NombreEmpresa = "Metalistería M Y B"
    Empresa = 1
    Load Form1
    Form1.Show vbModal
End Sub

Private Sub Command2_Click()
    NombreEmpresa = "Martín-Metal"
    Empresa = 2
    Load Form1
    Form1.Show vbModal
End Sub

Private Sub Command3_Click()
    NombreEmpresa = "Metalistería Bonilla"
    Empresa = 3
    Load Form1
    Form1.Show vbModal
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub
