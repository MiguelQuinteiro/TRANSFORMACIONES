VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTransformaciones 
   BackColor       =   &H00FFC0C0&
   Caption         =   "TRANSFORMACIONES"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   14595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIntercambiar 
      BackColor       =   &H00FF8080&
      Caption         =   "INTERCAMBIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdLimpiar 
      BackColor       =   &H00FF8080&
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Caption         =   " TRANSFORMACIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   11520
      TabIndex        =   90
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton cmdCanonica 
         BackColor       =   &H00FF8080&
         Caption         =   "CANONICA"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   9240
         Width           =   2415
      End
      Begin VB.CommandButton cmdGiro90 
         BackColor       =   &H00FF8080&
         Caption         =   "GIRO 90º"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton cmdGiro180 
         BackColor       =   &H00FF8080&
         Caption         =   "GIRO 180º"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton cmdGiro270 
         BackColor       =   &H00FF8080&
         Caption         =   "GIRO 270º"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton cmdFilas12 
         BackColor       =   &H00FF8080&
         Caption         =   "FILAS 1-2"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton cmdFilas34 
         BackColor       =   &H00FF8080&
         Caption         =   "FILAS 3-4"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CommandButton cmdFilas1234 
         BackColor       =   &H00FF8080&
         Caption         =   "FILAS 1-2,3-4"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   3480
         Width           =   2415
      End
      Begin VB.CommandButton cmdColumnas12 
         BackColor       =   &H00FF8080&
         Caption         =   "COLUMNAS 1-2"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   4080
         Width           =   2415
      End
      Begin VB.CommandButton cmdColumnas34 
         BackColor       =   &H00FF8080&
         Caption         =   "COLUMNAS 3-4"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4560
         Width           =   2415
      End
      Begin VB.CommandButton cmdColumnas1234 
         BackColor       =   &H00FF8080&
         Caption         =   "COLUMNAS 1-2,3-4"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   5040
         Width           =   2415
      End
      Begin VB.CommandButton cmdNivelesTorres 
         BackColor       =   &H00FF8080&
         Caption         =   "NIVELES - TORRES"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   6600
         Width           =   2415
      End
      Begin VB.CommandButton cmdTorres 
         BackColor       =   &H00FF8080&
         Caption         =   "TORRES"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   6120
         Width           =   2415
      End
      Begin VB.CommandButton cmdNiveles 
         BackColor       =   &H00FF8080&
         Caption         =   "NIVELES"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   5640
         Width           =   2415
      End
      Begin VB.CommandButton cmdHorizontal 
         BackColor       =   &H00FF8080&
         Caption         =   "HORIZONTAL"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   7200
         Width           =   2415
      End
      Begin VB.CommandButton cmdVertical 
         BackColor       =   &H00FF8080&
         Caption         =   "VERTICAL"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   7680
         Width           =   2415
      End
      Begin VB.CommandButton cmdTransponerDerecha 
         BackColor       =   &H00FF8080&
         Caption         =   "TRANSPONER   \"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   8280
         Width           =   2415
      End
      Begin VB.CommandButton cmdTransponerIzquierda 
         BackColor       =   &H00FF8080&
         Caption         =   "TRANSPONER   /"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   8760
         Width           =   2415
      End
      Begin VB.CommandButton cmdTransposicion 
         BackColor       =   &H00FF8080&
         Caption         =   "TRANSPOSICION"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CAMBIOS A FAMILIAS 1 Y 2 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   69
      Top             =   8520
      Width           =   11175
      Begin VB.CommandButton cmdConvierte3a2 
         BackColor       =   &H00FF8080&
         Caption         =   "3 a 2"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte4a2 
         BackColor       =   &H00FF8080&
         Caption         =   "4 a 2"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte5a1 
         BackColor       =   &H00FF8080&
         Caption         =   "5 a 1"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte6a2 
         BackColor       =   &H00FF8080&
         Caption         =   "6 a 2"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte7a2 
         BackColor       =   &H00FF8080&
         Caption         =   "7 a 2"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte8a1 
         BackColor       =   &H00FF8080&
         Caption         =   "8 a 1"
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte9a2 
         BackColor       =   &H00FF8080&
         Caption         =   "9 a 2"
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte10a2 
         BackColor       =   &H00FF8080&
         Caption         =   "10 a 2"
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
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte11a2 
         BackColor       =   &H00FF8080&
         Caption         =   "11 a 2"
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
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte12a1 
         BackColor       =   &H00FF8080&
         Caption         =   "12 a 1"
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte2a3 
         BackColor       =   &H00FF8080&
         Caption         =   "2 a 3"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte2a4 
         BackColor       =   &H00FF8080&
         Caption         =   "2 a 4"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte1a5 
         BackColor       =   &H00FF8080&
         Caption         =   "1 a 5"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte2a6 
         BackColor       =   &H00FF8080&
         Caption         =   "2 a 6"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte2a7 
         BackColor       =   &H00FF8080&
         Caption         =   "2 a 7"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte1a8 
         BackColor       =   &H00FF8080&
         Caption         =   "1 a 8"
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte2a9 
         BackColor       =   &H00FF8080&
         Caption         =   "2 a 9"
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte2a10 
         BackColor       =   &H00FF8080&
         Caption         =   "2 a 10"
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
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte2a11 
         BackColor       =   &H00FF8080&
         Caption         =   "2 a 11"
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
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdConvierte1a12 
         BackColor       =   &H00FF8080&
         Caption         =   "1 a 12"
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdTodosLosProblemas 
      BackColor       =   &H00FF8080&
      Caption         =   "TODOS LOS PROBLEMAS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdAnalizarTransformada 
      BackColor       =   &H00FF8080&
      Caption         =   "ANALIZAR TRANSFORMADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   7320
      Width           =   2895
   End
   Begin VB.CommandButton cmdAnalizarOriginal 
      BackColor       =   &H00FF8080&
      Caption         =   "ANALIZAR ORIGINAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TRANSFORMACIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   3240
      TabIndex        =   46
      Top             =   3000
      Width           =   2895
      Begin VB.TextBox txtSolTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtModTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   51
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtGrilla01Trans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   50
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtGrilla02Trans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   49
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtGrilla03Trans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   48
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtGrilla04Trans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   47
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Grilla 04"
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
         TabIndex        =   65
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Grilla 03"
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
         TabIndex        =   64
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Grilla 02"
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
         TabIndex        =   63
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Grilla 01"
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
         TabIndex        =   62
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Modelo"
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
         TabIndex        =   61
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Solución"
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
         TabIndex        =   60
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   " ORIGINAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   39
      Top             =   3000
      Width           =   2895
      Begin VB.TextBox txtGrilla04Original 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   45
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtGrilla03Original 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   44
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtGrilla02Original 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   43
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtGrilla01Original 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   42
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtModOriginal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   41
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSolOriginal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   40
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Grilla 04"
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
         TabIndex        =   59
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Grilla 03"
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
         TabIndex        =   58
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Grilla 02"
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
         TabIndex        =   57
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Grilla 01"
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
         TabIndex        =   56
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Modelo"
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
         TabIndex        =   55
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Solución"
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
         TabIndex        =   54
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame framLineaProblemas 
      BackColor       =   &H00FFC0C0&
      Caption         =   " CARGAR O EXTRAER PROBLEMAS   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6360
      TabIndex        =   35
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdExtraeProblema 
         BackColor       =   &H00FF8080&
         Caption         =   "EXTRAER PROBLEMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtCargaProblema 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   37
         Text            =   "1234341221434321"
         Top             =   360
         Width           =   4575
      End
      Begin VB.CommandButton cmdCargaProblema 
         BackColor       =   &H00FF8080&
         Caption         =   "CARGAR PROBLEMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TRANSFORMACIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   1
      Left            =   3240
      TabIndex        =   18
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   34
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   1920
         TabIndex        =   33
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1440
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1920
         TabIndex        =   30
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   840
         TabIndex        =   28
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1440
         TabIndex        =   27
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1920
         TabIndex        =   26
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   360
         TabIndex        =   25
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   840
         TabIndex        =   24
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   1440
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   1920
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   840
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtTrans 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   1440
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " ORIGINAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   1440
         TabIndex        =   17
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   840
         TabIndex        =   16
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   1920
         TabIndex        =   14
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   1440
         TabIndex        =   13
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   840
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtProblema 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   1920
         TabIndex        =   2
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdTransfomacionTotal 
      BackColor       =   &H00FF8080&
      Caption         =   "TRANSFORMACION TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4095
      Left            =   6360
      TabIndex        =   53
      Top             =   3120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmTransformaciones.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTransformaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : TRANSFORMACIONES
'* CONTENIDO     : ESTUDIA LAS TRANSFORMACIONES A UNA SOLUCION DE SUDOKU
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 18 DE JULIO DE 2014
'* ACTUALIZACION : 18 DE JULIO DE 2014
'****************************************************************************************
Option Explicit

' LAS 288 SOLUCIONES QUE EXISTEN PARA EL SUDOKU DE 4X4
Private Type misSoluciones
    miNumero As Long
    miCasilla(1 To 16) As Integer
End Type

' DECLARACIÓN DE VARIABLES CON TIPOS
Dim miSolucion(1 To 288) As misSoluciones

' DECLARACIÓN DE VARIABLES LOCALES
Dim miLineInput As String

Dim miVectorAnalisis(1 To 16) As Integer

Dim miSolucionEnEstudio As Integer
Dim miGrilla01EnEstudio As Integer
Dim miGrilla02EnEstudio As Integer
Dim miGrilla03EnEstudio As Integer
Dim miGrilla04EnEstudio As Integer
Dim miModeloEnEstudio As Integer
Dim mi_PrintLine As String

Private Sub cmdAnalizarOriginal_Click()
    ' DECLARACIÓN DE VARIABLES PRIVADAS
    Dim i As Integer
    Dim miCadenaAnalisis As String
    
    ' CARGA LOS DATOS PARA ANALIZARLOS
    miCadenaAnalisis = ""
    For i = 1 To 16
        miVectorAnalisis(i) = Val(txtProblema(i).Text)
        miCadenaAnalisis = miCadenaAnalisis + Trim(txtProblema(i).Text)
    Next i
    
    If miCadenaAnalisis = "" Then
        ' MOSTRAR LOS RESULTADOS TOTALES
        txtSolOriginal = ""
        txtGrilla01Original = ""
        txtGrilla02Original = ""
        txtGrilla03Original = ""
        txtGrilla04Original = ""
        txtModOriginal = ""
    Else
        ' DETERMINAR NUMERO DE SOLUCION
        Call miDeterminaSolucion
        
        If miSolucionEnEstudio <> 0 Then
            ' DETERMINAR EL NUMERO DE LAS GRILLAS QUE INTERVIENEN
            Call miDeterminaGrilla01
            Call miDeterminaGrilla02
            Call miDeterminaGrilla03
            Call miDeterminaGrilla04
            
            ' DETERMINAR MODELO AL QUE PERTENECE
            Call miDeterminaModelo
            
            ' MOSTRAR LOS RESULTADOS TOTALES
            txtSolOriginal = miSolucionEnEstudio
            txtGrilla01Original = miGrilla01EnEstudio
            txtGrilla02Original = miGrilla02EnEstudio
            txtGrilla03Original = miGrilla03EnEstudio
            txtGrilla04Original = miGrilla04EnEstudio
            txtModOriginal = miModeloEnEstudio
        Else
            ' MOSTRAR LOS RESULTADOS TOTALES
            txtSolOriginal = 0
            txtGrilla01Original = 0
            txtGrilla02Original = 0
            txtGrilla03Original = 0
            txtGrilla04Original = 0
            txtModOriginal = 0
        End If
    End If
End Sub

Private Sub cmdAnalizarTransformada_Click()
    ' DECLARACIÓN DE VARIABLES PRIVADAS
    Dim i As Integer
    Dim miCadenaAnalisis As String
    
    ' CARGA LOS DATOS PARA ANALIZARLOS
    miCadenaAnalisis = ""
    For i = 1 To 16
        miVectorAnalisis(i) = Val(txtTrans(i).Text)
        miCadenaAnalisis = miCadenaAnalisis + Trim(txtTrans(i).Text)
    Next i
    
    If miCadenaAnalisis = "" Then
        ' MOSTRAR LOS RESULTADOS TOTALES
        txtSolTrans = ""
        txtGrilla01Trans = ""
        txtGrilla02Trans = ""
        txtGrilla03Trans = ""
        txtGrilla04Trans = ""
        txtModTrans = ""
    Else
        ' DETERMINAR NUMERO DE SOLUCION
        Call miDeterminaSolucion
        
        If miSolucionEnEstudio <> 0 Then
            ' DETERMINAR EL NUMERO DE LAS GRILLAS QUE INTERVIENEN
            Call miDeterminaGrilla01
            Call miDeterminaGrilla02
            Call miDeterminaGrilla03
            Call miDeterminaGrilla04
            
            ' DETERMINAR MODELO AL QUE PERTENECE
            Call miDeterminaModelo
            
            ' MOSTRAR LOS RESULTADOS TOTALES
            txtSolTrans = miSolucionEnEstudio
            txtGrilla01Trans = miGrilla01EnEstudio
            txtGrilla02Trans = miGrilla02EnEstudio
            txtGrilla03Trans = miGrilla03EnEstudio
            txtGrilla04Trans = miGrilla04EnEstudio
            txtModTrans = miModeloEnEstudio
        Else
            ' MOSTRAR LOS RESULTADOS TOTALES
            txtSolTrans = 0
            txtGrilla01Trans = 0
            txtGrilla02Trans = 0
            txtGrilla03Trans = 0
            txtGrilla04Trans = 0
            txtModTrans = 0
        End If
    End If
End Sub

' CALCULA LA MATRIZ DE COMBINACION DE LAS 16 X 16 TRANSFORMACIONES (256)
Private Sub CalculaDerivados4096(miEnviado As String)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim miResultadoI As String
    Dim miResultadoJ As String
    Dim miResultadoK As String
    
    For i = 1 To 16
        For j = 1 To 16
        For k = 1 To 16
            'Aplicar transformacion(i)
            Select Case i
                Case 1
                    miResultadoI = Giro90(miEnviado)
                Case 2
                    miResultadoI = Giro180(miEnviado)
                Case 3
                    miResultadoI = Giro270(miEnviado)
                Case 4
                    miResultadoI = Filas12(miEnviado)
                Case 5
                    miResultadoI = Filas34(miEnviado)
                Case 6
                    miResultadoI = Filas1234(miEnviado)
                Case 7
                    miResultadoI = Columnas12(miEnviado)
                Case 8
                    miResultadoI = Columnas34(miEnviado)
                Case 9
                    miResultadoI = Columnas1234(miEnviado)
                Case 10
                    miResultadoI = Niveles(miEnviado)
                Case 11
                    miResultadoI = Torres(miEnviado)
                Case 12
                    miResultadoI = NivelesTorres(miEnviado)
                Case 13
                    miResultadoI = Horizontal(miEnviado)
                Case 14
                    miResultadoI = Vertical(miEnviado)
                Case 15
                    miResultadoI = TransponerIzquierda(miEnviado)
                Case 16
                    miResultadoI = TransponerDerecha(miEnviado)
            End Select
                    
            'Aplicar transformacion(j) al resultado de (i)
            Select Case j
                Case 1
                    miResultadoJ = Giro90(miResultadoI)
                Case 2
                    miResultadoJ = Giro180(miResultadoI)
                Case 3
                    miResultadoJ = Giro270(miResultadoI)
                Case 4
                    miResultadoJ = Filas12(miResultadoI)
                Case 5
                    miResultadoJ = Filas34(miResultadoI)
                Case 6
                    miResultadoJ = Filas1234(miResultadoI)
                Case 7
                    miResultadoJ = Columnas12(miResultadoI)
                Case 8
                    miResultadoJ = Columnas34(miResultadoI)
                Case 9
                    miResultadoJ = Columnas1234(miResultadoI)
                Case 10
                    miResultadoJ = Niveles(miResultadoI)
                Case 11
                    miResultadoJ = Torres(miResultadoI)
                Case 12
                    miResultadoJ = NivelesTorres(miResultadoI)
                Case 13
                    miResultadoJ = Horizontal(miResultadoI)
                Case 14
                    miResultadoJ = Vertical(miResultadoI)
                Case 15
                    miResultadoJ = TransponerIzquierda(miResultadoI)
                Case 16
                    miResultadoJ = TransponerDerecha(miResultadoI)
            End Select
            
            'Aplicar transformacion(j) al resultado de (i)
            Select Case k
                Case 1
                    miResultadoK = Giro90(miResultadoJ)
                Case 2
                    miResultadoK = Giro180(miResultadoJ)
                Case 3
                    miResultadoK = Giro270(miResultadoJ)
                Case 4
                    miResultadoK = Filas12(miResultadoJ)
                Case 5
                    miResultadoK = Filas34(miResultadoJ)
                Case 6
                    miResultadoK = Filas1234(miResultadoJ)
                Case 7
                    miResultadoK = Columnas12(miResultadoJ)
                Case 8
                    miResultadoK = Columnas34(miResultadoJ)
                Case 9
                    miResultadoK = Columnas1234(miResultadoJ)
                Case 10
                    miResultadoK = Niveles(miResultadoJ)
                Case 11
                    miResultadoK = Torres(miResultadoJ)
                Case 12
                    miResultadoK = NivelesTorres(miResultadoJ)
                Case 13
                    miResultadoK = Horizontal(miResultadoJ)
                Case 14
                    miResultadoK = Vertical(miResultadoJ)
                Case 15
                    miResultadoK = TransponerIzquierda(miResultadoJ)
                Case 16
                    miResultadoK = TransponerDerecha(miResultadoJ)
            End Select
            
            'Imprimir resultado de (j)
            Print #77, miResultadoK
        Next k
        Next j
    Next i
End Sub


' CONVIERTE PROBLEMA DE UN MODELO A SU ISOMORFICO EN LA FAMILIA ORIGINAL
Private Sub cmdConvierte3a2_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = TransponerDerecha(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte4a2_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Giro90(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte5a1_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte6a2_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas34(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte7a2_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte8a1_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte9a2_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Giro270(miDato)
    miResultado = Filas34(miResultado)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte10a2_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    miResultado = TransponerIzquierda(miResultado)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte11a2_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte12a1_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    miResultado = Filas12(miResultado)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub



' CONVIERTE PROBLEMA DE LA FAMILIA ORIGINAL A SU ISOMORFICO DEL UN MODELO CORRESPONDIENTE
Private Sub cmdConvierte2a3_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = TransponerDerecha(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte2a4_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Giro90(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte1a5_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte2a6_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas34(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte2a7_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte1a8_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte2a9_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Giro90(miDato)
    miResultado = Columnas12(miResultado)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte2a10_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas12(miDato)
    miResultado = Giro90(miResultado)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte2a11_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdConvierte1a12_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    miResultado = Filas12(miResultado)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdIntercambiar_Click()
    Dim i As Integer
    Dim miAuxiliar As String
    
    For i = 1 To 16
        miAuxiliar = txtProblema(i).Text
        txtProblema(i).Text = txtTrans(i).Text
        txtTrans(i).Text = miAuxiliar
    Next i
    
    Call cmdAnalizarOriginal_Click
    Call cmdAnalizarTransformada_Click

End Sub

' LIMPIA LOS DATOS DE LOS DISPLAY
Private Sub cmdLimpiar_Click()
    Dim i As Integer
    For i = 1 To 16
        txtProblema(i).Text = ""
        txtTrans(i).Text = ""
    Next i
    txtCargaProblema = ""
    txtSolOriginal = ""
    txtModOriginal = ""
    txtGrilla01Original = ""
    txtGrilla02Original = ""
    txtGrilla03Original = ""
    txtGrilla04Original = ""
    txtSolTrans = ""
    txtModTrans = ""
    txtGrilla01Trans = ""
    txtGrilla02Trans = ""
    txtGrilla03Trans = ""
    txtGrilla04Trans = ""
    RichTextBox1 = ""
End Sub

' GENERA UN ARCHIVO CON TODOS LOS PROBLEMAS DE 4 PISTAS
Private Sub cmdTodosLosProblemas_Click()
    Dim miLineInput As String
    Dim miLineOutput As String
    
    Dim miDato As String
    
    Open "Todos los Problemas por Familias de Modelos.txt" For Input As #77
    Open "TodosLosProblemas4Pistas.txt" For Output As #78

    Do Until EOF(77)
        Line Input #77, miLineInput
        miDato = Trim(miLineInput)
        
        Print #78, miDato
        Call Transposicion(miDato)
    Loop
    
    Close #77
    Close #77
End Sub

' ENCUENTRA CADA UNA DE LAS 24 TRANSPOSICIONES DE UN PROBLEMA O SOLUCION
Private Sub cmdTransposicion_Click()
    ' DECLARACIÓN DE VARIABLES PRIVADAS
    Dim miProblemaOriginal As String
    Dim miProblemaDerivado As String
    Dim miProblemaTransformado As String
    Dim miPrimero As Integer
    Dim miSegundo As Integer
    Dim miTercero As Integer
    Dim miCuarto As Integer
    Dim miAuxiliar As String
    Dim x As Integer
    
    Open "LosDerivadosRepetidos.txt" For Output As #77
    Open "LosDerivadosUnicos.txt" For Output As #78

    ' INICIALIZACIÓN DE VARIABLES PRIVADAS
    miProblemaOriginal = ExtraerProblema()
    miProblemaDerivado = ""
    miProblemaTransformado = ""
    miAuxiliar = ""
    
    ' BUSQUEDA DE LAS TRANSPOSICIONES
    For miPrimero = 1 To 4
        For miSegundo = 1 To 4
            If miPrimero <> miSegundo Then
            
                For miTercero = 1 To 4
                    If miPrimero <> miSegundo And _
                           miPrimero <> miTercero And _
                           miSegundo <> miTercero Then
                    
                        For miCuarto = 1 To 4
                            If miPrimero <> miCuarto And _
                               miSegundo <> miCuarto And _
                               miTercero <> miCuarto Then
                                
                                ' CAMBIO DE DATOS POR SU TRANSPOSICION CORRESPONDIENTE
                                For x = 1 To 16
                                    If Val(Mid(miProblemaOriginal, x, 1)) = 0 Then
                                        miProblemaDerivado = miProblemaDerivado + "0"
                                    End If
                                    If Val(Mid(miProblemaOriginal, x, 1)) = 1 Then
                                        miProblemaDerivado = miProblemaDerivado + Trim(Str(miPrimero))
                                    End If
                                    If Val(Mid(miProblemaOriginal, x, 1)) = 2 Then
                                        miProblemaDerivado = miProblemaDerivado + Trim(Str(miSegundo))
                                    End If
                                    If Val(Mid(miProblemaOriginal, x, 1)) = 3 Then
                                        miProblemaDerivado = miProblemaDerivado + Trim(Str(miTercero))
                                    End If
                                    If Val(Mid(miProblemaOriginal, x, 1)) = 4 Then
                                        miProblemaDerivado = miProblemaDerivado + Trim(Str(miCuarto))
                                    End If
                                Next x
                                
                                ' MUESTRS LOS RESULTADOS EN EL FORMULARIO
                                RichTextBox1.Text = RichTextBox1.Text + miProblemaDerivado + vbCr
                                
                                ' IMPRIME EN EL ARCHIVO LOS PROBLEMAS DERIVADOS
                                Print #77, miProblemaDerivado
                                
                                ' CALCULA TODOS LOS POSIBLES PARA EL PROBLEMA DERIVADO
                                Call CalculaDerivados4096(miProblemaDerivado)
                                
                                miProblemaDerivado = ""
                                
                            End If
                        Next miCuarto
                    End If
                Next miTercero
            End If
        Next miSegundo
    Next miPrimero
    
    Close #77
    Close #78
    
    ' CLACULA EL TOTAL DE DERIVADOS
    Call DerivadosTotal
End Sub

' CLACULA EL TOTAL DE DERIVADOS
Private Sub DerivadosTotal()
    Dim miLineInput As String
    Dim miDato As String
    Dim miContadorProblemasUnicos
    Dim ProblemasUnicos(30000) As String
    Dim NoEsUnico As Boolean
    Dim x As Long
    
    Open "LosDerivadosRepetidos.txt" For Input As #77
    Open "LosDerivadosUnicos.txt" For Output As #78

    miContadorProblemasUnicos = 1
    
    Do Until EOF(77)
        Line Input #77, miLineInput
        miDato = Trim(miLineInput)
        
        NoEsUnico = False
        For x = 1 To miContadorProblemasUnicos
            If x = 1 And ProblemasUnicos(x) = "" Then
                'primera vez
                ProblemasUnicos(x) = miDato
                miContadorProblemasUnicos = miContadorProblemasUnicos + 1
            End If
            
            If miDato = ProblemasUnicos(x) Then
                NoEsUnico = True
                ' condicion de salida
                x = miContadorProblemasUnicos + 1
            End If
        Next x
        
        If NoEsUnico = False Then
            ' como es unico lo agrgo a la lista
            ProblemasUnicos(miContadorProblemasUnicos) = miDato
            miContadorProblemasUnicos = miContadorProblemasUnicos + 1
        End If
    Loop
    
    For x = 1 To miContadorProblemasUnicos
        If ProblemasUnicos(x) <> "" Then
            Print #78, x, "   ", ProblemasUnicos(x)
        End If
    Next x

    Close #77
    Close #78
End Sub

' AL CARGAR EL FORMULARIO
Private Sub Form_Load()
    ' DECLARACIÓN DE VARIABLES PRIVADAS
    Dim i As Integer
    
    ' CARGA VALORES INICIALES PARA LAS SOLUCIONES
    ' ABRE EL ARCHIVO CON LAS 288 SOLUCIONES
    Dim x As Integer
    Dim miNumero As Integer
    Open "miSolucionesTotal.txt" For Input As #10
    
    Do Until EOF(10)
        Line Input #10, miLineInput
        miNumero = Val(Mid(miLineInput, 32, 3))
        
        ' CARGA EL NUMERO DE LA GRILLA
        miSolucion(miNumero).miNumero = Val(Mid(miLineInput, 32, 3))
        
        ' CARGA LOS VALORES DE LOS DÍGITOS QUE COMPONEN LA SOLUCION
        For x = 1 To 16
            miSolucion(miNumero).miCasilla(x) = Val(Mid(miLineInput, x, 1))
        Next x
    Loop
    
    Close #10
End Sub

' APLICA TODAS LAS TRANSFORMACIONES
Private Sub cmdTransfomacionTotal_Click()
    Dim i As Integer
    Dim j As Integer
    
    Open "LasTransformaciones.txt" For Output As #1
    
    ' PARA SOLUCIONES = 1 TO 288
    For i = 1 To 288
        ' COLOCAR LA SOLUCION EN LA LINEA DE CARGA
        txtCargaProblema = ""
        For j = 1 To 16
            txtCargaProblema = txtCargaProblema + Trim(Str(miSolucion(i).miCasilla(j)))
        Next j
    
        ' CARGAR EL PROBLEMA
        Call cmdCargaProblema_Click
        
        ' ANALIZAR
        Call Analizar
        
        ' MUESTRA EN PANTALLA LOS RESULTADOS
        DoEvents
        
    ' SIGUIENTE PARA
    Next i
    Close #1
End Sub

' CARGA PROBLEMA DESDE LA LINEA
Private Sub cmdCargaProblema_Click()
    CargarProblema (txtCargaProblema)
End Sub

' EXTRAE PROBLEMA HACIA LA LINEA
Private Sub cmdExtraeProblema_Click()
    txtCargaProblema = ExtraerProblema()
End Sub

' CARGA PROBLEMA QUE SE PRESENTAN EN FORMA DE LINEA
Private Function CargarProblema(miEnviado As String)
    Dim x As Integer
    Call Limpia
    For x = 1 To 16
        If Val(Mid(miEnviado, x, 1)) = 0 Then
            txtProblema(x) = ""
        Else
            txtProblema(x) = Mid(miEnviado, x, 1)
        End If
    Next x
End Function

' EXTRAER PROBLEMA HACIA UNA VARIABLE DE TIPO CARACTER
Private Function ExtraerProblema() As String
    Dim x As Integer
    Dim ProblemaExtraido As String
    ProblemaExtraido = ""
    For x = 1 To 16
        If txtProblema(x) = "" Then
            ProblemaExtraido = ProblemaExtraido & "0"
        Else
            ProblemaExtraido = ProblemaExtraido & txtProblema(x)
        End If
    Next x
    ExtraerProblema = ProblemaExtraido
End Function

' LIMPIA LOS VALORES VISIBLES EN TODO EL FORMULARIO
Private Sub Limpia()
    Dim x As Integer
    Dim i As Integer
    For x = 1 To 16
        txtProblema(x) = ""
        txtProblema(x).Enabled = True
    Next x
End Sub

' ANALIZAR LA SOLUCIÓN PLANTEADA PARA BUSCAR TODAS SUS TRANSFORMACIONES POSIBLES
Private Sub Analizar()
    Call miCargaVector
    
    ' DETERMINAR NUMERO DE SOLUCION
    Call miDeterminaSolucion
    
    ' DETERMINAR EL NUMERO DE LAS GRILLAS QUE INTERVIENEN
    Call miDeterminaGrilla01
    Call miDeterminaGrilla02
    Call miDeterminaGrilla03
    Call miDeterminaGrilla04
    
    ' DETERMINAR MODELO AL QUE PERTENECE
    Call miDeterminaModelo
    
    ' MOSTRAR LOS RESULTADOS TOTALES
    txtSolOriginal = miSolucionEnEstudio
    txtGrilla01Original = miGrilla01EnEstudio
    txtGrilla02Original = miGrilla02EnEstudio
    txtGrilla03Original = miGrilla03EnEstudio
    txtGrilla04Original = miGrilla04EnEstudio
    txtModOriginal = miModeloEnEstudio
        
    ' IMPRIMIR LOS RESULTADOS A UN ARCHIVO DE TEXTO
    Call ImprimeArchivo(" -- Original")
    
    '1
    ' GIRO 90º
    Call cmdGiro90_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Giro 90º")
    
    ' GIRO 180º
    Call cmdGiro180_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Giro 180º")
    
    ' GIRO 270º
    Call cmdGiro270_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Giro 270º")
    
    
    ' INTERCAMBIAR FILAS (1-2)
    Call cmdFilas12_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Filas 1-2")
    
    ' INTERCAMBIAR FILAS (3-4)
    Call cmdFilas34_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Filas 3-4")
    
    ' INTERCAMBIAR FILAS (1-2, 3-4)
    Call cmdFilas1234_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Filas 1-2, 3-4")
    
    
    ' INTERCAMBIAR COLUMNAS (1-2)
    Call cmdColumnas12_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Columnas 1-2")
    
    ' INTERCAMBIAR COLUMNAS (3-4)
    Call cmdColumnas34_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Columnas 3-4")
    
    ' INTERCAMBIAR COLUMNAS (1-2, 3-4)
    Call cmdColumnas1234_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Columnas 1-2, 3-4")
    
    
    ' INTERCAMBIAR REGIONES (1,2 - 3,4)
    Call cmdNiveles_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Niveles")
    
    ' INTERCAMBIAR REGIONES (1,3 - 2,4)
    Call cmdTorres_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Torres")
    
    ' INTERCAMBIAR REGIONES (1,2 - 3,4) - (1,3 - 2,4)
    Call cmdNivelesTorres_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Niveles y Torres")
    
    ' REFLEJAR HORIZONTAL
    Call cmdHorizontal_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Horizontal")
    
    ' REFLEJAR VERTICAL
    Call cmdVertical_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Vertical")
    
    ' TRANSPONER IZQUIERDA (\)
    Call cmdTransponerIzquierda_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Trasponer (\)")
    
    ' TRANSPONER DERECHA (/)
    Call cmdTransponerDerecha_Click
    Call cmdAnalizarTransformada_Click
    Call ImprimeArchivo(" -- Trasponer (/)")
    
    Print #1, ""
    Print #1, ""
    
    ' ACUMULAR LOS RESULTADOS
    'miContadorTotalEstudiados = miContadorTotalEstudiados + Val(lblAnalizados)
    'miContadorTotalEncontrados = miContadorTotalEncontrados + Val(lblEncontrados)
End Sub

' CARGA LOS DATOS EN EL VECTOR DE ANALISIS
Private Sub miCargaVector()
    Dim i As Integer
    
    ' CARGA LOS DATOS PARA ANALIZARLOS
    For i = 1 To 16
        miVectorAnalisis(i) = Val(txtProblema(i).Text)
    Next i
End Sub

' DETERMINA EL NUMERO DE LA SOLUCION QUE SE ESTA UTILIZANDO
Private Sub miDeterminaSolucion()
    Dim i As Integer
    Dim miDato As String
    
    miDato = ""
    
    For i = 1 To 16
        miDato = miDato + Trim(Str(miVectorAnalisis(i)))
    Next i
    
    miSolucionEnEstudio = DeterminaNumeroSolucion(miDato)
End Sub

' DETERMINA EL NUMERO DEL MODELO QUE SE ESTA UTILIZANDO
Private Sub miDeterminaModelo()
    ' 01 07 10 16   nº:  1
    If miGrilla01EnEstudio = 1 And _
       miGrilla02EnEstudio = 7 And _
       miGrilla03EnEstudio = 10 And _
       miGrilla04EnEstudio = 16 Then
           miModeloEnEstudio = 1
    End If
    
    ' 01 07 12 14   nº:  2
    If miGrilla01EnEstudio = 1 And _
       miGrilla02EnEstudio = 7 And _
       miGrilla03EnEstudio = 12 And _
       miGrilla04EnEstudio = 14 Then
           miModeloEnEstudio = 2
    End If
    
    ' 01 08 10 15   nº:  3
    If miGrilla01EnEstudio = 1 And _
       miGrilla02EnEstudio = 8 And _
       miGrilla03EnEstudio = 10 And _
       miGrilla04EnEstudio = 15 Then
           miModeloEnEstudio = 3
    End If
    
    ' 02 07 09 16   nº:  4
    If miGrilla01EnEstudio = 2 And _
       miGrilla02EnEstudio = 7 And _
       miGrilla03EnEstudio = 9 And _
       miGrilla04EnEstudio = 16 Then
           miModeloEnEstudio = 4
    End If
    
    ' 02 08 09 15   nº:  5
    If miGrilla01EnEstudio = 2 And _
       miGrilla02EnEstudio = 8 And _
       miGrilla03EnEstudio = 9 And _
       miGrilla04EnEstudio = 15 Then
           miModeloEnEstudio = 5
    End If
    
    ' 02 08 11 13   nº:  6
    If miGrilla01EnEstudio = 2 And _
       miGrilla02EnEstudio = 8 And _
       miGrilla03EnEstudio = 11 And _
       miGrilla04EnEstudio = 13 Then
           miModeloEnEstudio = 6
    End If
    
    ' 03 05 10 16   nº:  7
    If miGrilla01EnEstudio = 3 And _
       miGrilla02EnEstudio = 5 And _
       miGrilla03EnEstudio = 10 And _
       miGrilla04EnEstudio = 16 Then
           miModeloEnEstudio = 7
    End If
    
    ' 03 05 12 14   nº:  8
    If miGrilla01EnEstudio = 3 And _
       miGrilla02EnEstudio = 5 And _
       miGrilla03EnEstudio = 12 And _
       miGrilla04EnEstudio = 14 Then
           miModeloEnEstudio = 8
    End If
    
    ' 03 06 11 14   nº:  9
    If miGrilla01EnEstudio = 3 And _
       miGrilla02EnEstudio = 6 And _
       miGrilla03EnEstudio = 11 And _
       miGrilla04EnEstudio = 14 Then
           miModeloEnEstudio = 9
    End If
    
    ' 04 05 12 13   nº:  10
    If miGrilla01EnEstudio = 4 And _
       miGrilla02EnEstudio = 5 And _
       miGrilla03EnEstudio = 12 And _
       miGrilla04EnEstudio = 13 Then
           miModeloEnEstudio = 10
    End If
    
    ' 04 06 09 15   nº:  11
    If miGrilla01EnEstudio = 4 And _
       miGrilla02EnEstudio = 6 And _
       miGrilla03EnEstudio = 9 And _
       miGrilla04EnEstudio = 15 Then
           miModeloEnEstudio = 11
    End If
    
    ' 04 06 11 13   nº:  12
    If miGrilla01EnEstudio = 4 And _
       miGrilla02EnEstudio = 6 And _
       miGrilla03EnEstudio = 11 And _
       miGrilla04EnEstudio = 13 Then
           miModeloEnEstudio = 12
    End If
End Sub

' DETERMINA EL NUMERO DE LA GRILLA 01 QUE SE ESTA UTILIZANDO
Private Sub miDeterminaGrilla01()
    ' 01 07 10 16   nº:  1
    If miVectorSolucion(1) = miVectorSolucion(7) And _
       miVectorSolucion(1) = miVectorSolucion(10) And _
       miVectorSolucion(1) = miVectorSolucion(16) Then
           miGrilla01EnEstudio = 1
    End If
    
    ' 01 07 12 14   nº:  2
    If miVectorSolucion(1) = miVectorSolucion(7) And _
       miVectorSolucion(1) = miVectorSolucion(12) And _
       miVectorSolucion(1) = miVectorSolucion(14) Then
           miGrilla01EnEstudio = 2
    End If
    
    ' 01 08 10 15   nº:  3
    If miVectorSolucion(1) = miVectorSolucion(8) And _
       miVectorSolucion(1) = miVectorSolucion(10) And _
       miVectorSolucion(1) = miVectorSolucion(15) Then
           miGrilla01EnEstudio = 3
    End If
    
    ' 01 08 11 14   nº:  4
    If miVectorSolucion(1) = miVectorSolucion(8) And _
       miVectorSolucion(1) = miVectorSolucion(11) And _
       miVectorSolucion(1) = miVectorSolucion(14) Then
           miGrilla01EnEstudio = 4
    End If
End Sub

' DETERMINA EL NUMERO DE LA GRILLA 02 QUE SE ESTA UTILIZANDO
Private Sub miDeterminaGrilla02()
    ' 02 07 09 16   nº:  5
    If miVectorSolucion(2) = miVectorSolucion(7) And _
       miVectorSolucion(2) = miVectorSolucion(9) And _
       miVectorSolucion(2) = miVectorSolucion(16) Then
           miGrilla02EnEstudio = 5
    End If
    
    ' 02 07 12 13   nº:  6
    If miVectorSolucion(2) = miVectorSolucion(7) And _
       miVectorSolucion(2) = miVectorSolucion(12) And _
       miVectorSolucion(2) = miVectorSolucion(13) Then
           miGrilla02EnEstudio = 6
    End If
    
    ' 02 08 09 15   nº:  7
    If miVectorSolucion(2) = miVectorSolucion(8) And _
       miVectorSolucion(2) = miVectorSolucion(9) And _
       miVectorSolucion(2) = miVectorSolucion(15) Then
           miGrilla02EnEstudio = 7
    End If
    
    ' 02 08 11 13   nº:  8
    If miVectorSolucion(2) = miVectorSolucion(8) And _
       miVectorSolucion(2) = miVectorSolucion(11) And _
       miVectorSolucion(2) = miVectorSolucion(13) Then
           miGrilla02EnEstudio = 8
    End If
End Sub

' DETERMINA EL NUMERO DE LA GRILLA 03 QUE SE ESTA UTILIZANDO
Private Sub miDeterminaGrilla03()
    ' 03 05 10 16   nº:  9
    If miVectorSolucion(3) = miVectorSolucion(5) And _
       miVectorSolucion(3) = miVectorSolucion(10) And _
       miVectorSolucion(3) = miVectorSolucion(16) Then
           miGrilla03EnEstudio = 9
    End If
    
    ' 03 05 12 14   nº:  10
    If miVectorSolucion(3) = miVectorSolucion(5) And _
       miVectorSolucion(3) = miVectorSolucion(12) And _
       miVectorSolucion(3) = miVectorSolucion(14) Then
           miGrilla03EnEstudio = 10
    End If
    
    ' 03 06 09 16   nº:  11
    If miVectorSolucion(3) = miVectorSolucion(6) And _
       miVectorSolucion(3) = miVectorSolucion(9) And _
       miVectorSolucion(3) = miVectorSolucion(16) Then
           miGrilla03EnEstudio = 11
    End If
    
    ' 03 06 12 13   nº:  12
    If miVectorSolucion(3) = miVectorSolucion(6) And _
       miVectorSolucion(3) = miVectorSolucion(12) And _
       miVectorSolucion(3) = miVectorSolucion(13) Then
           miGrilla03EnEstudio = 12
    End If
End Sub

' DETERMINA EL NUMERO DE LA GRILLA 04 QUE SE ESTA UTILIZANDO
Private Sub miDeterminaGrilla04()
    ' 04 05 10 15   nº:  13
    If miVectorSolucion(4) = miVectorSolucion(5) And _
       miVectorSolucion(4) = miVectorSolucion(10) And _
       miVectorSolucion(4) = miVectorSolucion(15) Then
           miGrilla04EnEstudio = 13
    End If
    
    ' 04 05 11 14   nº:  14
    If miVectorSolucion(4) = miVectorSolucion(5) And _
       miVectorSolucion(4) = miVectorSolucion(11) And _
       miVectorSolucion(4) = miVectorSolucion(14) Then
           miGrilla04EnEstudio = 14
    End If
    
    ' 04 06 09 15   nº:  15
    If miVectorSolucion(4) = miVectorSolucion(6) And _
       miVectorSolucion(4) = miVectorSolucion(9) And _
       miVectorSolucion(4) = miVectorSolucion(15) Then
           miGrilla04EnEstudio = 15
    End If
    
    ' 04 06 11 13   nº:  16
    If miVectorSolucion(4) = miVectorSolucion(6) And _
       miVectorSolucion(4) = miVectorSolucion(11) And _
       miVectorSolucion(4) = miVectorSolucion(13) Then
           miGrilla04EnEstudio = 16
    End If
End Sub

Private Sub cmdCargaCeldas_Click()
    Dim i As Integer
    For i = 1 To 16
        txtProblema(i) = i
    Next i
End Sub

Private Sub cmdGiro90_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Giro90(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdGiro180_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Giro180(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdGiro270_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Giro270(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdFilas12_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdFilas34_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas34(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdFilas1234_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Filas1234(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdColumnas12_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas12(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdColumnas34_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas34(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdColumnas1234_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Columnas1234(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdNiveles_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Niveles(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdTorres_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Torres(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdNivelesTorres_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = NivelesTorres(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdHorizontal_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Horizontal(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdVertical_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = Vertical(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdTransponerIzquierda_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = TransponerIzquierda(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdTransponerDerecha_Click()
    Dim miDato As String
    Dim miResultado As String
    miDato = ExtraerProblema()
    miResultado = TransponerDerecha(miDato)
    MuestraTrasnformacion (miResultado)
    Call cmdAnalizarTransformada_Click
End Sub

Private Sub cmdCanonica_Click()
    Dim miCanonica(1 To 12) As String
    Dim miCadenaAnalisis As String
    Dim miModelo As Integer
    Dim miResultadoCanonico(1 To 16) As String
    Dim miCadenaResultado As String
    Dim i As Integer
    
    miCanonica(1) = "1234341221434321"
    miCanonica(2) = "1234431221433421"
    miCanonica(3) = "1234341241232341"
    miCanonica(4) = "1234341241232341"
    miCanonica(5) = "1234341243212143"
    miCanonica(6) = "1234431234212143"
    miCanonica(7) = "1234342121434312"
    miCanonica(8) = "1234432121433412"
    miCanonica(9) = "1234432131422413"
    miCanonica(10) = "1234432124133142"
    miCanonica(11) = "1234342143122143"
    miCanonica(12) = "1234432134122143"
    
    ' CARGA LOS DATOS PARA ANALIZARLOS
    miCadenaAnalisis = ""
    For i = 1 To 16
        miVectorAnalisis(i) = Val(txtProblema(i).Text)
        miCadenaAnalisis = miCadenaAnalisis + Trim(txtProblema(i).Text)
    Next i
    If miCadenaAnalisis <> "" Then
        Call cmdAnalizarOriginal_Click
        miModelo = Val(Trim(txtModOriginal.Text))
        
        miCadenaResultado = ""
        For i = 1 To 16
            If miVectorAnalisis(i) = 0 Then
                miResultadoCanonico(i) = 0
            Else
                miResultadoCanonico(i) = Mid(miCanonica(miModelo), i, 1)
            End If
            miCadenaResultado = miCadenaResultado + miResultadoCanonico(i)
        Next i
        
    End If
    
    MuestraTrasnformacion (miCadenaResultado)
    Call cmdAnalizarTransformada_Click
End Sub


Private Sub ImprimeArchivo(miTitulo As String)
    ' IMPRIMIR LOS RESULTADOS A UN ARCHIVO DE TEXTO
    '***************************************************************
    mi_PrintLine = ""
    
    If Len(Trim(Str(miSolucionEnEstudio))) = 1 Then
        mi_PrintLine = mi_PrintLine + "Solución: 00" + Trim(Str(miSolucionEnEstudio))
    End If
    If Len(Trim(Str(miSolucionEnEstudio))) = 2 Then
        mi_PrintLine = mi_PrintLine + "Solución: 0" + Trim(Str(miSolucionEnEstudio))
    End If
    If Len(Trim(Str(miSolucionEnEstudio))) = 3 Then
        mi_PrintLine = mi_PrintLine + "Solución: " + Trim(Str(miSolucionEnEstudio))
    End If
    
    mi_PrintLine = mi_PrintLine + "  --  Modelo :"
    If Len(Trim(Str(miModeloEnEstudio))) = 1 Then
        mi_PrintLine = mi_PrintLine + " 0" + Trim(Str(miModeloEnEstudio))
    Else
        mi_PrintLine = mi_PrintLine + " " + Trim(Str(miModeloEnEstudio))
    End If
    
    mi_PrintLine = mi_PrintLine + "  --  Grillas: "
    If Len(Trim(Str(miGrilla01EnEstudio))) = 1 Then
        mi_PrintLine = mi_PrintLine + " 0" + Trim(Str(miGrilla01EnEstudio))
    Else
        mi_PrintLine = mi_PrintLine + " " + Trim(Str(miGrilla01EnEstudio))
    End If
    
    mi_PrintLine = mi_PrintLine + " -"
    If Len(Trim(Str(miGrilla02EnEstudio))) = 1 Then
        mi_PrintLine = mi_PrintLine + " 0" + Trim(Str(miGrilla02EnEstudio))
    Else
        mi_PrintLine = mi_PrintLine + " " + Trim(Str(miGrilla02EnEstudio))
    End If
    
    mi_PrintLine = mi_PrintLine + " -"
    If Len(Trim(Str(miGrilla03EnEstudio))) = 1 Then
        mi_PrintLine = mi_PrintLine + " 0" + Trim(Str(miGrilla03EnEstudio))
    Else
        mi_PrintLine = mi_PrintLine + " " + Trim(Str(miGrilla03EnEstudio))
    End If
    
    mi_PrintLine = mi_PrintLine + " -"
    If Len(Trim(Str(miGrilla04EnEstudio))) = 1 Then
        mi_PrintLine = mi_PrintLine + " 0" + Trim(Str(miGrilla04EnEstudio))
    Else
        mi_PrintLine = mi_PrintLine + " " + Trim(Str(miGrilla04EnEstudio))
    End If
    
    mi_PrintLine = mi_PrintLine + miTitulo
    
    '****************************************************************************************
    ' OJO IMPRESION CONDICIONAL
    '****************************************************************************************
    'If miModeloEnEstudio = 1 Then
        Print #1, mi_PrintLine
    'End If
    '****************************************************************************************
End Sub

' MUESTRA LA CADENA DE CARACTERES EN EL DISPLAY DE TRANSFORMACION
Public Sub MuestraTrasnformacion(miEnviado As String)
    Dim i As Integer
    For i = 1 To 16
        If Mid(miEnviado, i, 1) = "0" Then
            txtTrans(i).Text = ""
        Else
            txtTrans(i).Text = Mid(miEnviado, i, 1)
        End If
    Next i
End Sub


