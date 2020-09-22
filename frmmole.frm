VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmole 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Mole Converter"
   ClientHeight    =   5130
   ClientLeft      =   3375
   ClientTop       =   2955
   ClientWidth     =   7800
   Icon            =   "frmmole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmole.frx":0E42
   ScaleHeight     =   5130
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Tag             =   "80"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   4080
      Top             =   2520
   End
   Begin VB.TextBox txtg 
      Height          =   285
      Left            =   5880
      TabIndex        =   95
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtelm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H008080FF&
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   94
      Top             =   480
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "_"
      Height          =   255
      Left            =   1680
      TabIndex        =   93
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox subs 
      Height          =   285
      Left            =   1200
      TabIndex        =   91
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton E1 
      BackColor       =   &H008080FF&
      Caption         =   "H"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton E3 
      BackColor       =   &H000000C0&
      Caption         =   "Li"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E4 
      BackColor       =   &H000000C0&
      Caption         =   "Be"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E11 
      BackColor       =   &H000000C0&
      Caption         =   "Na"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E37 
      BackColor       =   &H000000C0&
      Caption         =   "Rb"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E19 
      BackColor       =   &H000000C0&
      Caption         =   "K"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E56 
      BackColor       =   &H000000C0&
      Caption         =   "Cs"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E87 
      BackColor       =   &H000000C0&
      Caption         =   "Fr"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E12 
      BackColor       =   &H000000C0&
      Caption         =   "Mg"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E88 
      BackColor       =   &H000000C0&
      Caption         =   "Ra"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E57 
      BackColor       =   &H000000C0&
      Caption         =   "Ba"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E20 
      BackColor       =   &H000000C0&
      Caption         =   "Ca"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E38 
      BackColor       =   &H000000C0&
      Caption         =   "Sr"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E21 
      BackColor       =   &H00000080&
      Caption         =   "Sc"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E23 
      BackColor       =   &H00000080&
      Caption         =   "V"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E30 
      BackColor       =   &H00000080&
      Caption         =   "Zn"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E27 
      BackColor       =   &H00000080&
      Caption         =   "Co"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E24 
      BackColor       =   &H00000080&
      Caption         =   "Cr"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E22 
      BackColor       =   &H00000080&
      Caption         =   "Ti"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E29 
      BackColor       =   &H00000080&
      Caption         =   "Cu"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E26 
      BackColor       =   &H00000080&
      Caption         =   "Fe"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E28 
      BackColor       =   &H00000080&
      Caption         =   "Ni"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E25 
      BackColor       =   &H00000080&
      Caption         =   "Mn"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Al"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E14 
      BackColor       =   &H008080FF&
      Caption         =   "Si"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E17 
      BackColor       =   &H008080FF&
      Caption         =   "Cl"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E18 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ar"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E16 
      BackColor       =   &H008080FF&
      Caption         =   "S"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E15 
      BackColor       =   &H008080FF&
      Caption         =   "P"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton E9 
      BackColor       =   &H008080FF&
      Caption         =   "F"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E8 
      BackColor       =   &H008080FF&
      Caption         =   "O"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E7 
      BackColor       =   &H008080FF&
      Caption         =   "N"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E6 
      BackColor       =   &H008080FF&
      Caption         =   "C"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ne"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E5 
      BackColor       =   &H008080FF&
      Caption         =   "B"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton E2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "He"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton E31 
      BackColor       =   &H00000080&
      Caption         =   "Ga"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E36 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Kr"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E35 
      BackColor       =   &H008080FF&
      Caption         =   "Br"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E33 
      BackColor       =   &H008080FF&
      Caption         =   "As"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E32 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ge"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E34 
      BackColor       =   &H008080FF&
      Caption         =   "Se"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton E40 
      BackColor       =   &H00000080&
      Caption         =   "Zr"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E41 
      BackColor       =   &H00000080&
      Caption         =   "Nb"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E44 
      BackColor       =   &H00000080&
      Caption         =   "Ru"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E45 
      BackColor       =   &H00000080&
      Caption         =   "Rh"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E48 
      BackColor       =   &H00000080&
      Caption         =   "Cd"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E42 
      BackColor       =   &H00000080&
      Caption         =   "Mo"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E43 
      BackColor       =   &H00000080&
      Caption         =   "Tc"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E46 
      BackColor       =   &H00000080&
      Caption         =   "Pd"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E47 
      BackColor       =   &H00000080&
      Caption         =   "Ag"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E53 
      BackColor       =   &H008080FF&
      Caption         =   "I"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E54 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Xe"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E52 
      BackColor       =   &H008080FF&
      Caption         =   "Te"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E49 
      BackColor       =   &H00000080&
      Caption         =   "In"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E50 
      BackColor       =   &H00000080&
      Caption         =   "Sn"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E39 
      BackColor       =   &H00000080&
      Caption         =   "Y"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E51 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sb"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton E73 
      BackColor       =   &H00000080&
      Caption         =   "Ta"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E76 
      BackColor       =   &H00000080&
      Caption         =   "Os"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E71 
      BackColor       =   &H00000080&
      Caption         =   "Lu"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E75 
      BackColor       =   &H00000080&
      Caption         =   "Re"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E80 
      BackColor       =   &H00000080&
      Caption         =   "Hg"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E77 
      BackColor       =   &H00000080&
      Caption         =   "Ir"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E78 
      BackColor       =   &H00000080&
      Caption         =   "Pt"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E79 
      BackColor       =   &H00000080&
      Caption         =   "Au"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E85 
      BackColor       =   &H008080FF&
      Caption         =   "At"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E81 
      BackColor       =   &H00000080&
      Caption         =   "Ti"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E72 
      BackColor       =   &H00000080&
      Caption         =   "Hf"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E74 
      BackColor       =   &H00000080&
      Caption         =   "W"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E82 
      BackColor       =   &H00000080&
      Caption         =   "Pb"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E84 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Po"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E83 
      BackColor       =   &H00000080&
      Caption         =   "Bi"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E86 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rn"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton E103 
      BackColor       =   &H00000080&
      Caption         =   "Lr"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E104 
      BackColor       =   &H00000080&
      Caption         =   "Rf"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E105 
      BackColor       =   &H00000080&
      Caption         =   "Db"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E106 
      BackColor       =   &H00000080&
      Caption         =   "Sg"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E107 
      BackColor       =   &H00000080&
      Caption         =   "Bh"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E108 
      BackColor       =   &H00000080&
      Caption         =   "Hs"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E109 
      BackColor       =   &H00000080&
      Caption         =   "Mt"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E110 
      BackColor       =   &H00000080&
      Caption         =   "Ds"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton E111 
      BackColor       =   &H00000080&
      Caption         =   "Rg"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton CmdOP 
      Caption         =   "("
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton CmdCP 
      Caption         =   ")"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CmdSolve 
      Caption         =   "&Solve"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox TxtExp 
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   4080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   529
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmmole.frx":86356
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2520
      TabIndex        =   98
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5040
      TabIndex        =   97
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mass In Grams:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   96
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SubScript:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   92
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label LblExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   7080
      TabIndex        =   6
      Top             =   0
      Width           =   285
   End
   Begin VB.Label LblMin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   0
      Width           =   210
   End
   Begin VB.Label LblResult 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Top             =   960
      Width           =   7575
   End
End
Attribute VB_Name = "frmmole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FontFile As String
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_MINIMIZE = 6
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim StartMoving As Boolean
Dim InitialX As Long, InitialY As Long
Dim Expressions() As String, nExp As Integer, nTemp As Integer
Private Sub CmdCP_Click()
TxtExp.SelStart = TxtExp.SelStart - 1
txtelm.Text = txtelm.Text + ")"
AddExp ")"
End Sub

Private Sub CmdOP_Click()
txtelm.Text = txtelm.Text + "("
AddExp "("
End Sub
Private Sub CmdSolve_Click()
If txtg.Text = "" Then
MsgBox "Weight Needed", vbExclamation, "Error:"
Else
AddExp "0)"
TxtExp.SelStart = 0
AddExp txtg.Text & "/"
TxtExp.SelStart = 1000
Dim Exp As String
Dim Result As Double

CmdSolve.Enabled = False


Exp = Trim(TxtExp.Text)
Exp = Replace(Exp, " ", "")
TxtExp = Exp
LblResult = "0"
TxtExp.SelStart = 0
TxtExp.SelLength = Len(Exp)
TxtExp.SelColor = &HFF8080
TxtExp.SelBold = False
'
If EvaluateExp(Exp, Result, 0, TxtExp, LblResult) Then
    'Result
    LblResult = Result
    If Exp <> Expressions(nExp) Then
        nExp = nExp + 1
        ReDim Preserve Expressions(nExp)
        Expressions(nExp) = TxtExp.Text
        nTemp = nExp
    End If
End If

CmdSolve.Enabled = True
TxtExp.SetFocus
CmdSolve.Enabled = False
LblResult.Caption = LblResult.Caption + "  mol"
End If
End Sub

Private Sub CmdClear_Click()
LblResult.Caption = "0"
txtelm = ""
TxtExp.Text = "(0+"
TxtExp.SelStart = 3
subs.Text = ""
txtg.Text = ""
CmdSolve.Enabled = True
End Sub

Private Sub CmdSub_Click()
AddExp "-"
End Sub

Private Sub Command1_Click()
If Val(subs.Text) <> 0 Then
TxtExp.SelStart = TxtExp.SelStart - 1
AddExp "*" & subs.Text
Call script
Else
MsgBox "Needed Subscript"
End If
End Sub


Private Sub E1_Click()
txtelm.Text = txtelm.Text + "H"
AddExp "1.00794+"
End Sub

Private Sub E10_Click()
txtelm.Text = txtelm.Text + "Ne"
AddExp "20.1797+"
End Sub

Private Sub E103_Click()
txtelm.Text = txtelm.Text + "Lr"
AddExp "262+"
End Sub

Private Sub E104_Click()
txtelm.Text = txtelm.Text + "Rf"
AddExp "261+"
End Sub

Private Sub E105_Click()
txtelm.Text = txtelm.Text + "Db"
AddExp "262+"
End Sub

Private Sub E106_Click()
txtelm.Text = txtelm.Text + "Sg"
AddExp "263+"
End Sub

Private Sub E107_Click()
txtelm.Text = txtelm.Text + "Bh"
AddExp "262+"
End Sub

Private Sub E108_Click()
txtelm.Text = txtelm.Text + "Hs"
AddExp "265+"
End Sub

Private Sub E109_Click()
txtelm.Text = txtelm.Text + "Mt"
AddExp "266+"
End Sub

Private Sub E11_Click()
txtelm.Text = txtelm.Text + "Na"
AddExp "22.989768+"
End Sub

Private Sub E110_Click()
txtelm.Text = txtelm.Text + "Ds"
AddExp "271+"
End Sub

Private Sub E111_Click()
txtelm.Text = txtelm.Text + "Rg"
AddExp "272+"
End Sub

Private Sub E12_Click()
txtelm.Text = txtelm.Text + "Mg"
AddExp "24.3050+"
End Sub

Private Sub E13_Click()
txtelm.Text = txtelm.Text + "Al"
AddExp "26.981539+"
End Sub

Private Sub E14_Click()
txtelm.Text = txtelm.Text + "Si"
AddExp "28.0855+"
End Sub

Private Sub E15_Click()
txtelm.Text = txtelm.Text + "P"
AddExp "30.973762+"
End Sub

Private Sub E16_Click()
txtelm.Text = txtelm.Text + "S"
AddExp "32.066+"
End Sub

Private Sub E17_Click()
txtelm.Text = txtelm.Text + "Cl"
AddExp "35.4527+"
End Sub

Private Sub E18_Click()
txtelm.Text = txtelm.Text + "Ar"
AddExp "39.948+"
End Sub

Private Sub E19_Click()
txtelm.Text = txtelm.Text + "K"
AddExp "39.0983+"
End Sub

Private Sub E2_Click()
txtelm.Text = txtelm.Text + "He"
AddExp "4.002602+"
End Sub

Private Sub E20_Click()
txtelm.Text = txtelm.Text + "Ca"
AddExp "40.078+"
End Sub

Private Sub E21_Click()
txtelm.Text = txtelm.Text + "Sc"
AddExp "44.955910+"
End Sub

Private Sub E22_Click()
txtelm.Text = txtelm.Text + "Ti"
AddExp "47.867+"
End Sub

Private Sub E23_Click()
txtelm.Text = txtelm.Text + "V"
AddExp "50.9415+"
End Sub

Private Sub E24_Click()
txtelm.Text = txtelm.Text + "Cr"
AddExp "51.9961+"
End Sub

Private Sub E25_Click()
txtelm.Text = txtelm.Text + "Mn"
AddExp "54.93805+"
End Sub

Private Sub E26_Click()
txtelm.Text = txtelm.Text + "Fe"
AddExp "55.845+"
End Sub

Private Sub E27_Click()
txtelm.Text = txtelm.Text + "Co"
AddExp "58.93320+"
End Sub

Private Sub E28_Click()
txtelm.Text = txtelm.Text + "Ni"
AddExp "58.6934+"
End Sub

Private Sub E29_Click()
txtelm.Text = txtelm.Text + "Cu"
AddExp "63.546+"
End Sub

Private Sub E3_Click()
txtelm.Text = txtelm.Text + "Li"
AddExp "6.941+"
End Sub

Private Sub E30_Click()
txtelm.Text = txtelm.Text + "Zn"
AddExp "65.39+"
End Sub

Private Sub E31_Click()
txtelm.Text = txtelm.Text + "Ga"
AddExp "69.723+"
End Sub

Private Sub E32_Click()
txtelm.Text = txtelm.Text + "Ge"
AddExp "72.61+"
End Sub

Private Sub E33_Click()
txtelm.Text = txtelm.Text + "As"
AddExp "74.92159+"
End Sub

Private Sub E34_Click()
txtelm.Text = txtelm.Text + "Se"
AddExp "78.96+"
End Sub

Private Sub E35_Click()
txtelm.Text = txtelm.Text + "Br"
AddExp "79.904+"
End Sub

Private Sub E36_Click()
txtelm.Text = txtelm.Text + "Kr"
AddExp "83.80+"
End Sub

Private Sub E37_Click()
txtelm.Text = txtelm.Text + "Rb"
AddExp "85.4678+"
End Sub

Private Sub E38_Click()
txtelm.Text = txtelm.Text + "Sr"
AddExp "87.62+"
End Sub

Private Sub E39_Click()
txtelm.Text = txtelm.Text + "Y"
AddExp "88.90585+"
End Sub

Private Sub E4_Click()
txtelm.Text = txtelm.Text + "Be"
AddExp "9.012182+"
End Sub

Private Sub E40_Click()
txtelm.Text = txtelm.Text + "Zr"
AddExp "91.224+"
End Sub

Private Sub E41_Click()
txtelm.Text = txtelm.Text + "Nb"
AddExp "92.90638+"
End Sub

Private Sub E42_Click()
txtelm.Text = txtelm.Text + "Mo"
AddExp "95.94+"
End Sub

Private Sub E43_Click()
txtelm.Text = txtelm.Text + "Tc"
AddExp "98+"
End Sub

Private Sub E44_Click()
txtelm.Text = txtelm.Text + "Ru"
AddExp "101.07+"
End Sub

Private Sub E45_Click()
txtelm.Text = txtelm.Text + "Rh"
AddExp "102.90550+"
End Sub

Private Sub E46_Click()
txtelm.Text = txtelm.Text + "Pd"
AddExp "106.42+"
End Sub

Private Sub E47_Click()
txtelm.Text = txtelm.Text + "Ag"
AddExp "107.8682+"
End Sub

Private Sub E48_Click()
txtelm.Text = txtelm.Text + "Cd"
AddExp "112.411+"
End Sub

Private Sub E49_Click()
txtelm.Text = txtelm.Text + "In"
AddExp "114.818"
End Sub

Private Sub E5_Click()
txtelm.Text = txtelm.Text + "B"
AddExp "10.81+"
End Sub

Private Sub E50_Click()
txtelm.Text = txtelm.Text + "Sn"
AddExp "118.710+"
End Sub

Private Sub E51_Click()
txtelm.Text = txtelm.Text + "Sb"
AddExp "121.760+"
End Sub

Private Sub E52_Click()
txtelm.Text = txtelm.Text + "Te"
AddExp "127.60+"
End Sub

Private Sub E53_Click()
txtelm.Text = txtelm.Text + "I"
AddExp "126.90447+"
End Sub

Private Sub E54_Click()
txtelm.Text = txtelm.Text + "Xe"
AddExp "131.29+"
End Sub

Private Sub E56_Click()
txtelm.Text = txtelm.Text + "Cs"
AddExp "132.90543+"
End Sub

Private Sub E57_Click()
txtelm.Text = txtelm.Text + "Ba"
AddExp "137.327+"
End Sub

Private Sub E6_Click()
txtelm.Text = txtelm.Text + "C"
AddExp "12.011+"
End Sub

Private Sub E7_Click()
txtelm.Text = txtelm.Text + "N"
AddExp "14.00674+"
End Sub

Private Sub E71_Click()
txtelm.Text = txtelm.Text + "Lu"
AddExp "174.967+"
End Sub

Private Sub E72_Click()
txtelm.Text = txtelm.Text + "Hf"
AddExp "178.49+"
End Sub

Private Sub E73_Click()
txtelm.Text = txtelm.Text + "Ta"
AddExp "180.9479+"
End Sub

Private Sub E74_Click()
txtelm.Text = txtelm.Text + "W"
AddExp "183.84+"
End Sub

Private Sub E75_Click()
txtelm.Text = txtelm.Text + "Re"
AddExp "186.207+"
End Sub

Private Sub E76_Click()
txtelm.Text = txtelm.Text + "Os"
AddExp "190.23+"
End Sub

Private Sub E77_Click()
txtelm.Text = txtelm.Text + "Ir"
AddExp "192.217+"
End Sub

Private Sub E78_Click()
txtelm.Text = txtelm.Text + "Pt"
AddExp "195.08+"
End Sub

Private Sub E79_Click()
txtelm.Text = txtelm.Text + "Au"
AddExp "196.96654+"
End Sub

Private Sub E8_Click()
txtelm.Text = txtelm.Text + "O"
AddExp "15.9994+"
End Sub

Private Sub E80_Click()
txtelm.Text = txtelm.Text + "Hg"
AddExp "200.59+"
End Sub

Private Sub E81_Click()
txtelm.Text = txtelm.Text + "Tl"
AddExp "204.3833+"
End Sub

Private Sub E82_Click()
txtelm.Text = txtelm.Text + "Pb"
AddExp "207.2+"
End Sub

Private Sub E83_Click()
txtelm.Text = txtelm.Text + "Bi"
AddExp "208.98037+"
End Sub

Private Sub E84_Click()
txtelm.Text = txtelm.Text + "Po"
AddExp "209+"
End Sub

Private Sub E85_Click()
txtelm.Text = txtelm.Text + "At"
AddExp "210+"
End Sub

Private Sub E86_Click()
txtelm.Text = txtelm.Text + "Rn"
AddExp "222+"
End Sub

Private Sub E87_Click()
txtelm.Text = txtelm.Text + "Fr"
AddExp "223+"
End Sub

Private Sub E88_Click()
txtelm.Text = txtelm.Text + "Ra"
AddExp "226+"
End Sub

Private Sub E9_Click()
txtelm.Text = txtelm.Text + "F"
AddExp "18.9984032+"
End Sub

Private Sub Form_Load()

On Error Resume Next
 
 FontFile = App.Path & "\Bede.ttf"
 Call AddFontResource(FontFile)
txtelm.FontName = "Bede"
txtelm.FontSize = "12"
nExp = 0
ReDim Preserve Expressions(nExp)
Expressions(0) = ""

TxtExp.SelColor = &HFF8080

TxtExp = "(0+"
TxtExp.SelStart = 3
TxtExp.SelLength = Len(TxtExp)
TxtExp.SelColor = &HFF8080
TxtExp.SelLength = 0

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
InitialX = X
InitialY = Y
StartMoving = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If StartMoving Then
    Me.Left = Me.Left + (X - InitialX)
    Me.Top = Me.Top + (Y - InitialY)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartMoving = False
End Sub

Private Sub Label3_Click()
frmAbout.Visible = True
End Sub

Private Sub Label4_DblClick()
Dim strfile As String
 strfile = App.Path & "\source.txt"
 Dim LF As Variant
    LF = Shell("Notepad " & strfile, 3)
End Sub

Private Sub LblExit_Click()
PlayFromRes "101"
Timer1.Enabled = True
End Sub

Private Sub LblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblExit.BorderStyle = 1
End Sub

Private Sub LblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblExit.BorderStyle = 0
End Sub
Private Sub LblMin_Click()
ShowWindow Me.hwnd, SW_MINIMIZE
End Sub
Private Sub LblMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblMin.BorderStyle = 1
End Sub
Private Sub LblMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblMin.BorderStyle = 0
End Sub

Private Sub LblResult_Click()

If IsNumeric(LblResult.Caption) Then
    Clipboard.Clear
    Clipboard.SetText LblResult.Caption
End If
End Sub

Private Sub subs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Command1.Value = True
End If
End Sub

Private Sub AddExp(ByVal Exp As String)
Dim sStart As Integer

'input position
sStart = TxtExp.SelStart
TxtExp.Text = Mid(TxtExp.Text, 1, sStart) + Exp + Mid(TxtExp.Text, sStart + 1 + TxtExp.SelLength)
TxtExp.SelStart = 0
TxtExp.SelLength = Len(TxtExp)
TxtExp.SelColor = &HFF8080
TxtExp.SelBold = False
TxtExp.SelLength = 0
TxtExp.SelStart = sStart + 100
TxtExp.SetFocus
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

Private Sub txtg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
CmdSolve.Value = True
End If
End Sub
