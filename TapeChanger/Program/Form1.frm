VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changer management utility V1.0"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Caption         =   "Programs launched after backup"
      Height          =   1155
      Left            =   4200
      TabIndex        =   63
      Top             =   270
      Width           =   4035
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "3."
         Height          =   225
         Left            =   120
         TabIndex        =   69
         Top             =   870
         Width           =   135
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "2."
         Height          =   195
         Left            =   90
         TabIndex        =   68
         Top             =   570
         Width           =   165
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "1."
         Height          =   225
         Left            =   120
         TabIndex        =   67
         Top             =   270
         Width           =   135
      End
      Begin VB.Label Shutdown3 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   300
         TabIndex        =   66
         ToolTipText     =   "Third program run after tape backup is finished"
         Top             =   840
         Width           =   3645
      End
      Begin VB.Label Shutdown2 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   300
         TabIndex        =   65
         ToolTipText     =   "Second program run after tape backup is finished"
         Top             =   540
         Width           =   3645
      End
      Begin VB.Label Shutdown1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   300
         TabIndex        =   64
         ToolTipText     =   "First program run after tape backup is finished"
         Top             =   240
         Width           =   3645
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Programs launched after tape load"
      Height          =   1155
      Left            =   90
      TabIndex        =   56
      Top             =   270
      Width           =   4065
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "3."
         Height          =   195
         Left            =   60
         TabIndex        =   62
         Top             =   870
         Width           =   165
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "2."
         Height          =   195
         Left            =   60
         TabIndex        =   61
         Top             =   570
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "1."
         Height          =   195
         Left            =   60
         TabIndex        =   60
         Top             =   270
         Width           =   165
      End
      Begin VB.Label Program3 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   270
         TabIndex        =   59
         ToolTipText     =   "Program 3 to be run after tape is loaded"
         Top             =   840
         Width           =   3705
      End
      Begin VB.Label Program2 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   270
         TabIndex        =   58
         ToolTipText     =   "Program 2 to be run after tape is loaded"
         Top             =   540
         Width           =   3705
      End
      Begin VB.Label Program1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   270
         TabIndex        =   57
         ToolTipText     =   "Program 1 run after tape is loaded"
         Top             =   240
         Width           =   3705
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Settings"
      Height          =   1185
      Left            =   90
      TabIndex        =   49
      Top             =   1470
      Width           =   2145
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "Next tape loaded"
         Height          =   195
         Left            =   330
         TabIndex        =   55
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label NextLoaded 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   1650
         TabIndex        =   54
         ToolTipText     =   "Tape to be used by backup program"
         Top             =   810
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Last tape loaded"
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label LastLoaded 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   1650
         TabIndex        =   52
         ToolTipText     =   "Tape last used by backup program"
         Top             =   510
         Width           =   405
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cleaning tape in slot"
         Height          =   225
         Left            =   90
         TabIndex        =   51
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label CleaningTape 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   1650
         TabIndex        =   50
         ToolTipText     =   "Cleaning tape is loaded in this slot"
         Top             =   210
         Width           =   405
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Changer information"
      Height          =   2595
      Left            =   2280
      TabIndex        =   11
      Top             =   1470
      Width           =   5955
      Begin VB.CommandButton SaveSettings 
         Caption         =   "Save settings"
         Height          =   375
         Left            =   3780
         TabIndex        =   74
         ToolTipText     =   "Save adapter and device selection for use with auto mode"
         Top             =   1710
         Width           =   2085
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5400
         Top             =   540
      End
      Begin VB.CommandButton CancelAutoMode 
         Caption         =   "Cancel auto mode"
         Height          =   375
         Left            =   3780
         TabIndex        =   73
         ToolTipText     =   "Start or stop auto mode"
         Top             =   1290
         Width           =   2085
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "seconds"
         Height          =   225
         Left            =   4470
         TabIndex        =   72
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4440
         TabIndex        =   71
         ToolTipText     =   "Countdown timer"
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "Program will go automatic in"
         Height          =   255
         Left            =   3750
         TabIndex        =   70
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label Label41 
         Caption         =   "is in the drive"
         Height          =   225
         Left            =   4860
         TabIndex        =   48
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4320
         TabIndex        =   47
         ToolTipText     =   "This tape is loaded in the drive"
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "Tape"
         Height          =   195
         Left            =   3870
         TabIndex        =   46
         Top             =   2250
         Width           =   405
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "Tape changer ID:"
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Top             =   2220
         Width           =   2265
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   43
         Top             =   1920
         Width           =   285
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   42
         Top             =   1920
         Width           =   285
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   41
         Top             =   1920
         Width           =   285
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number of data transfer elements MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   40
         Top             =   1920
         Width           =   2595
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   39
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   38
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   37
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First data transfer element address MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   36
         Top             =   1680
         Width           =   2595
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   35
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   34
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   33
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number of import/export elements MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   1440
         Width           =   2595
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   31
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   30
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   29
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First import/export element addess MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   1200
         Width           =   2595
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   27
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   26
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   25
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number of storage elements MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   960
         Width           =   2595
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   23
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   22
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   21
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First storage element address MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   720
         Width           =   2595
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   19
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   18
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   17
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number of medium transport elements MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   480
         Width           =   2595
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   15
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   14
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2700
         TabIndex        =   13
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First medium transport element address MSB"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Adapter list"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   2730
      Width           =   2115
      Begin VB.ListBox Adapterlist 
         Height          =   1035
         Left            =   90
         TabIndex        =   10
         ToolTipText     =   "Lists compatible adapters"
         Top             =   210
         Width           =   1905
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Functions"
      Height          =   2655
      Left            =   5820
      TabIndex        =   4
      Top             =   4140
      Width           =   2415
      Begin VB.CommandButton Command4 
         Caption         =   "Reset changer"
         Height          =   435
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Re-initialise the changer (inventory the slots)"
         Top             =   1470
         Width           =   2145
      End
      Begin VB.CommandButton Quit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Exit this fine program"
         Top             =   2040
         Width           =   2145
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clean head"
         Enabled         =   0   'False
         Height          =   435
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Clean the tape drive's head"
         Top             =   900
         Width           =   2145
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Load or unload a tape"
         Top             =   330
         Width           =   2145
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Slot information"
      Height          =   2655
      Left            =   3240
      TabIndex        =   2
      Top             =   4140
      Width           =   2535
      Begin VB.ListBox TapesInSlots 
         Height          =   2205
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Lists physical tapes loaded in slots"
         Top             =   300
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Device listing"
      Height          =   2655
      Left            =   60
      TabIndex        =   0
      Top             =   4140
      Width           =   3105
      Begin VB.ListBox DeviceListing 
         Height          =   2205
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Lists devices connected to adapter selected"
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      Caption         =   "Written by MarkBorchers."
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   6390
      TabIndex        =   75
      Top             =   30
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adapterlist_Click()
    
    DeviceListing.Clear
        
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &H24
    ExecIO.SRB_BufPointer = VarPtr(Databuffer2)
    ExecIO.SRB_CDBLen = &H6
    ExecIO.SRB_CDBByte(0) = &H12
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = &H0
    ExecIO.SRB_CDBByte(3) = &H0
    ExecIO.SRB_CDBByte(4) = &H24
    ExecIO.SRB_CDBByte(5) = &H0
    
    For i = 0 To 15
    ExecIO.SRB_Target = i
    nRet = SendASPI32ExecIOEx(ExecIO)
    While ExecIO.SRB_Status = SS_PENDING
    DoEvents
    Wend
    
    a$ = ""
    
    For o = 0 To 94
        If Val(Databuffer2.Info(o)) < 32 Or Val(Databuffer2.Info(o)) > 127 Then
        GoTo dd
        End If
        a$ = a$ + Chr(Databuffer2.Info(o))
dd:
    Next o
    
    DeviceListing.AddItem i & " " & a$
    
    Erase Databuffer2.Info()

Next i
End Sub

Private Sub CancelAutoMode_Click()

Select Case CancelAutoMode.Caption

Case "Cancel auto mode"
    Timer1.Enabled = False
    Label47.Enabled = False
    Label48.Enabled = False
    Label49.Enabled = False
    Label48.Caption = "10"
    CancelAutoMode.Caption = "Start auto mode"
Case "Start auto mode"
    Timer1.Enabled = True
    Label47.Enabled = True
    Label48.Enabled = True
    Label49.Enabled = True
    CancelAutoMode.Caption = "Cancel auto mode"
End Select
End Sub

Private Sub Command1_Click()

If Left$(Command1.Caption, 4) = "Retu" Then

    'According to HP's changer SCSI spec, when a move element
    'command is issued where the source element is the drive,
    'the tape is automatically ejected from the drive. Then the
    'tape can be moved back to it's home slot. (You must record
    'which slot the tape came from, of course!)
    
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer1)
    ExecIO.SRB_CDBLen = &HC
    ExecIO.SRB_CDBByte(0) = &HA5
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = &H0
    ExecIO.SRB_CDBByte(3) = &H0
    ExecIO.SRB_CDBByte(4) = &H0
    ExecIO.SRB_CDBByte(5) = &H1
    ExecIO.SRB_CDBByte(6) = &H0
    ExecIO.SRB_CDBByte(7) = (TapeNowLoaded + Label16.Caption)
    ExecIO.SRB_CDBByte(8) = &H0
    ExecIO.SRB_CDBByte(9) = &H0
    ExecIO.SRB_CDBByte(10) = &H0
    ExecIO.SRB_CDBByte(11) = &H0
    
    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    
    TapesInSlots.List(TapesInSlots.ListIndex) = "Slot " & (TapesInSlots.ListIndex + 1) & " tape loaded"
    Label40.Caption = "-"
    
    Command1.Caption = "Load tape " & (TapesInSlots.ListIndex + 1)
        
    Exit Sub
    
    End If
    
    'Command1's caption is load tape...
    
    'Load tape from slot n in to the drive

    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer1)
    ExecIO.SRB_CDBLen = &HC
    ExecIO.SRB_CDBByte(0) = &HA5            'move medium command
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = Label6.Caption  'MSB picker
    ExecIO.SRB_CDBByte(3) = Label8.Caption  'LSB picker
    ExecIO.SRB_CDBByte(4) = Label14.Caption 'MSB tape slot
    ExecIO.SRB_CDBByte(5) = (TapesInSlots.ListIndex + Label16.Caption) 'LSB tape slot
    ExecIO.SRB_CDBByte(6) = Label30.Caption 'MSB tape drive
    ExecIO.SRB_CDBByte(7) = Label32.Caption 'LSB tape drive
    ExecIO.SRB_CDBByte(8) = &H0
    ExecIO.SRB_CDBByte(9) = &H0
    ExecIO.SRB_CDBByte(10) = &H0
    ExecIO.SRB_CDBByte(11) = &H0
    
    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend

    TapesInSlots.List(TapesInSlots.ListIndex) = "Slot " & (TapesInSlots.ListIndex + 1) & " tape in drive"
    Label40.Caption = TapesInSlots.ListIndex + 1
    
    TapeNowLoaded = TapesInSlots.ListIndex
    
    Command1.Caption = "Return tape to slot " & (TapesInSlots.ListIndex + 1)
    
    End Sub
    
Private Sub Command2_Click()

'Clean the tape drive's head. Note this uses the "CleanHeadTimeout"
'variable to determine when to return the cleaning tape to the
'slot. Otherwise, a nightmare programming requirement would be to
'poll the tape drive for when it ejects the tape. Maybe in the
'next release! Note that this button is almost identical to
'command1, but with the load/unload around the other way, and a
'timer has been inserted between the load and unload functions.

'This basic timer idea was also used by the big guns like Cheyenne's
'ArcServe!

'Load the cleaning tape

    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer1)
    ExecIO.SRB_CDBLen = &HC
    ExecIO.SRB_CDBByte(0) = &HA5            'move medium command
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = Label6.Caption  'MSB picker
    ExecIO.SRB_CDBByte(3) = Label8.Caption  'LSB picker
    ExecIO.SRB_CDBByte(4) = Label14.Caption 'MSB tape slot
    ExecIO.SRB_CDBByte(5) = (CleaningTape.Caption - 1) + Label16.Caption 'LSB tape slot
    ExecIO.SRB_CDBByte(6) = Label30.Caption 'MSB tape drive
    ExecIO.SRB_CDBByte(7) = Label32.Caption 'LSB tape drive
    ExecIO.SRB_CDBByte(8) = &H0
    ExecIO.SRB_CDBByte(9) = &H0
    ExecIO.SRB_CDBByte(10) = &H0
    ExecIO.SRB_CDBByte(11) = &H0
    
    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    
    Call Sleep(CleaningTapeTimeout * 1000)
    
    'Return the (ejected?) tape to the slot
    
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer1)
    ExecIO.SRB_CDBLen = &HC
    ExecIO.SRB_CDBByte(0) = &HA5
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = &H0
    ExecIO.SRB_CDBByte(3) = &H0
    ExecIO.SRB_CDBByte(4) = &H0
    ExecIO.SRB_CDBByte(5) = &H1
    ExecIO.SRB_CDBByte(6) = &H0
    ExecIO.SRB_CDBByte(7) = (CleaningTape.Caption - 1) + Label16.Caption
    ExecIO.SRB_CDBByte(8) = &H0
    ExecIO.SRB_CDBByte(9) = &H0
    ExecIO.SRB_CDBByte(10) = &H0
    ExecIO.SRB_CDBByte(11) = &H0
    
    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    
    MsgBox "Tape drive has been cleaned."
    
End Sub

Private Sub Command4_Click()
    'Reset the changer - essentially, inventory the slots.
    '(note. Inventory is different from *read*. All inventory does
    'is check for the *presence* of a tape, which can be read
    'by issuing command B8h.
    
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer1)
    ExecIO.SRB_CDBLen = &H6
    ExecIO.SRB_CDBByte(0) = &H7      'rezero the unit
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = &H0
    ExecIO.SRB_CDBByte(3) = &H0
    ExecIO.SRB_CDBByte(4) = &H0
    ExecIO.SRB_CDBByte(5) = &H0

    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    
End Sub
Private Sub DeviceListing_Click()
    
    TapesInSlots.Clear
    
'First, check to make sure this device is a medium changer

'scan device types
    DevType.SRB_Cmd = SC_GET_DEV_TYPE
    DevType.SRB_HaID = Adapterlist.ListIndex
    DevType.SRB_Flags = 0
    DevType.SRB_Hdr_Rsvd = 0
    DevType.SRB_Target = DeviceListing.ListIndex
    DevType.SRB_Lun = 0
    
    nRet = SendASPI32DevTypeEx(DevType)
    
    If DevType.DEV_DeviceType = DTYPE_JUKE Then

    'it is a changer
    
    SaveSettings.Enabled = True
    
    'get inquiry data regarding number of slots, addressing
    'and name info
    
    'This bit gets the name of the changer
    
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer2)
    ExecIO.SRB_CDBLen = &H6
    ExecIO.SRB_CDBByte(0) = &H12
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = &H0
    ExecIO.SRB_CDBByte(3) = &H0
    ExecIO.SRB_CDBByte(4) = &H24
    ExecIO.SRB_CDBByte(5) = &H0
    
    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    
    a$ = ""
    
    For o = 0 To 94
    If Val(Databuffer2.Info(o)) < 32 Or Val(Databuffer2.Info(o)) > 127 Then
    GoTo dd
    End If
    a$ = a$ + Chr(Databuffer2.Info(o))
dd:
    Next o
    
    Label37.Caption = a$
        
    Erase Databuffer2.Info()
    
    'This bit will get the element addressing information
    
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer2)
    ExecIO.SRB_CDBLen = &H6
    ExecIO.SRB_CDBByte(0) = &H1A
    ExecIO.SRB_CDBByte(1) = &H0
    ExecIO.SRB_CDBByte(2) = &H1D
    ExecIO.SRB_CDBByte(3) = &H0
    ExecIO.SRB_CDBByte(4) = &H16
    ExecIO.SRB_CDBByte(5) = &H0
    
    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    
    'Start from 6th byte
    
    Label6.Caption = Databuffer2.Info(6)
    Label8.Caption = Databuffer2.Info(7)
    Label10.Caption = Databuffer2.Info(8)
    Label12.Caption = Databuffer2.Info(9)
    Label14.Caption = Databuffer2.Info(10)
    Label16.Caption = Databuffer2.Info(11)
    Label18.Caption = Databuffer2.Info(12)
    Label20.Caption = Databuffer2.Info(13)
    Label22.Caption = Databuffer2.Info(14)
    Label24.Caption = Databuffer2.Info(15)
    Label26.Caption = Databuffer2.Info(16)
    Label28.Caption = Databuffer2.Info(17)
    Label30.Caption = Databuffer2.Info(18)
    Label32.Caption = Databuffer2.Info(19)
    Label34.Caption = Databuffer2.Info(20)
    Label36.Caption = Databuffer2.Info(21)
        
    'device is a medium changer... get slot information
    
    ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
    ExecIO.SRB_HaID = Adapterlist.ListIndex
    ExecIO.SRB_Flags = SRB_DIR_IN
    ExecIO.SRB_Hdr_Rsvd = 0
    ExecIO.SRB_Target = DeviceListing.ListIndex
    ExecIO.SRB_Lun = 0
    ExecIO.SRB_SenseLen = 14
    ExecIO.SRB_BufLen = &HFFF
    ExecIO.SRB_BufPointer = VarPtr(Databuffer1)
    ExecIO.SRB_CDBLen = &HC
    ExecIO.SRB_CDBByte(0) = &HB8
    ExecIO.SRB_CDBByte(1) = &H2
    ExecIO.SRB_CDBByte(2) = &H0
    ExecIO.SRB_CDBByte(3) = &H0
    ExecIO.SRB_CDBByte(4) = &H0
    ExecIO.SRB_CDBByte(5) = Label20.Caption
    ExecIO.SRB_CDBByte(6) = &H0
    ExecIO.SRB_CDBByte(7) = &HF
    ExecIO.SRB_CDBByte(8) = &HF
    ExecIO.SRB_CDBByte(9) = &H0
    ExecIO.SRB_CDBByte(10) = &H0
    ExecIO.SRB_CDBByte(11) = &H0
    
    nRet = SendASPI32ExecIOEx(ExecIO)
        While ExecIO.SRB_Status = SS_PENDING
        DoEvents
    Wend
    
    'Slot information is buried in Element status data.
    'It is rather messy, but at least it's constant!
    'It is defined thus:
    
    'Header info
    
    'Byte 0 is the element type (in our case - 2 = media)
    'Byte 1 is the Primary and Alternate volume tag
    'Byte 2 is the MSB of the element descriptor length
    'Byte 3 is the LSB of the element descriptor length
    'Byte 4 is reserved
    'Byte 5 is the MSB of the Byte count of descriptor data available
    'Byte 6 is the middle byte of the byte count of decriptor data available
    'Byte 7 is the LSB of the byte count of descriptor data available
    
    'From byte 8 to end of data, are the storage element descriptors
    
    'Byte 0 of the SED is the MSB of the element address
    'Byte 1 of the SED is the LSB of the element address
    'Byte 2 of the SED is the critical data for determining slot status.
    
    'Byte 2 looks like this.
    
    'Bit 7   Bit 6    Bit 5    Bit 4    Bit 3    Bit 2    Bit 1    Bit 0
    '---------------Reserved--------   Access    Except  Reserved  Full
    
    'So, the first four MSB bits are reserved.
    'Bit 3 indicates whether the tape can be used
    'Bit 2 indicates something unknown to me
    'Bit 1 indicates whether the slot has been reserved by the operator
    'Bit 0 indicates slot full(1) or empty (0)
    
    'If the bits were valued like this: "00001001" then
    'the changer's indicating access to this slot is allowed,
    '"Except" is off, the slot is not reserved, and a tape is loaded.
    
    'Or, if the bits were "00001000", all the above, and the slot is empty.
    
    'I assume the easist way to read this byte is to convert it
    'to binary, and read the LSB. That way we guarantee we
    'correctly read the full or empty bit.
    
    'The first 8 bytes (0-7) are the 'header' - we do not need this.
    
    j = 0
    c = 1
    
    For i = 8 To (Label20.Caption * 5) '5 bytes per tape
    
    x = Databuffer1.ElementStatusPages(i)
    
    If j = 2 Then
    
    'Convert x to long (decimal) (it's the number we need) and read the LSB
    
    lng = CLng("&H" & x)
    
    binValue = convDecToBin(lng)
    
    If Right(binValue, 1) = "1" Then
    
    If c = CleaningTape.Caption Then
    TapesInSlots.AddItem "Slot " & c & " (cleaning tape)"
    Else
    TapesInSlots.AddItem "Slot " & c & " tape loaded"
    End If
    
    Else
    TapesInSlots.AddItem "Slot " & c & " no tape loaded"
    End If
    c = c + 1
    End If
    
    j = j + 1
        
    If j = 4 Then j = 0
    Next i
    
    'Update the 'Next loaded' value. Label20.caption is the key.
    'Total slots must be taken in to account, as must the tape
    'tape cleaner, and whether a tape is actually loaded in the
    'slot!
    
    LastLoaded = CLng(LastLoaded.Caption)
    
    If Len(CleaningTape.Caption) = 0 Then
        TheCleaningTape = 0
    Else
        TheCleaningTape = CLng(CleaningTape.Caption)
    End If
    
    TotalTapes = CLng(Label20.Caption)
    
    TheNextTape = LastLoaded + 1
    
    If TheNextTape = TheCleaningTape Then TheNextTape = TheNextTape + 1
    
    If TheNextTape > TotalTapes Or TheNextTape = TotalTapes Then TheNextTape = 1
       
    NextLoaded.Caption = TheNextTape
      
    Command1.Enabled = False
    
    Command1.Caption = "Select tape from list"
    
    If Len(CleaningTape.Caption) = 0 Then
        Command2.Enabled = False
        Else
        Command2.Enabled = True
    End If
    
    Command4.Enabled = True
    SaveSettings.Enabled = True
    
    Else
    
    SaveSettings.Enabled = False
    Command1.Caption = ""
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    MsgBox "Device is not a changer"
    
    End If

End Sub

Private Sub Form_Load()

Dim f As Long
Dim a As String
Dim TheData As String

Open App.Path & "\changer.ini" For Input As #1
Do While Not EOF(1)
Line Input #1, a
    f = InStr(a, "=")
        If f Then
        TheData = UCase(Left$(a, f - 1))
    Else
        TheData = UCase(Trim$(a))
    End If
Select Case TheData
    Case "CLEANTAPE": CleaningTape.Caption = Mid$(a, f + 1)
    Case "CLEANHEADTIMEOUT": CleanHeadTimeout = CLng(Mid$(a, f + 1))
    Case "PROGRAM1": Program1.Caption = Mid$(a, f + 1)
    Case "PROGRAM2": Program2.Caption = Mid$(a, f + 1)
    Case "PROGRAM3": Program3.Caption = Mid$(a, f + 1)
    Case "SHUTDOWN1": Shutdown1.Caption = Mid$(a, f + 1)
    Case "SHUTDOWN2": Shutdown2.Caption = Mid$(a, f + 1)
    Case "SHUTDOWN3": Shutdown3.Caption = Mid$(a, f + 1)
    Case "LASTLOADED": LastLoaded.Caption = Mid$(a, f + 1)
    Case "ADAPTER": Adapter = CLng(Mid$(a, f + 1))
    Case "DEVICE": Device = CLng(Mid$(a, f + 1))
    End Select
Loop

Close #1

If Len(CleaningTape.Caption) < 1 Then
Command2.Enabled = False
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ASPI is installed properly if the following is TRUE
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'check available
    ASPI = AspiCheck
    If ASPI = False Then
        MsgBox "Error - no ASPI layer detected"
        Exit Sub
    End If

    'get adapter count
    cnt = AspiGetNumAdapters()
    If cnt = 0 Then
        MsgBox "Error - no adapters installed"
        Exit Sub
    End If

    'scan adapters for changers
    For i = 0 To (cnt - 1)
    
    'get inquiry data
        Inquiry.SRB_Cmd = SC_HA_INQUIRY
        Inquiry.SRB_HaID = i
        Inquiry.SRB_Flags = 0
        Inquiry.SRB_Hdr_Rsvd = 0

        nRet = SendASPI32InquiryEx(Inquiry)
        
        'info here can be shown from inquiry...
        
        Adapterlist.AddItem i & " " & Inquiry.HA_Ident
       
Next i

SaveSettings.Enabled = False

    End Sub

Private Sub Label50_Click()
Ms$ = "Written by Mark Borchers. Contact me at marximus27@hotmail.com."
Ms$ = Ms$ & vbCrLf
Ms$ = Ms$ & "Don't forget this is a work in progress. If it does work - great."
Ms$ = Ms$ & vbCrLf
Ms$ = Ms$ & "If not, sorry. I could only test this on Windows 2000, and an"
Ms$ = Ms$ & vbCrLf
Ms$ = Ms$ & "HP SureStore 818 DLT Autoloader. Please send me an email reporting"
Ms$ = Ms$ & vbCrLf
Ms$ = Ms$ & "if it does or does not work for you. Your OS and changer info would be appreciated."

MsgBox Ms$

End Sub

Private Sub Quit_Click()
Unload Me
End Sub

Private Sub SaveSettings_Click()

'This button saves the adapter and device selections so the
'program can go in to auto mode.

Open App.Path & "\changer.ini" For Output As #1

Print #1, "CleanTape=" & CleaningTape.Caption
Print #1, "CleanHeadTimeout=" & CleanHeadTimeout
Print #1, "Program1=" & Program1.Caption
Print #1, "Program2=" & Program2.Caption
Print #1, "Program3=" & Program3.Caption
Print #1, "Shutdown1=" & Shutdown1.Caption
Print #1, "Shutdown2=" & Shutdown2.Caption
Print #1, "Shutdown3=" & Shutdown3.Caption
Print #1, "LastLoaded=" & LastLoaded.Caption

'These are the adapter and device settings

Print #1, "Adapter=" & Adapterlist.ListIndex
Print #1, "Device=" & DeviceListing.ListIndex

Close #1

End Sub

Private Sub TapesInSlots_Click()

If Left(Command1.Caption, 4) = "Retu" Then Exit Sub

If Len(CleaningTape.Caption) > 0 Then

    If (TapesInSlots.ListIndex + 1) = CLng(CleaningTape.Caption) Then
    Command1.Enabled = False
    Command1.Caption = "Load tape not allowed"
    Exit Sub
    End If

If Right(TapesInSlots.List(TapesInSlots.ListIndex), 14) = "no tape loaded" Then
    Command1.Enabled = False
    Command1.Caption = "No tape loaded"
Else
    Command1.Enabled = True
    Command1.Caption = "Load tape " & (TapesInSlots.ListIndex + 1)
End If
End If

End Sub

Private Sub Timer1_Timer()

countdown = Label48.Caption

If countdown = 0 Then
    
Timer1.Enabled = False
    
    'load the next tape - do a basic select on the list items
    'to select the last saved configuration
    
    'select the adapter
    
    Adapterlist.Selected(Adapter) = True
    
    'select the device
    
    DeviceListing.Selected(Device) = True
       
    'select the tape
    
    TapesInSlots.Selected(NextLoaded.Caption - 1) = True
    
    'load the tape
    
    Command1_Click
        
    'do stuff
        
    If Len(Program1.Caption) > 1 Then
    nRet = ExecCmd(Program1.Caption)
    End If

    If Len(Program2.Caption) > 1 Then
    nRet = ExecCmd(Program2.Caption)
    End If
    
    If Len(Program3.Caption) > 1 Then
    nRet = ExecCmd(Program3.Caption)
    End If
       
    'Write out settings (basically, update 'tape last used' variable)
    
    Open App.Path & "\changer.ini" For Output As #1
    Print #1, "CleanTape=" & CleaningTape.Caption
    Print #1, "CleanHeadTimeout=" & CleanHeadTimeout
    Print #1, "Program1=" & Program1.Caption
    Print #1, "Program2=" & Program2.Caption
    Print #1, "Program3=" & Program3.Caption
    Print #1, "Shutdown1=" & Shutdown1.Caption
    Print #1, "Shutdown2=" & Shutdown2.Caption
    Print #1, "Shutdown3=" & Shutdown3.Caption
    Print #1, "LastLoaded=" & Label40.Caption

    'These are the adapter and device settings

    Print #1, "Adapter=" & Adapterlist.ListIndex
    Print #1, "Device=" & DeviceListing.ListIndex
    Close #1
       
    'Completely hide the program
    
    Me.Hide
    
    'Your external backup utility should now be launched.
    
    'When it's finished...
    
    'Unload the tape, then run the shutdown programs specified
    
    Me.Show
                
    Command1_Click
    
    If Len(Shutdown1.Caption) > 1 Then
    nRet = ExecCmd(Shutdown1.Caption)
    End If
    
    If Len(Shutdown2.Caption) > 1 Then
    nRet = ExecCmd(Shutdown2.Caption)
    End If
    
    If Len(Shutdown3.Caption) > 1 Then
    nRet = ExecCmd(Shutdown3.Caption)
    End If

Exit Sub

End If

countdown = countdown - 1

Label48.Caption = countdown

End Sub
