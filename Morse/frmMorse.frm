VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMorse 
   Caption         =   "Simple Morse Encoder"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSpk 
      Caption         =   "Use Big Speaker"
      Height          =   360
      Left            =   6900
      TabIndex        =   10
      Top             =   1440
      Width           =   2115
   End
   Begin VB.CommandButton cmdSwTimer 
      Caption         =   "Play Using &Software Timer"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   2460
   End
   Begin VB.TextBox txtFreq 
      Height          =   360
      Left            =   1485
      TabIndex        =   8
      Text            =   "510"
      Top             =   945
      Width           =   1230
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "&Encode"
      Height          =   375
      Left            =   7650
      TabIndex        =   5
      Top             =   90
      Width           =   1320
   End
   Begin MSComctlLib.Slider sldSpeed 
      Height          =   420
      Left            =   2880
      TabIndex        =   4
      Top             =   1380
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   741
      _Version        =   393216
      Min             =   5
      Max             =   50
      SelStart        =   15
      Value           =   15
      TextPosition    =   1
   End
   Begin VB.Timer tmr10 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1395
      Top             =   1800
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play using VB Timer"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   940
      Width           =   2535
   End
   Begin VB.TextBox txtMorse 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   450
      TabIndex        =   1
      Text            =   " .... . .-.. .-.. ---"
      Top             =   495
      Width           =   8520
   End
   Begin VB.TextBox txtMesg 
      Height          =   330
      Left            =   450
      TabIndex        =   0
      Text            =   "hello"
      Top             =   90
      Width           =   7125
   End
   Begin VB.TextBox txtTime 
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Text            =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Frequency:"
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Interval (Inverse of Speed):"
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   1485
      Width           =   2310
   End
End
Attribute VB_Name = "frmMorse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------
'   MORSE ENCODER v 2.0
'--------------------------------------------------------------------
' WHODUNNIT:
'   The whole code presented here is written by    : Harshad Sharma
'   (except where mentioned otherwise)
'
' WHAT IT DOES:
'   1. Takes in a string of charachers
'   2. Encodes it into a series of dots (.) and dashes (-) based on
'      International Morse Code
'   3. Plays back the sound (dah-di-dit) from the computer speaker
'   4. UPDATE !!! It also plays back the sound from the MAIN SPEAKERS!!!
'
' PURPOSE: To provide a template for building CW apps
'   Basically intended for those interested in HAM.
'   This code is just an example showing ease of programming in VB
'
'
' DISCLAIMER: This code is provided as-is, without warranty express or
'   implied. The author takes no responsibility for problems arising
'   from the use of this code. You can use it in way you like, but while
'   distributing either forward the set of files as you got them, or if
'   you have modified any file(s) state it that way.
'--------------------------------------------------------------------
' CONTACT: If you have any questions or comments, please feel free to
'   My e-mail  : harshad.sharma@bigfoot.com
'   My homepage: www20.brinkster.com/tanhadil
'--------------------------------------------------------------------
'   ALWAYS REMEMBER:
'                 "When everything else fails, read the instructions"
'--------------------------------------------------------------------


' For those who are totally blank about what the line below means:
' Option Explicit tells Visual Basic that it should not allow YOU
' to use any variables without first Declaring them.
Option Explicit

' Below is an Application Programming Interface (API) Call to windows
' It is used to make the sound from the PC Speaker
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

' I like to declare most of the often-used variables at once.
' It saves time in two ways... first I don't have to type more than once.
' Second VB does not have to reallocate memory for these variables
' every now and then. Thus even the program runs faster this way.
Private x As Integer            ' Used as counter in various functions
Private y As Integer            ' Used as counter
Private temp As String          ' Used to store temporary data
Private char As String          ' Used to store on charachter at a time



Private Sub cmdEncode_Click()
    ' take the text in txtMesg, encode it and put it into txtMorse
    txtMorse.Text = Morse.EncodeToMorse(txtMesg.Text)
End Sub

Private Sub cmdPlay_Click()
    ' select the text in the txtMesg textbox so that user can
    ' start typing next message or repeat the message if he/she wishes.
    txtMesg.SetFocus
    SendKeys "{HOME}+{END}"
    ' if the txtMorse textbox is empty, then first encode the text from
    ' txtMesg into txtMorse then call the function
    If txtMorse.Text = "" Then
        txtMorse.Text = Morse.EncodeToMorse(txtMesg.Text)
    End If
    ' call the function to play the encoded morse code
    PlayMorse
End Sub

Private Sub PlayMorse()
    ' the variables are declared publicly to make things faster
    ' make the variables ready for use.
    
Do While Len(txtMorse.Text) > 0
    ' reset the variables so that the function operates normally
    ' (The variables are changed while each loop,
    ' and must be reset before the next loop)
    If x = 0 Then x = 1
    y = 0
    
    ' store the string from the textbox into a temporary
    ' variable to speed up calculations...
    temp = txtMorse.Text
    ' get the next charachter
    char = Left(temp, 1)
    ' remove the charachter from the string and update the textbox
    txtMorse.Text = Right(temp, Len(temp) - 1)
    tmr10.Enabled = True
    Do While y < 1
        If char = "." Then
            ' give a short gap
            ' the line below makes the s/w go through the loop twice
            ' this is to allow even shorter time gap for the space
            ' and invalid charachters
            If Val(txtTime.Text) >= 2 Then
                txtTime.Text = 0        ' reset the time counter
                tmr10.Enabled = False   ' stop the timer for now
                                        ' (else it might trigger
                                        ' another instance of this function)
                                        
                If chkSpk.Value = 1 Then ' if user want to hear it loud...
                                        ' beep using our SimplestOscillator
                    dBeep Val(txtFreq.Text), tmr10.Interval * 1, 100
                Else
                                        ' beep using the API function
                    Beep Val(txtFreq.Text), tmr10.Interval * 8
                End If
                
                DoEvents                ' give windows some time to process other events
                y = 2                   ' set the flag so that the loop exits
            End If
        ElseIf char = "-" Then
            ' give a short gap
            If Val(txtTime.Text) >= 2 Then
                txtTime.Text = 0        ' reset the time counter
                tmr10.Enabled = False   ' stop the timer for now
                                        
                If chkSpk.Value = 1 Then ' if user want to hear it loud...
                                        ' beep using our SimplestOscillator
                    dBeep Val(txtFreq.Text), tmr10.Interval * 2, 100
                Else
                                        ' beep using the API function
                    Beep Val(txtFreq.Text), tmr10.Interval * 16
                End If
                
                DoEvents                ' give windows some time to process other events
                y = 2                   ' set the flag so that the loop exits
            End If
        ElseIf char = " " Then
            ' give a shorter gap
            If Val(txtTime.Text) >= tmr10.Interval Then
                txtTime.Text = 0        ' reset the time counter
                tmr10.Enabled = False   ' stop the timer for now
                DoEvents                ' give windows some time to process other events
                y = 2                   ' set the flag so that the loop exits
            End If
        Else
            'if anything else has creeped in, leave it there
                txtTime.Text = 0        ' reset the time counter
                tmr10.Enabled = False   ' stop the timer for now
                y = 2                   ' set the flag so that the loop exits
        End If
        DoEvents
    Loop
    DoEvents
Loop
End Sub

Private Sub cmdSwTimer_Click()
    ' select the text in the txtMesg textbox so that user can
    ' start typing next message or repeat the message if he/she wishes.
    txtMesg.SetFocus
    SendKeys "{HOME}+{END}"
    ' if the txtMorse textbox is empty, then first encode the text from
    ' txtMesg into txtMorse then call the function
    ' (NOTE 1) : If the textbox is filled, we simply transmits whatever is in it
    '            This feature it to allow users to type their own morse code and transmit it!
    If txtMorse.Text = "" Then
        txtMorse.Text = Morse.EncodeToMorse(txtMesg.Text)
    End If
    ' call the function to play the encoded morse code
    Call Morse.PlayMorseOnPCSpeaker(txtMorse.Text, sldSpeed.Value, Val(txtFreq.Text))
    ' because the software timer method does not update the textbox,
    ' we must clear it once the message is transmitted or else, our next message
    ' will not be transmitted (see NOTE 1)
    txtMorse.Text = ""
End Sub

Private Sub Form_Load()
    Me.Show
    DoEvents
    ' we have to first iniatialize the SimplestOscillator because is uses
    ' DirectX for outputting the sound.
    modOsc.InitializeBeep Me.hWnd
    ' Uncomment the lines below to play a
    ' little music when the app starts.
    'Beep 2000, 200
    'Beep 1000, 100
    'Beep 2500, 200
End Sub

Private Sub sldSpeed_Click()
    ' update the timer's interval
    ' The advantage of the VB timer is that you can change the speed
    ' in the middle of a message while it is being transmitted!
    tmr10.Interval = sldSpeed.Value
End Sub

Private Sub tmr10_timer()
    ' Increment the value in txtTime by one
    txtTime.Text = Val(txtTime.Text) + 1
End Sub
