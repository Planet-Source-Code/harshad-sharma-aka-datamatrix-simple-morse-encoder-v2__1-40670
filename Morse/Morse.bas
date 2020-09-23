Attribute VB_Name = "Morse"
' For those who are totally blank about what the line below means:
' Option Explicit tells Visual Basic that it should not allow YOU
' to use any variables without first Declaring them.
Option Explicit

' Below is an Application Programming Interface (API) Call to windows
' It is used to make the sound from the PC Speaker
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
' This API call is used to run the inbuilt timer...
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
' This TYPE is necessary if we want to get the System Time through the API
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

' I like to declare most of the often-used variables at once.
' It saves time in two ways... first I don't have to type more than once.
' Second VB does not have to reallocate memory for these variables
' every now and then. Thus even the program runs faster this way.
Private x As Integer            ' Used as counter in various functions
Private y As Integer            ' Used as counter
Private temp As String          ' Used to store temporary data
Private char As String          ' Used to store on charachter at a time
Private MSec As Integer         ' Used for counting / storing Micro Seconds
Private MilliSec As Integer     ' Used for counting / storing Milli Seconds
Private SysTime As SYSTEMTIME   ' Used to store the System Time
Private RunTimer As Boolean     ' Used to tell the software to exit timer loop
                                ' from outside procedure. 'If RunTimer = False Then Exit Loop'

Public Function EncodeToMorse(TextToEncode As String) As String
Dim x As Integer
Dim y As Integer
Dim char As String
Dim EncodedMorse As String
If Len(TextToEncode) > 1 Then
    For x = 0 To Len(TextToEncode) - 1
        char = Left(TextToEncode, 1)
        TextToEncode = Right(TextToEncode, Len(TextToEncode) - 1)
        EncodedMorse = EncodedMorse & " " & EncodeCharachter(char)
    Next
    EncodeToMorse = EncodedMorse
Else
    EncodeToMorse = ""
End If
End Function

Public Function PlayMorseOnPCSpeaker(aEncodedMorse As String, Optional aInterval As Integer, Optional aFrequency As Integer)
'----------------------------
' Little bit errorhandling:
'----------------------------
' 1. If we encounter an unknown error...
'    just resume the next line of code
On Error Resume Next
' 2. If the aInterval is not specified, then give it some default value
If aInterval < 1 Then aInterval = 15
' 3. If the EncodedMorse variable is not filled, then fill it with
'    default value : a single space
' Use the trim function to first remove extra spaces
' and then see if there is anything else
If Trim(aEncodedMorse) = "" Then aEncodedMorse = " "
' Use the Len function to check the size of the string
' A valid string contains more than one charachter...
If Len(aEncodedMorse) = 0 Then aEncodedMorse = " "
' 4. If the aFrequency variable is blank or incorrect, then give it a default value:
' I have checked on my system speaker that 480 - 510 Hz gives the
' loudest signals at the lowest frequencies.
If aFrequency < 1 Then aFrequency = 510


'----------------------------
' Main  Function Starts Here
'----------------------------

Do While Len(aEncodedMorse) > 0
    ' takes the left-most single charachter from the given string
    char = Left(aEncodedMorse, 1)
    
    ' removes the left-most charachter from the string
    aEncodedMorse = Right(aEncodedMorse, Len(aEncodedMorse) - 1)
    
    ' we have to reset the flag 'y' before we proceed with the loop
    ' finction because, as the variable is not limited to this sub,
    ' it's tendency is to remain static... hence if
    ' the previous loop might have set it to 2 (which make it to exit)
    ' WE must reset it to 0 or else it will no loop after the first iteration!
    y = 0
    
    
    Do While y < 1
        ' the same resetting rule applies here...
        MSec = 0
        MilliSec = 0
        
        ' this is one easy method to implement a timer through software...
        ' altough I had to give it a lot of thought, it here for you!
        Do While MSec < aInterval
            ' this is the timer function...
            ' all it does it wait for the system clock to show
            ' increment in the time. till that time, it gives
            ' other events more priority
            DoEvents
            Call GetSystemTime(SysTime)     ' get the system time
            If Abs(SysTime.wMilliseconds - MilliSec) >= 1 Then
                MilliSec = SysTime.wMilliseconds
                MSec = MSec + 1
            End If
            
            ' This method is NOT very accurate, nor is it very efficient
            ' It also consumes a lot of system time.
            ' But it is enough for our use. If you encounter problems,
            ' please use the VB Timer Method
        Loop        ' end statement for: Do While MSec < Minterval
    '-----------------------------------
        ' check if the charachter is a dot,
        ' if yes then play a short beep (Interval * 8) milliseconds
        If char = "." Then
            If MSec >= 2 Then
                Beep aFrequency, aInterval * 8
                DoEvents
                y = 2
            End If
        ' check if the charachter is a dash,
        ' if yes then play a long beep (Interval * 16) milliseconds
        ElseIf char = "-" Then
            If MSec >= 2 Then
                Beep aFrequency, aInterval * 18
                DoEvents
                y = 2
            End If
        ' check if the charachter is a space,
        ' if yes then do not play anything
        ElseIf char = " " Then
                ' all we need is a little gap, just wait
                If MSec >= aInterval / 2 Then
                    y = 2
                End If
                DoEvents
                DoEvents
                DoEvents
                DoEvents
        Else
            ' if anything else has creeped in, leave it there
            ' and immediately to next charachter
                y = 2
        End If
        
        DoEvents
    Loop ' End statement for: Do While y < 1
    '-----------------------------------
    
    
    DoEvents    ' Tell windows to process other events from this app
Loop            ' End statement for: Do While Len(EncodedMorse > 0)
End Function


Private Function EncodeCharachter(aCharachter As String) As String
aCharachter = LCase(aCharachter)
Select Case aCharachter
    Case " "
        EncodeCharachter = "   "
    Case "a"
        EncodeCharachter = ".-"
    Case "b"
        EncodeCharachter = "-..."
    Case "c"
        EncodeCharachter = "-.-."
    Case "d"
        EncodeCharachter = "-.."
    Case "e"
        EncodeCharachter = "."
    Case "f"
        EncodeCharachter = "..-."
    Case "g"
        EncodeCharachter = "--."
    Case "h"
        EncodeCharachter = "...."
    Case "i"
        EncodeCharachter = ".."
    Case "j"
        EncodeCharachter = ".---"
    Case "k"
        EncodeCharachter = "-.-"
    Case "l"
        EncodeCharachter = ".-.."
    Case "m"
        EncodeCharachter = "--"
    Case "n"
        EncodeCharachter = "-."
    Case "o"
        EncodeCharachter = "---"
    Case "p"
        EncodeCharachter = ".--."
    Case "q"
        EncodeCharachter = "--.-"
    Case "r"
        EncodeCharachter = ".-."
    Case "s"
        EncodeCharachter = "..."
    Case "t"
        EncodeCharachter = "-"
    Case "u"
        EncodeCharachter = "..-"
    Case "v"
        EncodeCharachter = "...-"
    Case "w"
        EncodeCharachter = ".--"
    Case "x"
        EncodeCharachter = "-..-"
    Case "y"
        EncodeCharachter = "-.--"
    Case "z"
        EncodeCharachter = "--.."
    Case "1"
        EncodeCharachter = ".----"
    Case "2"
        EncodeCharachter = "..---"
    Case "3"
        EncodeCharachter = "...--"
    Case "4"
        EncodeCharachter = "....-"
    Case "5"
        EncodeCharachter = "....."
    Case "6"
        EncodeCharachter = "-...."
    Case "7"
        EncodeCharachter = "--..."
    Case "8"
        EncodeCharachter = "---.."
    Case "9"
        EncodeCharachter = "----."
    Case "0"
        EncodeCharachter = "-----"
    Case "."
        EncodeCharachter = ".-.-.-"
    Case "?"
        EncodeCharachter = "..--.."
    Case ","
        EncodeCharachter = "--..--"
    Case "'"
        EncodeCharachter = ".----."
    ' The other charachters which are not listed here
    ' are (as per my knowledge) NOT supported by the MORSE Code.
    '  If you know any which are mistakenly left out, then just
    ' copy some code from above lines and add the new charachters.
    ' And please mail me about it (so I can update this project accordingly!)
    ' My email is : harshad.sharma@bigfoot.com
End Select
End Function

