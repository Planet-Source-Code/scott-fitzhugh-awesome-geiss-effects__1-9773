VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Some nice effects"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Stop FPS Meter"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4335
      TabIndex        =   10
      Top             =   2085
      Width           =   1950
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start Fake Analyzer"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4335
      TabIndex        =   8
      Top             =   1470
      Width           =   1950
   End
   Begin VB.Timer Timer3 
      Left            =   6165
      Top             =   810
   End
   Begin VB.Timer Timer2 
      Left            =   5625
      Top             =   810
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Draw Horizontal Line"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6390
      TabIndex        =   6
      Top             =   1470
      Width           =   1950
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Prints Text onto Graphic"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6390
      TabIndex        =   5
      Top             =   2085
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Color Palette.  (CLICK ME FIRST!)"
      Height          =   495
      Left            =   6150
      TabIndex        =   3
      Top             =   4005
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Left            =   5100
      Top             =   810
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Begins Effect"
      Enabled         =   0   'False
      Height          =   600
      Left            =   6150
      TabIndex        =   1
      Top             =   4650
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   660
      Left            =   45
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   0
      Top             =   5415
      Width           =   8355
   End
   Begin VB.PictureBox Picture2 
      ForeColor       =   &H8000000D&
      Height          =   2850
      Left            =   120
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   2
      Top             =   75
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "©2000 - Scott Fitzhugh - All Rights Reserved - http://www.themeshops.com"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4590
      Width           =   6435
   End
   Begin VB.Line Line1 
      X1              =   4335
      X2              =   7245
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Palette Bar"
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
      Left            =   75
      TabIndex        =   11
      Top             =   5085
      Width           =   5640
   End
   Begin VB.Label Label3 
      Caption         =   "Frames/s"
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
      Left            =   4350
      TabIndex        =   9
      Top             =   150
      Width           =   930
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   285
      Left            =   5460
      TabIndex        =   7
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "click and drag mouse around in the little black box above this text"
      Height          =   600
      Left            =   180
      TabIndex        =   4
      Top             =   3015
      Visible         =   0   'False
      Width           =   3690
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################
'#                                                                              #
'#  Some Cool Effects without DirectX                                           #
'#                                                                              #
'#  Author: Scott Fitzhugh                                                      #
'#  Date: Friday, July 14, 2000                                                 #
'#  Description: Basically this is just something that                          #
'#      I did to teach myself some graphics programming                         #
'#      with palettes, and arrays, and the dredded matrix.                      #
'#      I'll admit that it's slow.  It's not directx.  But                      #
'#      for someone who's just trying to learn, this is an                      #
'#      EXCELLENT program to try to emulate.  If I can ever                     #
'#      port this over to directx someday, the fading effect                    #
'#      emulates the "Geiss" fader pretty well I think...                       #
'#      Who knows... maybe we'll be able to get it fullscreen                   #
'#      someday...  lol... yeah right (=                                        #
'#                                                                              #
'#      PS: if you compile it and run the compiled code, it's                   #
'#      ALOT faster.  ::wink::                                                  #
'#                                                                              #
'#      EMAIL ME if you have any comments, suggestions, or want                 #
'#      to help me make this in direct x                                        #
'#                                                                              #
'#             arsecannon@yahoo.com                                             #
'#                                                                              #
'#                                                                              #
'#  ©2000 - Scott Fitzhugh - All Rights Reserved - http://www.themeshops.com    #
'################################################################################

Dim map(999, 999)   'image map
Dim pallete(255)    'palette
Dim plheight, plwidth   'height and width of drawing area
Dim fader   'fading multiplier
Dim frames As Single    'frames per second variable

Private Sub Command1_Click()
    Picture1.Picture = LoadPicture(App.Path + "\palette.bmp")   'load palette image
    
    For v = 0 To 255 Step 1 'scroll through and
        pallete(v) = Picture1.Point(v, 1)   'create pallete from image
    Next v
    
    Picture2.ForeColor = pallete(254)   'set print color for drawing area
    
    Command2.Enabled = True 'make "begin" button enabled
    Command1.Enabled = False    'disable palette button
End Sub

Private Sub Command2_Click()
        'toggle timer1 (eg. main fading effect timer)
    If Timer1.Interval = 0 Then
        Timer1.Interval = 1
        Label1.Visible = True
    Else
        Timer1.Interval = 0
        Label1.Visible = False
    End If
    
    Command2.Enabled = False    'disable begin button and enable tiny effect buttons
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Label1.Visible = True
    
    Timer2.Interval = 1000  'enable frames per second timer
End Sub

Private Sub Command3_Click()
    
    Picture2.CurrentX = 10  'set x,y coordinates and print "arse" onto image
    Picture2.CurrentY = 10
    Picture2.Print "arse"
    For x = 0 To plwidth    'recurse through image map and incorporate new pixels
            For y = 0 To plheight
                If Picture2.Point(x, y) = pallete(254) Then map(x, y) = 254
            Next y
    Next x

End Sub

Private Sub Command4_Click()
        For x = 0 To plwidth    'draw line at a height of 25
            map(x, 25) = 254
        Next x
End Sub

Private Sub Command5_Click()
        'toggle timer3 or the Analyzer timer
    If Timer3.Interval = 500 Then
        Timer3.Interval = 0
        Command5.Caption = "Start Fake Analyzer"
    Else
        Timer3.Interval = 500
        Command5.Caption = "Stop Fake Analyzer"
    End If
    
End Sub

Private Sub Command6_Click()
    
        'toggle timer2 or the frames per second timer
    If Timer2.Interval = 0 Then
        Timer2.Interval = 1000
        Command6.Caption = "Stop FPS Meter"
    Else
        Timer2.Interval = 0
        Command6.Caption = "Start FPS Meter"
    End If
End Sub

Private Sub Form_Load()
    'set height and width for drawing area.
    'I made this on a AMD athlon 750 with 650meg of ram...
    'and 100x100 is the largest I can make it and still get a descent
    'frame rate...  Do a little testing to figure out what you can do.
    'It also helps to compile your code and run it like that.  Frame rates DO increase,
    'and graphics are much more fluid. (framerates of 8-40 are normal... it's not directx)
    plheight = 80
    plwidth = 80
    
    'set fade multiplier
    'a larger number fades more slowly and some numbers get REALLY weird results
    fader = 2
        
    'define loop variables as single to improve speed slightly
    Dim x As Single
    Dim y As Single
        
    'create black imagemap
    For x = 0 To plheight
        For y = 0 To plwidth
            map(x, y) = 1
        Next y
    Next x

    Picture2.Font = verdana
    Picture2.FontSize = 8
    Picture2.FontBold = True
End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'create drag effect when button 1 is pressed
    If Button = 1 Then
        'define loop variables as single to improve speed slightly
        Dim h As Single, g As Single
        For h = 1 To 10
            For g = 1 To 10
                If (x + h) > 0 And (x + h) < plwidth And (y + g) > 0 And (y + g) < plheight Then map(x + h, y + g) = 255
            Next g
        Next h
        
    End If
End Sub

Private Sub Timer1_Timer()
        
    'define loop variables as single to improve speed slightly
    Dim x As Single
    Dim y As Single
    
    For x = 0 To plwidth
        For y = (plheight / 2) To 0 Step -1
            
            'testcode
                If (y - 1) > 0 Then map(x, y - 1) = (((map(x, y - 1) * (fader - 1)) + map(x, y)) / fader) - 1
                If (y + 1) < plheight Then map(x, y + 1) = (((map(x, y + 1) * (fader - 1)) + map(x, y)) / fader) - 1
                If (x - 1) > 0 Then map(x - 1, y) = (((map(x - 1, y) * (fader - 1)) + map(x, y)) / fader) - 1
                If (x + 1) < plwidth Then map(x + 1, y) = (((map(x + 1, y) * (fader - 1)) + map(x, y)) / fader) - 1
                
            'testcode
            'If Timer2.Interval = 1000 And Picture2.Point(X, Y) = pallete(254) Then
            'Else
                If map(x, y) < 0 Then map(x, y) = 0
                If map(x, y) > 255 Then map(x, y) = 255
                Picture2.PSet (x, y), pallete(map(x, y))
            'End If
        Next y
        For y = ((plheight / 2) + 1) To plheight
            
            'testcode
                If (y - 1) > 0 Then map(x, y - 1) = (((map(x, y - 1) * (fader - 1)) + map(x, y)) / fader) - 1
                If (y + 1) < plheight Then map(x, y + 1) = (((map(x, y + 1) * (fader - 1)) + map(x, y)) / fader) - 1
                If (x - 1) > 0 Then map(x - 1, y) = (((map(x - 1, y) * (fader - 1)) + map(x, y)) / fader) - 1
                If (x + 1) < plwidth Then map(x + 1, y) = (((map(x + 1, y) * (fader - 1)) + map(x, y)) / fader) - 1
                
            'testcode
            'If Timer2.Interval = 1000 And Picture2.Point(X, Y) = pallete(254) Then
            'Else
                If map(x, y) < 0 Then map(x, y) = 0
                If map(x, y) > 255 Then map(x, y) = 255
                Picture2.PSet (x, y), pallete(map(x, y))
            'End If
        Next y
    Next x
    'increase frame/s variable
    frames = frames + 1
    
End Sub

Private Sub Timer2_Timer()
    'one second has passed, print it's value and reset it
    Label2.Caption = frames
    frames = 0
End Sub

Private Sub Timer3_Timer()
    'define loop variables as single to improve speed slightly
    Dim f As Single, h As Single

Randomize
'GenRndNumber = Int((Upper - Lower + 1) * Rnd + Lower)

Dim crazy As Single
crazy = 10  'change crazy to get a more wild analyzer pattern (higher more energy)

h = 25
f = 25

For j = 1 To plwidth
    If (j - 10) >= 0 Then
        f = Int((h + crazy) - (h - crazy) + 1) * Rnd + (h - crazy)
    Else
        f = Int(crazy + 1) * Rnd
    End If
    
    If f <= plheight And h <= plheight Then Picture2.Line (j, f)-(j - 1, h), 255    'draw new line
    
    h = f   'set new as old
Next j

    
    'incorperate pattern into imagemap
    For x = 0 To plwidth
        For y = 0 To plheight
            If Picture2.Point(x, y) = 255 Then map(x, y) = Picture2.Point(x, y)
        Next y
    Next x

End Sub
