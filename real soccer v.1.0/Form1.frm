VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Real Soccer v.1.0 : By Garz0r7"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   0
      Top             =   8880
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   360
      Top             =   8880
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   2640
      Top             =   8880
   End
   Begin VB.Timer gk_timer 
      Interval        =   100
      Left            =   3120
      Top             =   8880
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2160
      Top             =   8880
   End
   Begin VB.Timer ball_move_timer 
      Left            =   1680
      Top             =   8880
   End
   Begin VB.Timer gk_kick_timer 
      Left            =   840
      Top             =   8880
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   8880
   End
   Begin VB.Label story 
      Caption         =   "Label1"
      Height          =   6975
      Left            =   7080
      TabIndex        =   2
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label score 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   6975
   End
   Begin VB.Label what 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   6975
   End
   Begin VB.Shape b 
      Height          =   135
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   9000
      Width           =   135
   End
   Begin VB.Shape goal 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   375
      Left            =   2520
      Top             =   0
      Width           =   1935
   End
   Begin VB.Shape p2 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   3600
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape p1 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2520
      Shape           =   1  'Square
      Top             =   6360
      Width           =   135
   End
   Begin VB.Shape stadium 
      Height          =   6975
      Left            =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Shape gk 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   3360
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape ball 
      BackStyle       =   1  'Opaque
      FillStyle       =   7  'Diagonal Cross
      Height          =   135
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   135
   End
   Begin VB.Shape shadow 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00004000&
      Height          =   110
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   4200
      Width           =   110
   End
   Begin VB.Shape cir 
      BorderColor     =   &H80000007&
      Height          =   255
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape t2 
      BorderColor     =   &H80000005&
      Height          =   1695
      Left            =   1320
      Top             =   0
      Width           =   4455
   End
   Begin VB.Shape t1 
      BorderColor     =   &H80000005&
      Height          =   855
      Left            =   1920
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   7020
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem==============
Rem garz0r7 ;P
Rem aceidstudent
Rem patra,greece
Rem 210706
Rem=============

Const ball_speed = 50
Dim kk As Integer
Dim nul As Integer
Dim ball_x As Integer
Dim ball_y As Integer
Dim ball_sx As Integer
Dim ball_sy As Integer
Dim ball_h As Single
Dim player_has_ball As Integer
Dim ball_down As Integer
Dim g1 As Integer
Dim g2 As Integer
Dim sg1 As String
Dim sg2 As String
Const min_dist_player_ball = 100
Const min_dist_player_player = 100
Const min_dist_player_player_to_do_drible = 100 'player2 tries dribling after this value
Const min_dist_player_player_to_try_steal = 100 'player2 tries stealing after this value
'players stats
Dim gk_speed As Integer
Dim gk_ability As Integer
Dim p1_speed_with_ball As Integer
Dim p2_speed_with_ball As Integer
Dim p1_speed As Integer
Dim p2_speed As Integer
Dim p1_stealing As Integer
Dim p2_stealing As Integer
Dim p1_dribling As Integer
Dim p2_dribling As Integer
Dim p1_crossing As Integer
Dim p2_crossing As Integer
Dim p1_shoting As Integer
Dim p2_shoting As Integer
Dim p1_pressing As Integer
Dim p2_pressing As Integer
Dim p2_tries_long_shots As Integer
Dim p1_tries_long_shots As Integer
Dim p1_name As String
Dim p2_name As String
Option Explicit
Private Sub Form_Load()
'build the stadium
nul = build_stadium()
'loading player stats
p1_name = "Messi" 'your name
p2_name = "Ronaldinho"

gk_ability = 20
gk_speed = 40
p1_speed = 30 'when he has not ball.So must be>from speed with ball!
p2_speed = 25 'when he has not ball.So must be>from speed with ball!

'well dont make speed under 20..it's unrealistic to be sooo slow.
p2_speed_with_ball = 25
p1_speed_with_ball = 20

p1_stealing = 10
p2_stealing = 10

p1_dribling = 10
p2_dribling = 10

p2_crossing = 10
p1_crossing = 15

'if you have good shoting,you have more chances to score
p2_shoting = 10
p1_shoting = 10

'if pressing is bad..then excelent stealing means nothing.
'both is needed to steal oponents ball.
p1_pressing = 10
p2_pressing = 10

' p2 tries long shots,that means that he is talented on scoring!!
'from large distances!!
'try to be in frond of him and close to him
'to avoid his shots!!
'else,if you are back of him,he can shot as he
'realise that it is a good chance to score!!
p2_tries_long_shots = 10
p1_tries_long_shots = 10


'start game!
nul = gk_kick()
nul = write_story()
End Sub
Function write_story()
story = story + vbNewLine + ""
story = "Use Arrows To Move." + vbNewLine + "You Are The Red(player1) Player"
story = story + vbNewLine + "s->try drible the opponent"
story = story + vbNewLine + "a->try steal the ball ( you must be close to player2 )"
story = story + vbNewLine + "d->try shot ( closer to the goal,you have more chances )"


End Function
Private Sub ball_move_timer_Timer()
'handles ball moving when kicked
If shadow.Top < (ball_y) / 2 Then
    ball_h = ball_h + 100
End If
If shadow.Top >= (ball_y) / 2 Then
    ball_h = ball_h - 100
End If
If ball_h <= 0 Then
    ball_h = 0
End If

shadow.Top = shadow.Top + (ball_y - shadow.Top) / 10
shadow.Left = shadow.Left + (ball_x - shadow.Left) / 10
ball.Top = shadow.Top - ball_h
ball.Left = shadow.Left

If ball.Top >= ball_y - 100 Then
    ball_down = 1
    shadow.Visible = False
    ball_move_timer.Interval = 0
    gk_timer.Interval = 100
End If

End Sub


Function build_stadium()
' builds stadium
stadium.Left = 0
stadium.Top = 0
stadium.Width = 7000
stadium.Height = 6000
goal.Width = 2000
goal.Height = 400
goal.Left = stadium.Width / 2 - goal.Width / 2
goal.Top = stadium.Top + 50
gk.Left = goal.Left + goal.Width / 2
gk.Top = goal.Top + goal.Height - gk.Height / 2
p1.Top = stadium.Top + stadium.Height / 3
p2.Top = stadium.Top + stadium.Height / 3
p1.Left = stadium.Left - 200 + stadium.Width / 2
p2.Left = stadium.Left + 200 + stadium.Width / 2
ball.Width = 100
ball.Height = 100
shadow.Width = 100
shadow.Height = 100
score.Top = stadium.Top + stadium.Height + 100
score.Left = stadium.Left
score.Width = stadium.Width
what.Top = score.Top + score.Height
what.Width = stadium.Width
what.Left = stadium.Left
ball.Top = gk.Top + 200
ball.Left = gk.Left
t1.Width = 1.2 * goal.Width
t2.Width = 2 * goal.Width
t1.Height = 2.5 * goal.Height
t2.Height = 4 * goal.Height
t1.Left = stadium.Left + stadium.Width / 2 - t1.Width / 2
t1.Top = stadium.Top
t2.Left = stadium.Left + stadium.Width / 2 - t2.Width / 2
t2.Top = stadium.Top
b.Visible = False

End Function
Function gk_kick()

gk_timer.Interval = 0
ball_down = 0
player_has_ball = 0
gk.Left = goal.Left + goal.Width / 2
gk.Top = goal.Top + goal.Height - gk.Height / 2
ball.Top = gk.Top + 200
ball.Left = gk.Left
shadow.Top = ball.Top
shadow.Left = ball.Left
shadow.Visible = True
p1.Top = stadium.Top + stadium.Height / 3
p2.Top = stadium.Top + stadium.Height / 3
p1.Left = stadium.Left - 200 + stadium.Width / 2
p2.Left = stadium.Left + 200 + stadium.Width / 2

what = "Goalkeeper Kicks The Ball"
gk_kick_timer.Interval = 100

End Function
Private Sub gk_kick_timer_Timer()
'goalkeeper goes towards ball
gk.Top = gk.Top + gk_speed
'if he is close to ball ,he kick her
If gk.Top >= ball.Top - 40 Then
    Randomize Timer
    'ball_x,ball_y =where ball will go
    ball_x = (Rnd() * (0.8 * stadium.Width - stadium.Left)) + stadium.Left
    ball_y = (Rnd() * (0.5 * stadium.Height - stadium.Height / 3)) + stadium.Height / 3
    b.Top = ball_y
    b.Left = ball_x
    ball_sx = ball.Left
    ball_sy = ball.Top
    'ball was kicked and goes at potision ball_x,ball_y
    gk_kick_timer.Interval = 0
    ball_move_timer.Interval = 100
End If
End Sub

Private Sub gk_timer_Timer()
'goalkeeper artificial inteligence
'if player who has ball close to goalkeeper
'and goalkeeper has good ability stats then
'he goes steal the ball
Dim k As Integer

If p1.Top - (goal.Top + goal.Height) < 600 * gk_ability / 20 And player_has_ball = 1 Then
    k = move_gk(p1)
End If

If p2.Top - (goal.Top + goal.Height) < 600 * gk_ability / 20 And player_has_ball = 2 Then
    k = move_gk(p2)
End If

If player_has_ball = 0 Then
    k = move_gk_back()
End If

gk.Left = gk.Left + Rnd() * 20 - 10
End Sub
Function move_gk(p As Shape)
'goalkeeper follows player with ball
'cauze he is close to the goal
Randomize Timer
If gk.Left < p.Left Then
    gk.Left = gk.Left + gk_speed
End If

If gk.Left > p.Left Then
    gk.Left = gk.Left - gk_speed
End If

If gk.Top < p.Top Then
    gk.Top = gk.Top + gk_speed
End If

If gk.Top > p.Top Then
    gk.Top = gk.Top - gk_speed
End If

'goalkeeper clears the ball if close to player
If distance(gk, p) < 100 Then
    player_has_ball = 0
    ball.Top = ball.Top + (Rnd() * 1000) + 500
    what = "Goalkeeper Clears The Ball"
End If

End Function
Function move_gk_back()
'goalkeeper returns in his posiotion
If gk.Left + gk.Width / 2 < goal.Left + goal.Width / 2 And player_has_ball = 0 Then
    gk.Left = gk.Left + gk_speed
End If

If gk.Left + gk.Width / 2 > goal.Left + goal.Width / 2 And player_has_ball = 0 Then
    gk.Left = gk.Left - gk_speed
End If

If gk.Top + gk.Height / 2 > goal.Top + goal.Height And player_has_ball = 0 Then
    gk.Top = gk.Top - gk_speed
End If

If gk.Top + gk.Height / 2 < goal.Top + goal.Height And player_has_ball = 0 Then
    gk.Top = gk.Top + gk_speed
End If

End Function
Private Sub Timer1_Timer()
'handles your movement
Dim s As Integer

If player_has_ball = 1 Then
    s = p1_speed_with_ball
Else
    s = p1_speed
End If

If GetAsyncKeyState(vbKeyUp) Then
    p1.Top = p1.Top - s
End If

If GetAsyncKeyState(vbKeyDown) Then
    p1.Top = p1.Top + s
End If

If GetAsyncKeyState(vbKeyLeft) Then
    p1.Left = p1.Left - s
End If

If GetAsyncKeyState(vbKeyRight) Then
    p1.Left = p1.Left + s
End If

If GetAsyncKeyState(vbKeyS) And player_has_ball = 1 Then
'you dribling
    If GetAsyncKeyState(vbKeyUp) Then
        ball.Top = ball.Top - p1_crossing * 20
        player_has_ball = 0
    End If
    If GetAsyncKeyState(vbKeyDown) Then
        ball.Top = ball.Top + p1_crossing * 20
        player_has_ball = 0
    End If
    If GetAsyncKeyState(vbKeyLeft) Then
        ball.Left = ball.Left - p1_crossing * 20
        player_has_ball = 0
    End If
    If GetAsyncKeyState(vbKeyRight) Then
        ball.Left = ball.Left + p1_crossing * 20
        player_has_ball = 0
    End If
End If


If distance(p1, ball) < 200 And player_has_ball = 0 And ball_down = 1 Then
    player_has_ball = 1
    what = p1_name + " Get The Ball!"
End If

If player_has_ball = 1 Then
'if you have ball,then ball "follows" you
    If GetAsyncKeyState(vbKeyUp) Then
        ball.Top = p1.Top + p1.Height / 2 - ball.Height / 2 - Rnd() * 90
        ball.Left = p1.Left + p1.Width / 2 - ball.Width / 2
    End If
    If GetAsyncKeyState(vbKeyDown) Then
        ball.Top = p1.Top + p1.Height / 2 - ball.Height / 2 + Rnd() * 90
        ball.Left = p1.Left + p1.Width / 2 - ball.Width / 2
    End If
    If GetAsyncKeyState(vbKeyLeft) Then
        ball.Top = p1.Top + p1.Height / 2 - ball.Height / 2
        ball.Left = p1.Left + p1.Width / 2 - ball.Width / 2 - Rnd() * 90
    End If
    If GetAsyncKeyState(vbKeyRight) Then
        ball.Top = p1.Top + p1.Height / 2 - ball.Height / 2
        ball.Left = p1.Left + p1.Width / 2 - ball.Width / 2 + Rnd() * 90
    End If
End If

'player(you) shots!!!
If GetAsyncKeyState(vbKeyD) And player_has_ball = 1 Then
    what = p1_name + " Shots!!"
    nul = shot(p1)
End If

'player1(you) tries stealing
If GetAsyncKeyState(vbKeyA) And player_has_ball = 2 Then
    nul = try_steal(p1, p2)
End If
End Sub

Function distance(s1 As Shape, s2 As Shape)
distance = Sqr((s1.Left + s1.Width / 2 - s2.Left - s2.Width / 2) ^ 2 + (s1.Top + s1.Height / 2 - s2.Top - s2.Height / 2) ^ 2)
End Function
Function try_steal(s1 As Shape, s2 As Shape)
'**** s1=who want to steal the ball.
's2=who has the ball.


Randomize Timer
Dim x1 As Integer
Dim x2 As Integer
x1 = Rnd() * 100
x2 = Rnd() * 100

'if player1(you) wants to steal the ball
'then check if
'1)he is close to player2
'2)he has good pressing stats
'3)he has bettet stealing stats then player2  dribling stats

If s1 = p1 Then
'p1 tries steal from p2
    x1 = Rnd() * p1_stealing
    x2 = Rnd() * p2_dribling
    If x1 > x2 And distance(s1, s2) < min_dist_player_player_to_try_steal And x1 <= p1_pressing Then
        player_has_ball = 1
        p1.Top = p2.Top + 300
        what = p1_name + " Steals The Ball"
    End If
End If

'same with above..
If s1 = p2 Then
'p2 tries steal from p1
    x1 = Rnd() * p2_stealing
    x2 = Rnd() * p1_dribling
    If x1 > x2 And distance(s1, s2) < min_dist_player_player_to_try_steal And x2 <= p2_pressing Then
        player_has_ball = 2
        p2.Top = p1.Top + 300
        what = p2_name + " Steals The Ball"
    End If
End If


End Function
Private Sub Timer2_Timer()
'player2 main artificial intelligence


'if player1(you) has the ball,then follow him
If player_has_ball = 1 Then
    nul = move_to(p2, p1)
End If

'if player1 has the ball and p2 close to him then
'player2 tries to steal the ball
If player_has_ball = 1 And distance(p1, p2) < 100 Then
    nul = try_steal(p2, p1)
End If

'if nobody has the ball,go catch her
If player_has_ball = 0 Then
    nul = move_to(p2, ball)
End If

'if nobody has the ball,and p2 close to her
'then p2 takes the ball position
If distance(p2, ball) < min_dist_player_ball And player_has_ball = 0 And ball_down = 1 Then
    player_has_ball = 2
    ball.Top = p2.Top - 50
    ball.Left = p2.Left
    what = p2_name + " Gets The Ball"
End If

'if player2 has the ball then
'function move_with_ball will handle
'player2 artficil intelligence
If player_has_ball = 2 Then
    nul = move_with_ball(p2)
End If

End Sub
Function move_with_ball(p As Shape)
'handles artificial intelligence of player 2
'when he has the ball.
Dim x As Integer
Randomize Timer

'approach the goal
If p2.Left > goal.Left + goal.Width / 2 Then
    p2.Left = p2.Left - p2_speed_with_ball
    ball.Top = p2.Top + p2.Height / 2 - ball.Height / 2
    ball.Left = p2.Left + p2.Width / 2 - ball.Width / 2 - Rnd() * 90
End If

'approach the goal
If p2.Left < goal.Left + goal.Width / 2 Then
    p2.Left = p2.Left + p2_speed_with_ball
    ball.Top = p2.Top + p2.Height / 2 - ball.Height / 2
    ball.Left = p2.Left + p2.Width / 2 - ball.Width / 2 + Rnd() * 90
End If

'approach the goal
If p2.Top > goal.Top + goal.Height Then
    p2.Top = p2.Top - p2_speed_with_ball
    ball.Top = p2.Top + p2.Height / 2 - ball.Height / 2 - Rnd() * 90
    ball.Left = p2.Left + p2.Width / 2 - ball.Width / 2
End If



'if p2 has ball and p1 is close to him then
'p2 must drible,to avoid p1 from stealing the ball

'part1:player2 dribles,and aproaches the goal
If distance(p2, p1) < min_dist_player_player_to_do_drible And player_has_ball = 2 And p1.Top < p2.Top And p2.Top > goal.Top + goal.Height + 500 Then
    ball.Top = p1.Top - p2_crossing * 10
    p2.Top = p1.Top + 200 - p2_crossing * 20
    ball.Left = ball.Left + Rnd() * p2_crossing * 40 - p2_crossing * 20
    player_has_ball = 0
End If

'part2:player2 dribles,and gets away from the goal
If distance(p2, p1) < min_dist_player_player_to_do_drible And player_has_ball = 2 And p1.Top < p2.Top And p2.Top < goal.Top + goal.Height + 500 Then
    ball.Top = p2.Top + p2_crossing * 10
    p2.Top = p1.Top - 200 + p2_crossing * 20
    ball.Left = ball.Left + Rnd() * p2_crossing * 40 - p2_crossing * 20
    player_has_ball = 0
End If
'=========

'player2 checks if he can shot
If p2.Top > goal.Top + goal.Height And check_shot() = 1 Then
    what = p2_name + " Shots!!"
    nul = shot(p2)
End If

'player2 is very close to the goal and shots immediately.
If p2.Top < goal.Top + goal.Height + 400 Then
    what = p2_name + " Shots!!"
    nul = shot(p2)
End If


Rem newwwwwwww
If distance(p2, gk) < 200 Then
    x = Rnd() * 10
    If x < 4 Then
        what = p2_name + " Shots!!"
        nul = shot(p2)
    End If
End If

End Function
Function check_shot()
'handles if player2 has good position
'to do a shoot.
'a player that tries long shots then
'1)shots from far distances more often
'and 2) tries shots from more far distances.


Dim x As Integer
check_shot = 0
x = Int(Rnd() * 100)


If x < 10 * (1 - distance(p2, goal) / (2000 + p2_tries_long_shots * 50)) And ((distance(p1, p2) > 400 * 20 / (p2_tries_long_shots + 1) Or p2.Top < p1.Top)) Then
    check_shot = 1
End If

End Function
Function shot(p As Shape)
'handles where ball will go after a shot.
'And if ball get in the goal then apdates score.
Randomize Timer
Dim xx As Integer
Dim min As Integer
Dim max As Integer

'xx calculates ball left.
'a player with good shooting has
'more chances to score.But distance
'between player and the goal is important..
'thats where contributor is usefull.
'So a good shooter sends ball into goal with many chances
'but if he is far away from the goal then chances
'redused highly.

'if shot is done by player1(you).
'so p=p1.
If p = p1 Then
    'xx= rnd()*(max-min)+min!!!
    min = goal.Left - goal.Width + 2 * goal.Width * p1_shoting / (2 * 20)
    max = goal.Left + goal.Width + goal.Width - 2 * goal.Width * p1_shoting / (2 * 20)
    'xx keeps where ball would go,if
    'distance was small
    xx = Rnd() * (max - min) + min
    'now,with function contributor we change xx range.
    'ball range changes,depended on distance
    'between player and goal.

    'contributor returns +1 or -1.
    xx = xx + (distance(p1, goal) / 4000) * goal.Width * contributor()
End If

'same things happened above with player 2.

'if shot is done by player2.
'so p=p2
If p = p2 Then
    'xx= rnd()*(max-min)+min!!!
    min = goal.Left - goal.Width + 2 * goal.Width * p2_shoting / (2 * 20)
    max = goal.Left + goal.Width + goal.Width - 2 * goal.Width * p2_shoting / (2 * 20)
    xx = Rnd() * (max - min) + min
    Rem new
    xx = xx + (distance(p2, goal) / 4000) * goal.Width * contributor()
End If


'ball wents where we shooted her.
ball.Left = xx
ball.Top = goal.Top + 50

'and finally we update score
'if ball went into goal
If xx > goal.Left And xx + ball.Width < goal.Left + goal.Width And p = p1 Then
    g1 = g1 + 1
End If
If xx > goal.Left And xx + ball.Width < goal.Left + goal.Width And p = p2 Then
    g2 = g2 + 1
End If

'here we stop program for 2 seconds..
Sleep (2000)

'and gk kicks the ball again
nul = gk_kick()
End Function
Function contributor()
'funtion returns 1 or -1.
'it works with function shot.
'so take a look in function shot
'to understand where is usefull
Randomize Timer
Dim w As Integer
w = Int(Rnd() * 2)

If w = 0 Then
contributor = 1
End If

If w = 1 Then
contributor = -1
End If

End Function
Function move_to(p2 As Shape, s As Shape)
'this is artificial intelligence of player2
'when he has not the ball..


'if neither player1(you) has the ball
'then go get the ball
If s = ball Then
    If p2.Left + p2.Width / 2 < s.Left + s.Width / 2 Then
        p2.Left = p2.Left + p2_speed
    End If
    If p2.Left + p2.Width / 2 > s.Left + s.Width / 2 Then
        p2.Left = p2.Left - p2_speed
    End If
    If p2.Top + p2.Height / 2 < s.Top + s.Height / 2 Then
        p2.Top = p2.Top + p2_speed
    End If
    If p2.Top + p2.Height / 2 > s.Top + s.Height / 2 Then
        p2.Top = p2.Top - p2_speed
    End If
End If

'if player1(you) has the ball then
'follow player1.
If s = p1 Then
    If p2.Left + p2.Width / 2 < s.Left + s.Width / 2 Then
        p2.Left = p2.Left + p2_speed
    End If
    If p2.Left + p2.Width / 2 > s.Left + s.Width / 2 Then
        p2.Left = p2.Left - p2_speed
    End If
    If p2.Top + p2.Height / 2 < s.Top + s.Height / 2 - Rnd() * 300 Then
        p2.Top = p2.Top + p2_speed
    End If
    If p2.Top + p2.Height / 2 > s.Top + s.Height / 2 - Rnd() * 300 Then
        p2.Top = p2.Top - p2_speed
    End If
End If
End Function

Private Sub Timer3_Timer()
'displays score
'g1:goals of player 1(you)
'g2:goals of player 2
sg1 = g1
sg2 = g2
score = "SCORE :" + sg1 + " - " + sg2
End Sub
Private Sub Timer4_Timer()


If player_has_ball = 0 Then
    cir.Visible = False
End If

If player_has_ball = 2 Then
    cir.Visible = True
    cir.Left = p2.Left + p2.Width / 2 - cir.Width / 2
    cir.Top = p2.Top + p2.Height / 2 - cir.Height / 2
End If

If player_has_ball = 1 Then
    cir.Visible = True
    cir.Left = p1.Left + p1.Width / 2 - cir.Width / 2
    cir.Top = p1.Top + p1.Height / 2 - cir.Height / 2
End If


End Sub

Private Sub Timer5_Timer()
kk = kk + 1
If kk = 5 Then
    kk = 1
End If

If kk = 1 Then
    cir.BorderColor = vbRed
End If

If kk = 2 Then
    cir.BorderColor = vbBlack
End If

If kk = 3 Then
    cir.BorderColor = vbRed
End If

If kk = 4 Then
    cir.BorderColor = vbBlack
End If

End Sub
