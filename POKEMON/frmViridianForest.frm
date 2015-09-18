VERSION 5.00
Begin VB.Form frmViridianForest 
   Caption         =   "Form2"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form2"
   ScaleHeight     =   8265
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8430
      Left            =   0
      Picture         =   "frmViridianForest.frx":0000
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8370
      Begin VB.PictureBox bug2Mask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         Picture         =   "frmViridianForest.frx":E2B42
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   240
      End
      Begin VB.PictureBox bug2Sprite 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Picture         =   "frmViridianForest.frx":E2EB2
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   960
         Width           =   240
      End
      Begin VB.PictureBox bug1Mask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7560
         Picture         =   "frmViridianForest.frx":E3284
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   4515
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox bug1Sprite 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6960
         Picture         =   "frmViridianForest.frx":E35F4
         ScaleHeight     =   285
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   4440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   360
         Left            =   3720
         Picture         =   "frmViridianForest.frx":E39C6
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   2
         Top             =   7680
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   435
         Picture         =   "frmViridianForest.frx":E3D2E
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   600
         Top             =   7200
      End
   End
End
Attribute VB_Name = "frmViridianForest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ashX, ashY As Integer
Dim boardPos(1 To 31, -1 To 32), direction, trainerType As String
Dim canWalk, loadMap, searching, movingTrainer, battled1, battled2 As Boolean
Dim trainerY, trainer1X, trainer1Y, trainer2X, trainer2Y, searchingX, searchingY As Integer
Public nextPokemon As String
Const SPEED = 17
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Randomize
    Select Case KeyCode
        Case vbKeyLeft
        picSprite.Picture = LoadPicture("oakLeft.gif")
        picMask.Picture = LoadPicture("oakLeftMask.gif")
        Draw_ViridianForest
            If Not boardPos((ashX + 6 - SPEED) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX + 6 - SPEED) / 17, (ashY + 10) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX - SPEED 'moves guy left
                        Draw_ViridianForest 'call the Draw_ViridianForest function
                    Else:
                        'random encounter!
                        'init values
                        frmBattle.CH = 1
                        frmBattle.BRN = 1
                        frmBattle.pAccuracy = 1
                        frmBattle.eAccuracy = 1
                        frmBattle.eAtkName1 = ""
                        frmBattle.eAtkName2 = ""
                        frmBattle.eAtkName3 = ""
                        frmBattle.eAtkName4 = ""
                        'randomly choose what pokemon will appear
                        rndPoke = Int((Rnd * 9) + 1)
                        Select Case rndPoke
                            Case 1
                                frmBattle.ePokeNum = 11
                            Case 2
                                frmBattle.ePokeNum = 11
                            Case 3
                                frmBattle.ePokeNum = 10
                            Case 4
                                frmBattle.ePokeNum = 10
                            Case 5
                                frmBattle.ePokeNum = 13
                            Case 6
                                frmBattle.ePokeNum = 13
                            Case 7
                                frmBattle.ePokeNum = 14
                            Case 8
                                frmBattle.ePokeNum = 14
                            Case 9
                                frmBattle.ePokeNum = 25
                        End Select
                        'generate enemy pokemon stats
                        frmBattle.pokeOwner = "e"
                        frmBattle.eLVL = Int(Rnd * 4) + 1
                        Call frmBattle.getEPokemon
                        Call frmBattle.getStats
                        'generate your pokemon's stats
                        frmBattle.pokeNum = 1
                        frmBattle.pokeOwner = "p"
                        frmBattle.arrayNum = 1
                        Call form1.GetLVL
                        frmBattle.passThis = form1.ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.arrayNum = 1
                        Call form1.GetHP
                        frmBattle.pHP = form1.ovwHP
                        frmBattle.lblpHP = form1.ovwHP
                        frmViridianForest.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                Else:
                    ashX = ashX - SPEED 'moves guy left
                    Draw_ViridianForest 'call the Draw_ViridianForest function
                End If
            End If
        Case vbKeyRight
        picSprite.Picture = LoadPicture("oakRight.gif")
        picMask.Picture = LoadPicture("oakRightMask.gif")
        Draw_ViridianForest
            If Not boardPos((ashX + 6 + SPEED) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX + 6 + SPEED) / 17, (ashY + 10) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX + SPEED 'moves guy right
                        Draw_ViridianForest 'call the Draw_ViridianForest function
                    Else:
                        'random encounter!
                        'init values
                        frmBattle.CH = 1
                        frmBattle.BRN = 1
                        frmBattle.pAccuracy = 1
                        frmBattle.eAccuracy = 1
                        frmBattle.eAtkName1 = ""
                        frmBattle.eAtkName2 = ""
                        frmBattle.eAtkName3 = ""
                        frmBattle.eAtkName4 = ""
                        'randomly choose what pokemon will appear
                        rndPoke = Int((Rnd * 9) + 1)
                        Select Case rndPoke
                            Case 1
                                frmBattle.ePokeNum = 11
                            Case 2
                                frmBattle.ePokeNum = 11
                            Case 3
                                frmBattle.ePokeNum = 10
                            Case 4
                                frmBattle.ePokeNum = 10
                            Case 5
                                frmBattle.ePokeNum = 13
                            Case 6
                                frmBattle.ePokeNum = 13
                            Case 7
                                frmBattle.ePokeNum = 14
                            Case 8
                                frmBattle.ePokeNum = 14
                            Case 9
                                frmBattle.ePokeNum = 25
                        End Select
                        'generate enemy pokemon stats
                        frmBattle.pokeOwner = "e"
                        frmBattle.eLVL = Int(Rnd * 4) + 1
                        Call frmBattle.getEPokemon
                        Call frmBattle.getStats
                         'generate your pokemon's stats
                        frmBattle.pokeNum = 1
                        frmBattle.pokeOwner = "p"
                        frmBattle.arrayNum = 1
                        Call form1.GetLVL
                        frmBattle.passThis = form1.ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.arrayNum = 1
                        Call form1.GetHP
                        frmBattle.pHP = form1.ovwHP
                        frmBattle.lblpHP = form1.ovwHP
                        frmViridianForest.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                Else:
                    ashX = ashX + SPEED 'moves guy right
                    Draw_ViridianForest 'call the Draw_ViridianForest function
                End If
            End If
        Case vbKeyUp
        picSprite.Picture = LoadPicture("oakBack.gif")
        picMask.Picture = LoadPicture("oakBackMask.gif")
        Draw_ViridianForest
            If Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "barrier" And Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "" Or boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "grass" Or boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "lowerBound" Or boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "upperBound" Then
                    If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                        If (Rnd * 100) + 1 > 6 Then
                            ashY = ashY - SPEED 'moves guy up
                            Draw_ViridianForest 'call the Draw_ViridianForest function
                        Else:
                            'random encounter!
                            'init values
                            frmBattle.CH = 1
                            frmBattle.BRN = 1
                            frmBattle.pAccuracy = 1
                            frmBattle.eAccuracy = 1
                            frmBattle.eAtkName1 = ""
                            frmBattle.eAtkName2 = ""
                            frmBattle.eAtkName3 = ""
                            frmBattle.eAtkName4 = ""
                            'randomly choose what pokemon will appear
                            rndPoke = Int((Rnd * 9) + 1)
                            Select Case rndPoke
                                Case 1
                                    frmBattle.ePokeNum = 11
                                Case 2
                                    frmBattle.ePokeNum = 11
                                Case 3
                                    frmBattle.ePokeNum = 10
                                Case 4
                                    frmBattle.ePokeNum = 10
                                Case 5
                                    frmBattle.ePokeNum = 13
                                Case 6
                                    frmBattle.ePokeNum = 13
                                Case 7
                                    frmBattle.ePokeNum = 14
                                Case 8
                                    frmBattle.ePokeNum = 14
                                Case 9
                                    frmBattle.ePokeNum = 25
                            End Select
                            'generate enemy pokemon stats
                            frmBattle.pokeOwner = "e"
                            frmBattle.eLVL = Int(Rnd * 4) + 1
                            Call frmBattle.getEPokemon
                            Call frmBattle.getStats
                             'generate your pokemon's stats
                            frmBattle.pokeNum = 1
                            frmBattle.pokeOwner = "p"
                            frmBattle.arrayNum = 1
                            Call form1.GetLVL
                            frmBattle.passThis = form1.ovwLVL
                            Call frmBattle.passpLVL
                            Call frmBattle.getPPokemon
                            Call frmBattle.getStats
                            frmBattle.arrayNum = 1
                            Call form1.GetHP
                            frmBattle.pHP = form1.ovwHP
                            frmBattle.lblpHP = form1.ovwHP
                            frmViridianForest.Hide
                            frmTxtBox.List1.Clear
                            frmTxtBox.Hide
                            frmTxtBox.List1.Clear
                            frmTxtBox.Hide
                            frmBattle.Show
                        End If
                    ElseIf boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "upperBound" Then
                        frmViridianForest.Hide
                        frmPewterCity.Show
                        frmTxtBox.List1.Clear
                        frmTxtBox.List1.AddItem "Pewter City"
                        frmPewterCity.ashX = 266
                        frmPewterCity.ashY = 517
                        form1.MAP = "PEWTERCITY"
                    Else:
                        ashY = ashY - SPEED 'moves guy up
                        Draw_ViridianForest 'call the Draw_ViridianForest function
                    End If
                Else:
                    ashY = ashY - SPEED
                    Draw_ViridianForest
                    canWalk = False
                    trainerType = boardPos((ashX + 6) / 17, (ashY + 10) / 17)
                    Select Case trainerType
                    Case "BUGCATCHER1"
                        'have you battled this trainer already?
                        If battled1 = False Then
                            'move the enemy trainer to you
                            direction = "LEFT"
                            movingTrainer = True
                        Else:
                            canWalk = True
                        End If
                    Case "BUGCATCHER2"
                        'have you battled this trainer already?
                        If battled2 = False Then
                            'move the enemy trainer towards you
                            direction = "LEFT"
                            movingTrainer = True
                        Else:
                            canWalk = True
                        End If
                    End Select
                End If
            End If
        Case vbKeyDown
        picSprite.Picture = LoadPicture("oakFront.gif")
        picMask.Picture = LoadPicture("oakFrontMask.gif")
        Draw_ViridianForest
            If Not boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "barrier" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "lowerBound" Then
                    form1.MAP = "VIRIDIANCITY"
                    frmTxtBox.List1.Clear
                    frmTxtBox.List1.AddItem "Viridian City"
                    frmViridianCity.Show
                    frmViridianForest.Hide
                    frmViridianCity.Draw_ViridianCity
                    frmViridianForest.ashY = 0
                ElseIf boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "cliff" Then
                    ashY = ashY + SPEED + SPEED 'moves guy down 2
                    Draw_ViridianForest 'call the Draw_ViridianForest function
                ElseIf boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "" Or boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "upperBound" Or boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "grass" Then
                    If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                        If (Rnd * 100) + 1 > 6 Then
                            ashY = ashY + SPEED 'moves guy down
                            Draw_ViridianForest 'call the Draw_ViridianForest function
                        Else:
                            'random encounter!
                            'init values
                            frmBattle.CH = 1
                            frmBattle.BRN = 1
                            frmBattle.pAccuracy = 1
                            frmBattle.eAccuracy = 1
                            frmBattle.eAtkName1 = ""
                            frmBattle.eAtkName2 = ""
                            frmBattle.eAtkName3 = ""
                            frmBattle.eAtkName4 = ""
                        'randomly choose what pokemon will appear
                        rndPoke = Int((Rnd * 9) + 1)
                        Select Case rndPoke
                            Case 1
                                frmBattle.ePokeNum = 11
                            Case 2
                                frmBattle.ePokeNum = 11
                            Case 3
                                frmBattle.ePokeNum = 10
                            Case 4
                                frmBattle.ePokeNum = 10
                            Case 5
                                frmBattle.ePokeNum = 13
                            Case 6
                                frmBattle.ePokeNum = 13
                            Case 7
                                frmBattle.ePokeNum = 14
                            Case 8
                                frmBattle.ePokeNum = 14
                            Case 9
                                frmBattle.ePokeNum = 25
                        End Select
                            'generate enemy pokemon stats
                            frmBattle.pokeOwner = "e"
                            frmBattle.eLVL = Int(Rnd * 4) + 1
                            Call frmBattle.getEPokemon
                            Call frmBattle.getStats
                            'generate your pokemon's stats
                            frmBattle.pokeNum = 1
                            frmBattle.pokeOwner = "p"
                            frmBattle.arrayNum = 1
                            Call form1.GetLVL
                            frmBattle.passThis = form1.ovwLVL
                            Call frmBattle.passpLVL
                            Call frmBattle.getPPokemon
                            Call frmBattle.getStats
                            frmBattle.arrayNum = 1
                            Call form1.GetHP
                            frmBattle.pHP = form1.ovwHP
                            frmBattle.lblpHP = form1.ovwHP
                            frmViridianForest.Hide
                            frmTxtBox.List1.Clear
                            frmTxtBox.Hide
                            frmBattle.Show
                        End If
                    Else:
                        ashY = ashY + SPEED 'moves guy up
                        Draw_ViridianForest 'call the Draw_ViridianForest function
                    End If
                Else:
                    'move the player up
                    ashY = ashY + SPEED
                    Draw_ViridianForest
                    'freeze the player
                    canWalk = False
                    trainerType = boardPos((ashX + 6) / 17, (ashY + 10) / 17)
                    Select Case trainerType
                    Case "BUGCATCHER1"
                        'have you battled this trainer already?
                        If battled1 = False Then
                            'move the enemy trainer to you
                            direction = "LEFT"
                            movingTrainer = True
                        Else:
                            canWalk = True
                        End If
                    Case "BUGCATCHER2"
                        'have you battled this trainer already?
                        If battled2 = False Then
                            'move the enemy trainer towards you
                            direction = "LEFT"
                            movingTrainer = True
                        Else:
                            canWalk = True
                        End If
                    End Select
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    'you haven't battled anyone yet
    battled1 = False
    battled2 = False
    'you can walk
    canWalk = True
    loadMap = False
    'put the screen in the top left corner of the window
    frmViridianForest.Left = 0
    frmViridianForest.Top = 0
    'initialize boundary tiles
    For x = 3 To 30
        boardPos(x, 31) = "lowerBound"
    Next x
    For x = 2 To 3
        boardPos(x, 0) = "upperBound"
    Next x
    'initialize collisions
    For y = 0 To 31
        boardPos(1, y) = "barrier"
        boardPos(31, y) = "barrier"
    Next y
    For x = 4 To 31
        boardPos(x, 0) = "barrier"
    Next x
    For x = 1 To 12
        For y = 29 To 31
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 18 To 31
        For y = 29 To 31
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 1 To 7
        For y = 17 To 24
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 1 To 12
        For y = 11 To 14
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 4 To 6
        For y = 1 To 6
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 11 To 16
         For y = 19 To 24
            boardPos(x, y) = "barrier"
         Next y
    Next x
    For x = 11 To 25
        For y = 15 To 16
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 15 To 25
        For y = 7 To 16
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 21 To 25
        For y = 17 To 24
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For y = 19 To 24
        boardPos(20, y) = "barrier"
    Next y
    For x = 11 To 12
        For y = 3 To 8
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 13 To 14
        For y = 7 To 8
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 28 To 30
        For y = 19 To 24
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For y = y To 16
        boardPos(29, y) = "barrier"
    Next y
    For x = 18 To 19
        For y = 26 To 27
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 16 To 19
    For a = 22 To 29
        boardPos(x, 4) = "barrier"
        boardPos(x, 5) = "barrier"
        boardPos(a, 4) = "barrier"
        boardPos(a, 5) = "barrier"
    Next a
    Next x
    For x = 15 To 30
        boardPos(x, 1) = "barrier"
        boardPos(x, 2) = "barrier"
    Next x
    For y = 7 To 16
        boardPos(29, y) = "barrier"
    Next y
    'initialize grass tiles
    For x = 2 To 7
        For y = 25 To 28
            boardPos(x, y) = "grass"
        Next y
    Next x
    For x = 2 To 10
        For y = 15 To 16
            boardPos(x, y) = "grass"
        Next y
    Next x
    For y = 17 To 26
        boardPos(10, y) = "grass"
        boardPos(17, y) = "grass"
        boardPos(26, y) = "grass"
    Next y
    For x = 11 To 16
        boardPos(x, 17) = "grass"
        boardPos(x, 18) = "grass"
        boardPos(x, 25) = "grass"
        boardPos(x, 26) = "grass"
    Next x
    For x = 20 To 25
        boardPos(x, 25) = "grass"
        boardPos(x, 26) = "grass"
    Next x
    For x = 28 To 30
        For y = 25 To 28
            boardPos(x, y) = "grass"
        Next y
    Next x
    
    'define locations of trainers
    searching = True
    trainer1X = 488
    trainer1Y = 301
    trainerY = trainer1Y
    searchingX = trainer1X
    searchingY = trainer1Y
    trainerType = "BUGCATCHER1"
    direction = "LEFT"
    Call findTrainerRange
    
    searching = True
    trainer2X = 46
    trainer2Y = 64
    trainerY = trainer2Y
    searchingX = trainer2X
    searchingY = trainer2Y
    trainerType = "BUGCATCHER2"
    direction = "LEFT"
    Call findTrainerRange
    
    boardPos((trainer1X + 6) / 17, (trainer1Y + 10) / 17) = "barrier"
    boardPos((trainer2X + 6) / 17, (trainer2Y + 10) / 17) = "barrier"
    End Sub
    
Private Sub findTrainerRange()
    Do While searching = True
        Select Case direction
        Case "LEFT"
            If Not boardPos((searchingX + 6 - (SPEED * i)) / 17, (trainerY + 10) / 17) = "barrier" Then
                boardPos((searchingX + 6 - (SPEED * i)) / 17, (trainerY + 10) / 17) = trainerType
            Else:
                searching = False
            End If
        Case "RIGHT"
            If Not boardPos((searchingX + 6 + (SPEED * i)) / 17, (trainerY + 10) / 17) = "barrier" Then
                boardPos((searchingX + 6 + (SPEED * i)) / 17, (trainerY + 10) / 17) = trainerType
            Else:
                searching = False
            End If
        Case "UP"
        Case "DOWN"
        End Select
        i = i + 1
        If i > 6 Then
            searching = False
        End If
    Loop
    searching = True
End Sub

Public Sub Draw_ViridianForest()
    BitBlt frmViridianForest.hDC, 0, 0, picBack.Width, picBack.Height, picBack.hDC, 0, 0, vbSrcCopy 'bitblt background onto form
    BitBlt frmViridianForest.hDC, ashX, ashY, picMask.Width, picMask.Height, picMask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
    BitBlt frmViridianForest.hDC, ashX, ashY, picSprite.Width, picSprite.Height, picSprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
    BitBlt frmViridianForest.hDC, trainer2X, trainer2Y, bug2Mask.Width, bug2Mask.Height, bug2Mask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
    BitBlt frmViridianForest.hDC, trainer2X, trainer2Y, bug2Sprite.Width, bug2Sprite.Height, bug2Sprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
    BitBlt frmViridianForest.hDC, trainer1X, trainer1Y, bug1Mask.Width, bug1Mask.Height, bug1Mask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
    BitBlt frmViridianForest.hDC, trainer1X, trainer1Y, bug1Sprite.Width, bug1Sprite.Height, bug1Sprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
End Sub

Public Sub get_Battle()
    Select Case nextPokemon
        Case ""
            canWalk = True
            Call Draw_ViridianForest
        Case "weedle6"
            'weedle
            frmBattle.List1.AddItem "Bugcatcher sends out WEEDLE!"
            frmBattle.ePokeNum = 13
            'choose its level
            frmBattle.eLVL = 6
            'ini frmBattle trainer battle
            frmBattle.tEXP = 1.5
            frmBattle.canCatch = False
            frmBattle.cmdAtk1.Caption = "Start"
            frmBattle.cmdAtk2.Caption = ""
            frmBattle.cmdAtk3.Caption = ""
            frmBattle.cmdAtk4.Caption = ""
            frmBattle.trainerBattle.Visible = True
            'generate enemy pokemon stats
            frmBattle.pokeOwner = "e"
            Call frmBattle.getEPokemon
            Call frmBattle.getStats
            'generate your pokemon's stats
            frmBattle.pokeNum = 1
            frmBattle.pokeOwner = "p"
            frmBattle.arrayNum = 1
            Call form1.GetLVL
            frmBattle.passThis = form1.ovwLVL
            Call frmBattle.passpLVL
            Call frmBattle.getPPokemon
            Call frmBattle.getStats
            frmBattle.arrayNum = 1
            Call form1.GetHP
            frmBattle.pHP = form1.ovwHP
            frmBattle.lblpHP = form1.ovwHP
            frmViridianForest.Hide
            frmTxtBox.List1.Clear
            frmTxtBox.Hide
            frmBattle.Show
            nextPokemon = ""
            canWalk = True
            frmTxtBox.List1.AddItem "Bugcatcher defeated!"
    End Select
End Sub

Private Sub Timer1_Timer()
    If loadMap = False Then
        Draw_ViridianForest
        loadMap = True
    End If
    If movingTrainer = True Then
        Select Case direction
        Case "LEFT"
            Select Case trainerType
            Case "BUGCATCHER1"
                If trainer1X - SPEED = ashX - 1 Then
                    boardPos((trainer1X + 6) / 17, (trainer1Y + 10) / 17) = "barrier"
                    movingTrainer = False
                    'start the battle
                    'pick their first pokemon
                    'weedle
                    frmBattle.ePokeNum = 10
                    frmBattle.List1.AddItem "Bugcatcher sends out CATERPIE!"
                    'choose its level
                    frmBattle.eLVL = 6
                    'ini frmBattle trainer battle
                    frmBattle.tEXP = 1.5
                    frmBattle.canCatch = False
                    frmBattle.cmdAtk1.Caption = "Start"
                    frmBattle.cmdAtk2.Caption = ""
                    frmBattle.cmdAtk3.Caption = ""
                    frmBattle.cmdAtk4.Caption = ""
                    frmBattle.trainerBattle.Visible = True
                    'generate enemy pokemon stats
                    frmBattle.pokeOwner = "e"
                    Call frmBattle.getEPokemon
                    Call frmBattle.getStats
                    'generate your pokemon's stats
                    frmBattle.pokeNum = 1
                    frmBattle.pokeOwner = "p"
                    frmBattle.arrayNum = 1
                    Call form1.GetLVL
                    frmBattle.passThis = form1.ovwLVL
                    Call frmBattle.passpLVL
                    Call frmBattle.getPPokemon
                    Call frmBattle.getStats
                    frmBattle.arrayNum = 1
                    Call form1.GetHP
                    frmBattle.pHP = form1.ovwHP
                    frmBattle.lblpHP = form1.ovwHP
                    frmViridianForest.Hide
                    frmTxtBox.List1.Clear
                    frmTxtBox.Hide
                    frmBattle.Show
                    'choose their next pokemon
                    nextPokemon = "weedle6"
                    battled1 = True
                Else:
                    trainer1X = trainer1X - SPEED
                    boardPos((trainer1X + 6) / 17, (trainer1Y + 10) / 17) = ""
                    Call Draw_ViridianForest
                End If
            Case "BUGCATCHER2"
                If trainer2X - SPEED = ashX - 1 Then
                    boardPos((trainer2X + 6) / 17, (trainer2Y + 10) / 17) = "barrier"
                    movingTrainer = False
                    canWalk = True
                    'start the battle
                    'pick their first pokemon
                    'weedle
                    frmBattle.ePokeNum = 13
                    frmBattle.List1.AddItem "Bugcatcher sends out WEEDLE!"
                    'choose its level
                    frmBattle.eLVL = 9
                    'ini frmBattle trainer battle
                    frmBattle.tEXP = 1.5
                    frmBattle.canCatch = False
                    frmBattle.cmdAtk1.Caption = "Start"
                    frmBattle.cmdAtk2.Caption = ""
                    frmBattle.cmdAtk3.Caption = ""
                    frmBattle.cmdAtk4.Caption = ""
                    frmBattle.trainerBattle.Visible = True
                    'generate enemy pokemon stats
                    frmBattle.pokeOwner = "e"
                    Call frmBattle.getEPokemon
                    Call frmBattle.getStats
                    'generate your pokemon's stats
                    frmBattle.pokeNum = 1
                    frmBattle.pokeOwner = "p"
                    frmBattle.arrayNum = 1
                    Call form1.GetLVL
                    frmBattle.passThis = form1.ovwLVL
                    Call frmBattle.passpLVL
                    Call frmBattle.getPPokemon
                    Call frmBattle.getStats
                    frmBattle.arrayNum = 1
                    Call form1.GetHP
                    frmBattle.pHP = form1.ovwHP
                    frmBattle.lblpHP = form1.ovwHP
                    frmViridianForest.Hide
                    frmTxtBox.List1.Clear
                    frmTxtBox.Hide
                    frmBattle.Show
                    battled2 = True
                Else:
                    trainer2X = trainer2X - SPEED
                    boardPos((trainer2X + 6) / 17, (trainer2Y + 10) / 17) = ""
                    Call Draw_ViridianForest
                End If
            End Select
        Case "RIGHT"
            Select Case trainerType
            Case "BUGCATCHER1"
                If trainer1X + SPEED = ashX - 1 Then
                    movingTrainer = False
                    canWalk = True
                    'start the battle
                Else:
                    trainer1X = trainer1X + SPEED
                    Call Draw_ViridianForest
                End If
            Case "BUGCATCHER2"
                If trainer2X + SPEED = ashX - 1 Then
                    movingTrainer = False
                    canWalk = True
                    'start the battle
                Else:
                    trainer2X = trainer2X + SPEED
                    Call Draw_ViridianForest
                End If
            End Select
        Case "UP"
        Case "DOWN"
        End Select
        
    End If
    frmViridianForest.Left = 0
    frmViridianForest.Top = 0
End Sub


