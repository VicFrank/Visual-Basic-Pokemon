VERSION 5.00
Begin VB.Form frmViridianCity 
   Caption         =   "frmViridianCity"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form2"
   ScaleHeight     =   8295
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8430
      Left            =   0
      Picture         =   "frmViridianCity.frx":0000
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8370
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   600
         Top             =   7200
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4770
         Picture         =   "frmViridianCity.frx":E2B42
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   2
         Top             =   5760
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   360
         Left            =   3720
         Picture         =   "frmViridianCity.frx":E2F13
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   1
         Top             =   7680
         Visible         =   0   'False
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmViridianCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ashX, ashY As Integer
Dim boardPos(1 To 31, 1 To 32), person As String
Dim canWalk, loadMap, canTalk As Boolean
Public passThis, arrayNum As Integer
Const SPEED = 17
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Randomize
    Select Case KeyCode
        Case vbKeyLeft
        picSprite.Picture = LoadPicture("oakLeft.gif")
        picMask.Picture = LoadPicture("oakLeftMask.gif")
        Draw_ViridianCity
            If Not boardPos((ashX - SPEED) / 17, (ashY + 5) / 17) = "barrier" And Not boardPos((ashX - SPEED) / 17, (ashY + 5) / 17) = "cliff" And Not boardPos((ashX - SPEED) / 17, (ashY + 5) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX) / 17, (ashY + 5) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX - SPEED 'moves guy left
                        Draw_ViridianCity 'call the Draw_ViridianCity function
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
                        rndPoke = Int((Rnd * 2) + 1)
                        Select Case rndPoke
                            Case 1
                                frmBattle.ePokeNum = 16
                            Case 2
                                frmBattle.ePokeNum = 18
                        End Select
                        'generate enemy pokemon stats
                        frmBattle.pokeOwner = "e"
                        frmBattle.eLVL = Int(Rnd * 4) + 1
                        Call frmBattle.getEPokemon
                        Call frmBattle.getStats
                        'generate your pokemon's stats
                        frmBattle.pokeNum = 1
                        frmBattle.pokeOwner = "p"
                        arrayNum = 1
                        Call form1.GetLVL
                        passThis = ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        Call form1.GetHP
                        frmBattle.pHP = ovwHP
                        frmRoute1.Hide
                        frmBattle.Show
                    End If
                Else:
                    ashX = ashX - SPEED 'moves guy left
                    Draw_ViridianCity 'call the Draw_ViridianCity function
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX + 4 - SPEED) / 17, (ashY + 10) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4 - SPEED) / 17) = 11 And ((ashY + 10) / 17) = 20 Then
                    person = "HOUSE1PERSON"
                ElseIf ((ashX + 4 - SPEED) / 17) = 19 And ((ashY + 10) / 17) = 13 Then
                    person = "HOUSE2PERSON"
                ElseIf ((ashX + 4 - SPEED) / 17) = 13 And ((ashY + 10) / 17) = 2 Then
                    person = "OLDMAN"
                ElseIf ((ashX + 4 - SPEED) / 17) = 13 And ((ashY + 10) / 17) = 3 Then
                    person = "PIKACHU"
                ElseIf ((ashX + 4 - SPEED) / 17) = 25 And ((ashY + 10) / 17) = 22 Then
                    person = "GARDENER"
                ElseIf ((ashX + 4 - SPEED) / 17) = 25 And ((ashY + 10) / 17) = 6 Then
                    person = "GYMDUDE"
                ElseIf ((ashX + 4 - SPEED) / 17) = 24 And ((ashY + 10) / 17) = 6 Then
                    person = "GYMSIGN"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyRight
        picSprite.Picture = LoadPicture("oakRight.gif")
        picMask.Picture = LoadPicture("oakRightMask.gif")
        Draw_ViridianCity
            If Not boardPos((ashX + SPEED) / 17, (ashY + 5) / 17) = "barrier" And Not boardPos((ashX + SPEED) / 17, (ashY + 5) / 17) = "cliff" And Not boardPos((ashX + SPEED) / 17, (ashY + 5) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX) / 17, (ashY + 5) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX + SPEED 'moves guy right
                        Draw_ViridianCity 'call the Draw_ViridianCity function
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
                        rndPoke = Int((Rnd * 2) + 1)
                        Select Case rndPoke
                            Case 1
                                frmBattle.ePokeNum = 16
                            Case 2
                                frmBattle.ePokeNum = 18
                        End Select
                        'generate enemy pokemon stats
                        frmBattle.pokeOwner = "e"
                        frmBattle.eLVL = Int(Rnd * 4) + 1
                        Call frmBattle.getEPokemon
                        Call frmBattle.getStats
                        'generate your pokemon's stats
                        frmBattle.pokeNum = 1
                        frmBattle.pokeOwner = "p"
                        arrayNum = 1
                        Call form1.GetLVL
                        passThis = ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        Call form1.GetHP
                        frmBattle.pHP = ovwHP
                        frmRoute1.Hide
                        frmBattle.Show
                    End If
                    Else:
                        ashX = ashX + SPEED 'moves guy right
                        Draw_ViridianCity 'call the Draw_ViridianCity function
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX + 4 + SPEED) / 17, (ashY + 10) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4 + SPEED) / 17) = 11 And ((ashY + 10) / 17) = 20 Then
                    person = "HOUSE1PERSON"
                ElseIf ((ashX + 4 + SPEED) / 17) = 19 And ((ashY + 10) / 17) = 13 Then
                    person = "HOUSE2PERSON"
                ElseIf ((ashX + 4 + SPEED) / 17) = 13 And ((ashY + 10) / 17) = 2 Then
                    person = "OLDMAN"
                ElseIf ((ashX + 4 + SPEED) / 17) = 13 And ((ashY + 10) / 17) = 3 Then
                    person = "PIKACHU"
                ElseIf ((ashX + 4 + SPEED) / 17) = 25 And ((ashY + 10) / 17) = 22 Then
                    person = "GARDENER"
                ElseIf ((ashX + 4 + SPEED) / 17) = 25 And ((ashY + 10) / 17) = 6 Then
                    person = "GYMDUDE"
                ElseIf ((ashX + 4 + SPEED) / 17) = 24 And ((ashY + 10) / 17) = 6 Then
                    person = "GYMSIGN"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyUp
        picSprite.Picture = LoadPicture("oakBack.gif")
        picMask.Picture = LoadPicture("oakBackMask.gif")
        Draw_ViridianCity
            If Not boardPos((ashX) / 17, (ashY + 5 - SPEED) / 17) = "barrier" And Not boardPos((ashX) / 17, (ashY + 5 - SPEED) / 17) = "cliff" And Not boardPos((ashX) / 17, (ashY + 5 - SPEED) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX) / 17, ashY / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY - SPEED 'moves guy up
                        Draw_ViridianCity 'call the Draw_ViridianCity function
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
                        rndPoke = Int((Rnd * 2) + 1)
                        Select Case rndPoke
                            Case 1
                                frmBattle.ePokeNum = 16
                            Case 2
                                frmBattle.ePokeNum = 18
                        End Select
                        'generate enemy pokemon stats
                        frmBattle.pokeOwner = "e"
                        frmBattle.eLVL = Int(Rnd * 4) + 1
                        Call frmBattle.getEPokemon
                        Call frmBattle.getStats
                       'generate your pokemon's stats
                        frmBattle.pokeNum = 1
                        frmBattle.pokeOwner = "p"
                        arrayNum = 1
                        Call form1.GetLVL
                        passThis = ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        Call form1.GetHP
                        frmBattle.pHP = ovwHP
                        frmRoute1.Hide
                        frmBattle.Show
                    End If
                ElseIf boardPos((ashX) / 17, (ashY + 5) / 17) = "upperBound" Then
                    frmViridianCity.Hide
                    frmViridianForest.Show
                    form1.MAP = "VIRIDIANFOREST"
                    frmTxtBox.List1.Clear
                    frmTxtBox.List1.AddItem "Viridian Forest"
                    frmViridianForest.ashX = ashX
                    frmViridianForest.ashY = 500
                ElseIf boardPos((ashX) / 17, (ashY + 5 - SPEED) / 17) = "pkmnCenter" Then
                    frmViridianCity.Hide
                    frmPkmnCenter.Show
                    frmPkmnCenter.ashX = 102
                    frmPkmnCenter.ashY = 165
                    frmPkmnCenter.Draw_pkmnCenter
                Else:
                    ashY = ashY - SPEED 'moves guy up
                    Draw_ViridianCity 'call the Draw_ViridianCity function
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX + 4) / 17, (ashY + 10 - SPEED) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4) / 17) = 11 And ((ashY + 10 - SPEED) / 17) = 20 Then
                    person = "HOUSE1PERSON"
                ElseIf ((ashX + 4) / 17) = 19 And ((ashY + 10 - SPEED) / 17) = 13 Then
                    person = "HOUSE2PERSON"
                ElseIf ((ashX + 4) / 17) = 13 And ((ashY + 10 - SPEED) / 17) = 2 Then
                    person = "OLDMAN"
                ElseIf ((ashX + 4) / 17) = 13 And ((ashY + 10 - SPEED) / 17) = 3 Then
                    person = "PIKACHU"
                ElseIf ((ashX + 4) / 17) = 25 And ((ashY + 10 - SPEED) / 17) = 22 Then
                    person = "GARDENER"
                ElseIf ((ashX + 4) / 17) = 25 And ((ashY + 10 - SPEED) / 17) = 6 Then
                    person = "GYMDUDE"
                ElseIf ((ashX + 4) / 17) = 24 And ((ashY + 10 - SPEED) / 17) = 6 Then
                    person = "GYMSIGN"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyDown
        picSprite.Picture = LoadPicture("oakFront.gif")
        picMask.Picture = LoadPicture("oakFrontMask.gif")
        Draw_ViridianCity
            If Not boardPos((ashX) / 17, (ashY + 5 + SPEED) / 17) = "barrier" And Not boardPos((ashX) / 17, (ashY + 5 + SPEED) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX) / 17, (ashY + 5 + SPEED) / 17) = "lowerBound" Then
                    form1.MAP = "ROUTE1"
                    frmRoute1.Show
                    frmViridianCity.Hide
                    frmRoute1.Draw_Route1
                    frmRoute1.ashX = ashX
                    frmViridianCity.ashY = 0
                ElseIf boardPos((ashX) / 17, (ashY + 5 + SPEED) / 17) = "cliff" Then
                    ashY = ashY + SPEED + SPEED 'moves guy down 2
                    Draw_ViridianCity 'call the Draw_ViridianCity function
                Else:
                If boardPos((ashX) / 17, (ashY + 5) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY + SPEED 'moves guy down
                        Draw_ViridianCity 'call the Draw_ViridianCity function
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
                        rndPoke = Int((Rnd * 2) + 1)
                        Select Case rndPoke
                            Case 1
                                frmBattle.ePokeNum = 16
                            Case 2
                                frmBattle.ePokeNum = 18
                        End Select
                        'generate enemy pokemon stats
                        frmBattle.pokeOwner = "e"
                        frmBattle.eLVL = Int(Rnd * 4) + 1
                        Call frmBattle.getEPokemon
                        Call frmBattle.getStats
                        'generate your pokemon's stats
                        frmBattle.pokeNum = 1
                        frmBattle.pokeOwner = "p"
                        arrayNum = 1
                        Call form1.GetLVL
                        passThis = ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        Call form1.GetHP
                        frmBattle.pHP = ovwHP
                        frmRoute1.Hide
                        frmBattle.Show
                    End If
                    Else:
                        ashY = ashY + SPEED 'moves guy up
                        Draw_ViridianCity 'call the Draw_ViridianCity function
                End If
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX + 4) / 17, (ashY + 10 + SPEED) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4) / 17) = 11 And ((ashY + 10 + SPEED) / 17) = 20 Then
                    person = "HOUSE1PERSON"
                ElseIf ((ashX + 4) / 17) = 19 And ((ashY + 10 + SPEED) / 17) = 13 Then
                    person = "HOUSE2PERSON"
                ElseIf ((ashX + 4) / 17) = 13 And ((ashY + 10 + SPEED) / 17) = 2 Then
                    person = "OLDMAN"
                ElseIf ((ashX + 4) / 17) = 13 And ((ashY + 10 + SPEED) / 17) = 3 Then
                    person = "PIKACHU"
                ElseIf ((ashX + 4) / 17) = 25 And ((ashY + 10 + SPEED) / 17) = 22 Then
                    person = "GARDENER"
                ElseIf ((ashX + 4) / 17) = 25 And ((ashY + 10 + SPEED) / 17) = 6 Then
                    person = "GYMDUDE"
                ElseIf ((ashX + 4) / 17) = 24 And ((ashY + 10 + SPEED) / 17) = 6 Then
                    person = "GYMSIGN"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyReturn
            If canTalk = True Then
                frmTxtBox.List1.Clear
                Select Case person
                    Case "HOUSE1PERSON"
                        frmTxtBox.List1.AddItem "... ... ... ... ... ... .."
                    Case "HOUSE2PERSON"
                        frmTxtBox.List1.AddItem "Making houses is hard!"
                    Case "OLDMAN"
                        frmTxtBox.List1.AddItem "I used to teach kids how to catch"
                        frmTxtBox.List1.AddItem "pokemon, but they all disappeared"
                        frmTxtBox.List1.AddItem "around Cinnabar Island."
                    Case "PIKACHU"
                        frmTxtBox.List1.AddItem "Pika, Pika!"
                    Case "GARDENER"
                        frmTxtBox.List1.AddItem "I'm a gardener, but for some reason"
                        frmTxtBox.List1.AddItem "there are no flowers in this game."
                    Case "GYMDUDE"
                        frmTxtBox.List1.AddItem "Hey there, champ in the making!"
                        frmTxtBox.List1.AddItem "This city is looking for a new gym leader."
                    Case "GYMSIGN"
                        frmTxtBox.List1.AddItem "Viridian City Gym:"
                        frmTxtBox.List1.AddItem "Looking for a new leader!"
                End Select
                canTalk = False
                canWalk = False
            Else:
                frmTxtBox.List1.Clear
                canWalk = True
            End If
    End Select
End Sub

Private Sub Form_Load()
    frmViridianCity.Left = 0
    frmViridianCity.Top = 0
    canWalk = True
    loadMap = False
    'initialize boundary tiles
    For x = 3 To 25
        boardPos(x, 32) = "lowerBound"
    Next x
    For x = 3 To 25
        boardPos(x, 2) = "upperBound"
    Next x
    'initialzie collisions
    'pkmnCenter
    For x = 17 To 21
        For y = 18 To 22
            boardPos(x, y) = "barrier"
        Next y
    Next x
    boardPos(19, 22) = "pkmnCenter"
    'house1
    For x = 10 To 13
        For y = 16 To 19
            boardPos(x, y) = "barrier"
        Next y
    Next x
    'house2
    For x = 18 To 21
        For y = 9 To 12
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 1 To 14
        For y = 29 To 32
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 18 To 31
        For y = 29 To 32
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For y = 1 To 32
        boardPos(1, y) = "barrier"
        boardPos(31, y) = "barrier"
    Next y
    For x = 1 To 12
        For y = 1 To 3
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 18 To 30
        For y = 1 To 5
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 6 To 12
        For y = 5 To 12
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 1 To 3
        For y = 4 To 12
            boardPos(x, y) = "barrier"
        Next y
        For y = 15 To 20
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 1 To 7
        boardPos(x, 20) = "barrier"
        boardPos(x, 19) = "barrier"
    Next x
    For x = 28 To 30
        For y = 6 To 9
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For x = 7 To 11
        For y = 21 To 25
            boardPos(x, y) = "barrier"
        Next y
    Next x
    boardPos(12, 4) = "barrier"
    boardPos(11, 4) = "barrier"
    'poke mart
    For x = 25 To 28
        For y = 10 To 13
            boardPos(x, y) = "barrier"
        Next y
    Next x
    'define cliffs
    For x = 2 To 6
        boardPos(x, 25) = "cliff"
    Next x
    For x = 12 To 13
        boardPos(x, 25) = "cliff"
    Next x
    For x = 15 To 30
        boardPos(x, 25) = "cliff"
    Next x
    'define talkable people
    'house1 Person
    boardPos(11, 20) = "talkable"
    'house2 Person
    boardPos(19, 13) = "talkable"
    'old man
    boardPos(13, 2) = "talkable"
    'pikachu
    boardPos(13, 3) = "talkable"
    'gardener
    boardPos(25, 22) = "talkable"
    'gym dude
    boardPos(25, 6) = "talkable"
    'gym sign
    boardPos(24, 6) = "talkable"
    End Sub

Public Sub Draw_ViridianCity()
    If canWalk = True Then
        BitBlt frmViridianCity.hDC, 0, 0, picBack.Width, picBack.Height, picBack.hDC, 0, 0, vbSrcCopy 'bitblt background onto form
        BitBlt frmViridianCity.hDC, ashX, ashY, picMask.Width, picMask.Height, picMask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
        BitBlt frmViridianCity.hDC, ashX, ashY, picSprite.Width, picSprite.Height, picSprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
    End If
End Sub

Private Sub Timer1_Timer()
    If loadMap = False Then
        Draw_ViridianCity
        loadMap = True
    End If
End Sub


