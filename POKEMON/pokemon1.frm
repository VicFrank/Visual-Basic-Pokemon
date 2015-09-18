VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8430
      Left            =   -200
      Picture         =   "pokemon1.frx":0000
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   0
      Top             =   -240
      Visible         =   0   'False
      Width           =   8370
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   600
         Top             =   7200
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   360
         Left            =   4200
         Picture         =   "pokemon1.frx":E2B42
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   2
         Top             =   3360
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   360
         Left            =   3960
         Picture         =   "pokemon1.frx":E2F13
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   1
         Top             =   3360
         Visible         =   0   'False
         Width           =   270
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ashX, ashY As Integer
Dim boardPos(1 To 31, 0 To 31), person As String
Public MAP As String
Dim canWalk, canTalk As Boolean
Const SPEED = 17
'battle constant storage
Private overworldpHP(1 To 3), overworldLVL(1 To 3) As Integer
Public ovwHP, ovwLVL, passThis, arrayNum As Integer
'item storage
Public potionNum, pokeBallNum As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Randomize
    Select Case KeyCode
        Case vbKeyLeft
        picSprite.Picture = LoadPicture("oakLeft.gif")
        picMask.Picture = LoadPicture("oakLeftMask.gif")
        Draw_map
            If Not boardPos((ashX - SPEED + 6) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX - SPEED + 6) / 17, (ashY + 10) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX - SPEED 'moves guy left
                        Draw_map 'call the Draw_map function
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
                        frmBattle.arrayNum = 1
                        frmBattle.passThis = overworldLVL(1)
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.pHP = overworldpHP(1)
                        frmBattle.lblpHP = overworldpHP(1)
                        form1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                Else:
                    ashX = ashX - SPEED 'moves guy left
                    Draw_map 'call the Draw_map function
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX - SPEED + 6) / 17, (ashY + 9) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4 - SPEED) / 17) = 10 And ((ashY + 8) / 17) = 15 Then
                    person = "agatha"
                ElseIf ((ashX + 4 - SPEED) / 17) = 21 And ((ashY + 8) / 17) = 15 Then
                    person = "gary"
                ElseIf ((ashX + 4 - SPEED) / 17) = 11 And ((ashY + 8) / 17) = 22 Then
                    person = "fatman"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyRight
        picSprite.Picture = LoadPicture("oakRight.gif")
        picMask.Picture = LoadPicture("oakRightMask.gif")
        Draw_map
            If Not boardPos((ashX + SPEED + 6) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX + SPEED + 6) / 17, (ashY + 10) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX + SPEED 'moves guy right
                        Draw_map 'call the Draw_map function
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
                        frmBattle.arrayNum = 1
                        frmBattle.passThis = overworldLVL(1)
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.pHP = overworldpHP(1)
                        frmBattle.lblpHP = overworldpHP(1)
                        form1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                    Else:
                        ashX = ashX + SPEED 'moves guy right
                        Draw_map 'call the Draw_map function
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX + SPEED + 6) / 17, (ashY + 8) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4 + SPEED) / 17) = 10 And ((ashY + 8) / 17) = 15 Then
                    person = "agatha"
                ElseIf ((ashX + 4 + SPEED) / 17) = 21 And ((ashY + 8) / 17) = 15 Then
                    person = "gary"
                ElseIf ((ashX + 4 + SPEED) / 17) = 11 And ((ashY + 8) / 17) = 22 Then
                    person = "fatman"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyUp
        picSprite.Picture = LoadPicture("oakBack.gif")
        picMask.Picture = LoadPicture("oakBackMask.gif")
        Draw_map
            If Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "barrier" And Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY - SPEED 'moves guy up
                        Draw_map 'call the Draw_map function
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
                        frmBattle.arrayNum = 1
                        frmBattle.passThis = overworldLVL(1)
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.pHP = overworldpHP(1)
                        frmBattle.lblpHP = overworldpHP(1)
                        form1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                ElseIf boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "upperBound" Then
                    MAP = "ROUTE1"
                    frmTxtBox.List1.Clear
                    frmTxtBox.List1.AddItem "Route 1"
                    frmRoute1.Show
                    frmRoute1.ashX = ashX - 17
                    frmRoute1.ashY = 500
                    Call frmRoute1.Draw_Route1
                    form1.Hide
                Else:
                    ashY = ashY - SPEED 'moves guy up
                    Draw_map 'call the Draw_map function
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX + 6) / 17, (ashY + 8 - SPEED) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4) / 17) = 10 And ((ashY + 8 - SPEED) / 17) = 15 Then
                    person = "agatha"
                ElseIf ((ashX + 4) / 17) = 21 And ((ashY + 8 - SPEED) / 17) = 15 Then
                    person = "gary"
                ElseIf ((ashX + 4) / 17) = 11 And ((ashY + 8 - SPEED) / 17) = 22 Then
                    person = "fatman"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyDown
        picSprite.Picture = LoadPicture("oakFront.gif")
        picMask.Picture = LoadPicture("oakFrontMask.gif")
        Draw_map
            If Not boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "barrier" And Not boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "talkable" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY + SPEED 'moves guy down
                        Draw_map 'call the Draw_map function
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
                        frmBattle.arrayNum = 1
                        frmBattle.passThis = overworldLVL(1)
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.pHP = overworldpHP(1)
                        frmBattle.lblpHP = overworldpHP(1)
                        form1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                     Else:
                        ashY = ashY + SPEED 'moves guy up
                        Draw_map 'call the Draw_map function
                End If
            End If
            'are you facing the direction of someone you can talk to?
            If boardPos((ashX + 6) / 17, (ashY + 8 + SPEED) / 17) = "talkable" Then
                canTalk = True
                If ((ashX + 4) / 17) = 10 And ((ashY + 8 + SPEED) / 17) = 15 Then
                    person = "agatha"
                ElseIf ((ashX + 4) / 17) = 21 And ((ashY + 8 + SPEED) / 17) = 15 Then
                    person = "gary"
                ElseIf ((ashX + 4) / 17) = 11 And ((ashY + 8 + SPEED) / 17) = 22 Then
                    person = "fatman"
                End If
            Else:
                canTalk = False
            End If
        Case vbKeyReturn
            If canTalk = True Then
                frmTxtBox.List1.Clear
                Select Case person
                    Case "agatha"
                        frmTxtBox.List1.AddItem "AGATHA: OAK you old duff!"
                        frmTxtBox.List1.AddItem "I've moved into this house!"
                    Case "gary"
                        frmTxtBox.List1.AddItem "BLUE: Hey gramps! After I lost to RED, I moved back"
                        frmTxtBox.List1.AddItem "to Pallet Town. I'm going to get even stronger!"
                    Case "fatman"
                        frmTxtBox.List1.AddItem "I've heard that this game is riddled with bugs!"
                        frmTxtBox.List1.AddItem "TECHNOLOGY IS AMAZING!"
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
    frmTxtBox.Show
    frmTxtBox.List1.AddItem "Pallet Town"
    frmTxtBox.List1.AddItem "Welcome to VBPokemon Alpha!"
    frmTxtBox.List1.AddItem "Currently, you can only have up to 3 Pokemon"
    form1.Left = 0
    form1.Top = 0
    MAP = "PALLET"
    'get your poke's stats
    overworldLVL(1) = 5
    overworldpHP(1) = 18
    'init position
    ashX = 268 '+ 6
    ashY = 230 '+ 10
    canWalk = True
    Draw_map
    'initialize boundaries
    For y = 6 To 25
        boardPos(6, y) = "barrier"
        boardPos(27, y) = "barrier"
    Next y
    For x = 6 To 14
        boardPos(x, 5) = "barrier"
        boardPos(x, 25) = "barrier"
    Next x
    For x = 15 To 17
        boardPos(x, 1) = "upperBound"
        boardPos(x, 25) = "barrier"
    Next x
    For x = 18 To 27
        boardPos(x, 5) = "barrier"
        boardPos(x, 25) = "barrier"
    Next x
    For y = 1 To 5
        boardPos(14, y) = "barrier"
        boardPos(18, y) = "barrier"
    Next y
    'house1
    For x = 9 To 12
         For y = 11 To 14
            boardPos(x, y) = "barrier"
        Next y
    Next x
    'house2
    For x = 20 To 23
        For y = 11 To 14
            boardPos(x, y) = "barrier"
        Next y
    Next x
    'laboratory
    For x = 16 To 20
        For y = 18 To 22
            boardPos(x, y) = "barrier"
        Next y
    Next x
    For y = 20 To 22
        boardPos(15, y) = "barrier"
        boardPos(21, y) = "barrier"
        boardPos(22, y) = "barrier"
        boardPos(23, y) = "barrier"
    Next y
    'initialize grass tiles
    For x = 15 To 17
        For y = 2 To 8
            boardPos(x, y) = "grass"
        Next y
    Next x
    'Gary Oak
    boardPos(21, 15) = "talkable"
    'Agatha
    boardPos(10, 15) = "talkable"
    'Fat man
    boardPos(11, 22) = "talkable"
    'hide unnessecary forms
    frmBattle.Hide
    End Sub

Public Sub Draw_map()
    If canWalk = True Then
        BitBlt form1.hDC, 0, 0, picBack.Width, picBack.Height, picBack.hDC, 0, 0, vbSrcCopy 'bitblt background onto form
        BitBlt form1.hDC, ashX, ashY, picMask.Width, picMask.Height, picMask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
        BitBlt form1.hDC, ashX, ashY, picSprite.Width, picSprite.Height, picSprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
    End If
End Sub

Public Sub ArrayFixer() 'max2
    'for passing around arrays between forms
    overworldpHP(frmBattle.arrayNum) = frmBattle.passThis
End Sub

Public Sub GetHP()
    'for getting your poke's HP
    ovwHP = overworldpHP(frmBattle.arrayNum)
End Sub

Public Sub PassLVL()
    overworldLVL(frmBattle.arrayNum) = frmBattle.passThis
End Sub

Public Sub GetLVL()
    ovwLVL = overworldLVL(frmBattle.arrayNum)
End Sub

Private Sub Timer1_Timer()
    If loadMap = False Then
        Call Draw_map
        loadMap = True
    End If
    form1.Left = 0
    form1.Top = 0
End Sub
