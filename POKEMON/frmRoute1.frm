VERSION 5.00
Begin VB.Form frmRoute1 
   Caption         =   "Form2"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8430
      Left            =   0
      Picture         =   "frmRoute1.frx":0000
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8370
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   360
         Left            =   3720
         Picture         =   "frmRoute1.frx":E2B42
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
         Left            =   4020
         Picture         =   "frmRoute1.frx":E2EAA
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   1
         Top             =   7755
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
Attribute VB_Name = "frmRoute1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ashX, ashY As Integer
Dim boardPos(1 To 31, -1 To 32) As String
Dim canWalk, loadMap As Boolean
Const SPEED = 17
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Randomize
    Select Case KeyCode
        Case vbKeyLeft
        picSprite.Picture = LoadPicture("oakLeft.gif")
        picMask.Picture = LoadPicture("oakLeftMask.gif")
        Draw_Route1
            If Not boardPos((ashX + 6 - SPEED) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX + 6 - SPEED) / 17, (ashY + 10) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX - SPEED 'moves guy left
                        Draw_Route1 'call the Draw_Route1 function
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
                        Call form1.GetLVL
                        frmBattle.passThis = form1.ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.arrayNum = 1
                        Call form1.GetHP
                        frmBattle.pHP = form1.ovwHP
                        frmBattle.lblpHP = form1.ovwHP
                        frmRoute1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                Else:
                    ashX = ashX - SPEED 'moves guy left
                    Draw_Route1 'call the Draw_Route1 function
                End If
            End If
        Case vbKeyRight
        picSprite.Picture = LoadPicture("oakRight.gif")
        picMask.Picture = LoadPicture("oakRightMask.gif")
        Draw_Route1
            If Not boardPos((ashX + 6 + SPEED) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX + 6 + SPEED) / 17, (ashY + 10) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX + SPEED 'moves guy right
                        Draw_Route1 'call the Draw_Route1 function
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
                        Call form1.GetLVL
                        frmBattle.passThis = form1.ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.arrayNum = 1
                        Call form1.GetHP
                        frmBattle.pHP = form1.ovwHP
                        frmBattle.lblpHP = form1.ovwHP
                        frmRoute1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                    Else:
                        ashX = ashX + SPEED 'moves guy right
                        Draw_Route1 'call the Draw_Route1 function
                End If
            End If
        Case vbKeyUp
        picSprite.Picture = LoadPicture("oakBack.gif")
        picMask.Picture = LoadPicture("oakBackMask.gif")
        Draw_Route1
            If Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "barrier" And Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY - SPEED 'moves guy up
                        Draw_Route1 'call the Draw_Route1 function
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
                        Call form1.GetLVL
                        frmBattle.passThis = form1.ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.arrayNum = 1
                        Call form1.GetHP
                        frmBattle.pHP = form1.ovwHP
                        frmBattle.lblpHP = form1.ovwHP
                        frmRoute1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                ElseIf boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "upperBound" Then
                    frmRoute1.Hide
                    frmViridianCity.Show
                    frmTxtBox.List1.Clear
                    frmTxtBox.List1.AddItem "Viridian City"
                    frmViridianCity.ashX = ashX
                    frmViridianCity.ashY = 517
                    form1.MAP = "VIRIDIANCITY"
                Else:
                    ashY = ashY - SPEED 'moves guy up
                    Draw_Route1 'call the Draw_Route1 function
                End If
            End If
        Case vbKeyDown
        picSprite.Picture = LoadPicture("oakFront.gif")
        picMask.Picture = LoadPicture("oakFrontMask.gif")
        Draw_Route1
            If Not boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "barrier" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "lowerBound" Then
                    form1.MAP = "PALLET"
                    frmTxtBox.List1.Clear
                    frmTxtBox.List1.AddItem "Pallet Town"
                    form1.Show
                    frmRoute1.Hide
                    form1.Draw_map
                    form1.ashX = ashX + 17
                    frmRoute1.ashY = 0
                ElseIf boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "cliff" Then
                    ashY = ashY + SPEED + SPEED 'moves guy down 2
                    Draw_Route1 'call the Draw_Route1 function
                Else:
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY + SPEED 'moves guy down
                        Draw_Route1 'call the Draw_Route1 function
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
                        Call form1.GetLVL
                        frmBattle.passThis = form1.ovwLVL
                        Call frmBattle.passpLVL
                        Call frmBattle.getPPokemon
                        Call frmBattle.getStats
                        frmBattle.arrayNum = 1
                        Call form1.GetHP
                        frmBattle.pHP = form1.ovwHP
                        frmBattle.lblpHP = form1.ovwHP
                        frmRoute1.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                    Else:
                        ashY = ashY + SPEED 'moves guy up
                        Draw_Route1 'call the Draw_Route1 function
                End If
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    canWalk = True
    loadMap = False
    frmRoute1.Left = 0
    frmRoute1.Top = 0
    'initialize boundary tiles
    For x = 1 To 31
        boardPos(x, 0) = "barrier"
    Next x
    For x = 15 To 17
        boardPos(x, 0) = "upperBound"
    Next x
    For x = 3 To 25
        boardPos(x, 31) = "lowerBound"
    Next x
    'initialize collisions
    For x = 1 To 13
        boardPos(x, 30) = "barrier"
    Next x
    For x = 17 To 31
        boardPos(x, 30) = "barrier"
    Next x
    For y = 1 To 30
        boardPos(6, y) = "barrier"
        boardPos(28, y) = "barrier"
    Next y
    For x = 11 To 27
        boardPos(x, 25) = "cliff"
    Next x
    For x = 7 To 14
        boardPos(x, 20) = "barrier"
        boardPos(x, 19) = "barrier"
    Next x
    For x = 22 To 27
        boardPos(x, 20) = "cliff"
    Next x
    boardPos(7, 25) = "cliff"
    boardPos(7, 13) = "cliff"
    For x = 9 To 11
        boardPos(x, 13) = "cliff"
    Next x
    For x = 14 To 27
        boardPos(x, 13) = "cliff"
    Next x
    boardPos(7, 7) = "barrier"
    boardPos(7, 8) = "barrier"
    For x = 8 To 15
        boardPos(x, 8) = "cliff"
    Next x
    For x = 16 To 19
        boardPos(x, 8) = "barrier"
        boardPos(x, 7) = "barrier"
    Next x
    For x = 7 To 11
        boardPos(x, 5) = "cliff"
        boardPos(x, 2) = "cliff"
    Next x
    For x = 12 To 13
        For y = 2 To 5
            boardPos(x, y) = "barrier"
        Next y
    Next x
    
    'initialize grass tiles
    For x = 14 To 27
        For y = 3 To 4
            boardPos(x, y) = "grass"
        Next y
    Next x
    For x = 20 To 27
        For y = 6 To 10
            boardPos(x, y) = "grass"
        Next y
    Next x
    For x = 15 To 21
        For y = 18 To 22
            boardPos(x, y) = "grass"
        Next y
    Next x
    For x = 10 To 12
        For y = 26 To 30
            boardPos(x, y) = "grass"
        Next y
    Next x
    End Sub

Public Sub Draw_Route1()
    If canWalk = True Then
        BitBlt frmRoute1.hDC, 0, 0, picBack.Width, picBack.Height, picBack.hDC, 0, 0, vbSrcCopy 'bitblt background onto form
        BitBlt frmRoute1.hDC, ashX, ashY, picMask.Width, picMask.Height, picMask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
        BitBlt frmRoute1.hDC, ashX, ashY, picSprite.Width, picSprite.Height, picSprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
    End If
End Sub

Private Sub Timer1_Timer()
    If loadMap = False Then
        Draw_Route1
        loadMap = True
    End If
    frmRoute1.Left = 0
    frmRoute1.Top = 0
End Sub

