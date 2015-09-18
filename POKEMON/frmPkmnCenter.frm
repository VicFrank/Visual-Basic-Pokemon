VERSION 5.00
Begin VB.Form frmPkmnCenter 
   Caption         =   "Form2"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   2865
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2880
      Left            =   0
      Picture         =   "frmPkmnCenter.frx":0000
      ScaleHeight     =   188
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   204
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3120
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   2280
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2040
         Picture         =   "frmPkmnCenter.frx":1C1B2
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   2
         Top             =   2400
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   360
         Left            =   1200
         Picture         =   "frmPkmnCenter.frx":1C583
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   1
         Top             =   2400
         Visible         =   0   'False
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmPkmnCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ashX, ashY As Integer
Dim boardPos(-1 To 12, 3 To 11) As String
Dim canWalk, loadMap, canTalk As Boolean
Const SPEED = 17
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Randomize
    'are you standing next to Nurse Joy?
    If boardPos((ashX) / 17, (ashY + 5) / 17) = "heal" Then
        canTalk = True
    End If
    
    Select Case KeyCode
        Case vbKeyLeft
        picSprite.Picture = LoadPicture("oakLeft.gif")
        picMask.Picture = LoadPicture("oakLeftMask.gif")
        Draw_pkmnCenter
            If Not boardPos((ashX - SPEED) / 17, (ashY + 5) / 17) = "barrier" And Not boardPos((ashX - SPEED) / 17, (ashY + 5) / 17) = "cliff" And canWalk = True Then
                canTalk = False
                If boardPos((ashX) / 17, (ashY + 5) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX - SPEED 'moves guy left
                        Draw_pkmnCenter 'call the Draw_pkmnCenter function
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
                    End If
                Else:
                    ashX = ashX - SPEED 'moves guy left
                    Draw_pkmnCenter 'call the Draw_pkmnCenter function
                End If
            End If
        Case vbKeyRight
        picSprite.Picture = LoadPicture("oakRight.gif")
        picMask.Picture = LoadPicture("oakRightMask.gif")
        Draw_pkmnCenter
            If Not boardPos((ashX + SPEED) / 17, (ashY + 5) / 17) = "barrier" And Not boardPos((ashX + SPEED) / 17, (ashY + 5) / 17) = "cliff" And canWalk = True Then
                canTalk = False
                If boardPos((ashX) / 17, (ashY + 5) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX + SPEED 'moves guy right
                        Draw_pkmnCenter 'call the Draw_pkmnCenter function
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
                    End If
                    Else:
                        ashX = ashX + SPEED 'moves guy right
                        Draw_pkmnCenter 'call the Draw_pkmnCenter function
                End If
            End If
        Case vbKeyUp
        picSprite.Picture = LoadPicture("oakBack.gif")
        picMask.Picture = LoadPicture("oakBackMask.gif")
        Draw_pkmnCenter
            If Not boardPos((ashX) / 17, (ashY + 5 - SPEED) / 17) = "barrier" And Not boardPos((ashX) / 17, (ashY + 5 - SPEED) / 17) = "cliff" And canWalk = True Then
                ashY = ashY - SPEED 'moves guy up
                Draw_pkmnCenter 'call the Draw_pkmnCenter function
            End If
        Case vbKeyDown
        picSprite.Picture = LoadPicture("oakFront.gif")
        picMask.Picture = LoadPicture("oakFrontMask.gif")
        Draw_pkmnCenter
            If Not boardPos((ashX) / 17, (ashY + 5 + SPEED) / 17) = "barrier" And canWalk = True Then
                canTalk = False
                If boardPos((ashX) / 17, (ashY + 5 + SPEED) / 17) = "lowerBound" Then
                    Select Case form1.MAP
                    Case "VIRIDIANCITY"
                        frmViridianCity.Show
                        frmPkmnCenter.Hide
                        frmViridianCity.Draw_ViridianCity
                        frmViridianCity.ashX = 319
                        frmViridianCity.ashY = 383
                        frmPkmnCenter.ashY = 0
                    End Select
                Else:
                If boardPos((ashX) / 17, (ashY + 5) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY + SPEED 'moves guy down
                        Draw_pkmnCenter 'call the Draw_pkmnCenter function
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
                    End If
                    Else:
                        ashY = ashY + SPEED 'moves guy up
                        Draw_pkmnCenter 'call the Draw_pkmnCenter function
                End If
                End If
            End If
        Case vbKeyReturn
            If canTalk = True Then
                frmTxtBox.List1.AddItem "Welcome to the Pkmn Center"
                frmTxtBox.List1.AddItem "We will now heal your injured Pokemon"
                For x = 1 To frmBattle.partyNum
                    frmBattle.pokeNum = x
                    frmBattle.pokeOwner = "p"
                    Call frmBattle.getPPokemon
                    Call frmBattle.getStats
                    frmBattle.arrayNum = x
                    frmBattle.passThis = frmBattle.lblMaxHP
                    Call form1.ArrayFixer
                Next x
            End If
    End Select
End Sub

Private Sub Form_Load()
    frmPkmnCenter.Left = 0
    frmPkmnCenter.Top = 0
    canWalk = True
    canTalk = False
    loadMap = False
    'initialize boundary tiles
    boardPos(6, 11) = "lowerBound"
    For x = 0 To 5
        boardPos(x, 11) = "barrier"
        boardPos(x, 3) = "barrier"
    Next x
    boardPos(6, 3) = "barrier"
    For x = 7 To 11
        boardPos(x, 11) = "barrier"
        boardPos(x, 3) = "barrier"
    Next x
    For y = 3 To 11
        boardPos(-1, y) = "barrier"
        boardPos(12, y) = "barrier"
    Next y
    boardPos(6, 4) = "heal"
    End Sub

Public Sub Draw_pkmnCenter()
    If canWalk = True Then
        BitBlt frmPkmnCenter.hDC, 0, 0, picBack.Width, picBack.Height, picBack.hDC, 0, 0, vbSrcCopy 'bitblt background onto form
        BitBlt frmPkmnCenter.hDC, ashX, ashY, picMask.Width, picMask.Height, picMask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
        BitBlt frmPkmnCenter.hDC, ashX, ashY, picSprite.Width, picSprite.Height, picSprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
    End If
End Sub

Private Sub Timer1_Timer()
    If loadMap = False Then
        Draw_pkmnCenter
        loadMap = True
    End If
End Sub
