VERSION 5.00
Begin VB.Form Route28 
   Caption         =   "Form2"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form2"
   ScaleHeight     =   3780
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3855
      Left            =   0
      Picture         =   "Route28.frx":0000
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8385
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   360
         Left            =   5400
         Picture         =   "Route28.frx":670B6
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   2
         Top             =   2520
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
         Left            =   4680
         Picture         =   "Route28.frx":6741E
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1920
         Top             =   2880
      End
   End
End
Attribute VB_Name = "Route28"
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
        Draw_route28
            If Not boardPos((ashX + 6 - SPEED) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX + 6 - SPEED) / 17, (ashY + 10) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashX = ashX - SPEED 'moves guy left
                        Draw_route28 'call the Draw_route28 function
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
                        frmroute28.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                Else:
                    ashX = ashX - SPEED 'moves guy left
                    Draw_route28 'call the Draw_route28 function
                End If
            End If
        Case vbKeyRight
        picSprite.Picture = LoadPicture("oakRight.gif")
        picMask.Picture = LoadPicture("oakRightMask.gif")
        Draw_route28
            If Not boardPos((ashX + 6 + SPEED) / 17, (ashY + 10) / 17) = "barrier" And Not boardPos((ashX + 6 + SPEED) / 17, (ashY + 10) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "rightBound" Then
                        form1.MAP = "VIRIDIANCITY"
                        frmTxtBox.List1.Clear
                        frmTxtBox.List1.AddItem "Viridian City"
                        frmViridianCity.Show
                        frmroute28.Hide
                        frmViridianCity.Draw_ViridianCity
                        form1.ashX = ashX + 17
                        frmroute28.ashY = 0
                    ElseIf (Rnd * 100) + 1 > 6 Then
                        ashX = ashX + SPEED 'moves guy right
                        Draw_route28 'call the Draw_route28 function
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
                        frmroute28.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                    Else:
                        ashX = ashX + SPEED 'moves guy right
                        Draw_route28 'call the Draw_route28 function
                End If
            End If
        Case vbKeyUp
        picSprite.Picture = LoadPicture("oakBack.gif")
        picMask.Picture = LoadPicture("oakBackMask.gif")
        Draw_route28
            If Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "barrier" And Not boardPos((ashX + 6) / 17, (ashY + 10 - SPEED) / 17) = "cliff" And canWalk = True Then
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY - SPEED 'moves guy up
                        Draw_route28 'call the Draw_route28 function
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
                        frmroute28.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                ElseIf boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "upperBound" Then
                    frmroute28.Hide
                    frmViridianCity.Show
                    frmTxtBox.List1.Clear
                    frmTxtBox.List1.AddItem "Viridian City"
                    frmViridianCity.ashX = ashX
                    frmViridianCity.ashY = 517
                    form1.MAP = "VIRIDIANCITY"
                Else:
                    ashY = ashY - SPEED 'moves guy up
                    Draw_route28 'call the Draw_route28 function
                End If
            End If
        Case vbKeyDown
        picSprite.Picture = LoadPicture("oakFront.gif")
        picMask.Picture = LoadPicture("oakFrontMask.gif")
        Draw_route28
            If Not boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "barrier" And canWalk = True Then
                ElseIf boardPos((ashX + 6) / 17, (ashY + 10 + SPEED) / 17) = "cliff" Then
                    ashY = ashY + SPEED + SPEED 'moves guy down 2
                    Draw_route28 'call the Draw_route28 function
                Else:
                If boardPos((ashX + 6) / 17, (ashY + 10) / 17) = "grass" Then
                    If (Rnd * 100) + 1 > 6 Then
                        ashY = ashY + SPEED 'moves guy down
                        Draw_route28 'call the Draw_route28 function
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
                        frmroute28.Hide
                        frmTxtBox.List1.Clear
                        frmTxtBox.Hide
                        frmBattle.Show
                    End If
                    Else:
                        ashY = ashY + SPEED 'moves guy up
                        Draw_route28 'call the Draw_route28 function
                End If
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    canWalk = True
    loadMap = False
    frmroute28.Left = 0
    frmroute28.Top = 0
    'initialize boundary tiles
    End Sub

Public Sub Draw_route28()
    If canWalk = True Then
        BitBlt frmroute28.hDC, 0, 0, picBack.Width, picBack.Height, picBack.hDC, 0, 0, vbSrcCopy 'bitblt background onto form
        BitBlt frmroute28.hDC, ashX, ashY, picMask.Width, picMask.Height, picMask.hDC, 0, 0, vbSrcAnd 'bitblt mask onto form
        BitBlt frmroute28.hDC, ashX, ashY, picSprite.Width, picSprite.Height, picSprite.hDC, 0, 0, vbSrcPaint 'bitblt sprite onto form
    End If
End Sub

Private Sub Timer1_Timer()
    If loadMap = False Then
        Draw_route28
        loadMap = True
    End If
    frmroute28.Left = 0
    frmroute28.Top = 0
End Sub


