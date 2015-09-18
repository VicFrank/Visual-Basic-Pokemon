VERSION 5.00
Begin VB.Form frmBattle 
   Caption         =   "Form2"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form2"
   ScaleHeight     =   4125
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.CommandButton cmdAtk4 
      Caption         =   "Run"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdAtk3 
      Caption         =   "Pokemon"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdAtk2 
      Caption         =   "Items"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdAtk1 
      Caption         =   "Attack"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.PictureBox picBattle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   0
      Picture         =   "frmBattle.frx":0000
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin VB.PictureBox trainerBattle 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   2235
         Left            =   0
         Picture         =   "frmBattle.frx":1B586
         ScaleHeight     =   2175
         ScaleWidth      =   3855
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   0
         Picture         =   "frmBattle.frx":36B0C
         ScaleHeight     =   1260
         ScaleWidth      =   1845
         TabIndex        =   13
         Top             =   930
         Width           =   1845
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   1950
         Picture         =   "frmBattle.frx":3E55E
         ScaleHeight     =   1575
         ScaleWidth      =   1905
         TabIndex        =   10
         Top             =   0
         Width           =   1905
      End
      Begin VB.Label lblMaxHP 
         BackColor       =   &H00C0FFFF&
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   3480
         TabIndex        =   12
         Top             =   1935
         Width           =   255
      End
      Begin VB.Label lblpHP 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   3225
         TabIndex        =   11
         Top             =   1935
         Width           =   135
      End
      Begin VB.Label lblYourPoke 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2475
         TabIndex        =   9
         Top             =   1635
         Width           =   975
      End
      Begin VB.Label lblEPoke 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   3
         Top             =   380
         Width           =   855
      End
      Begin VB.Label lblELvl 
         BackColor       =   &H00C0FFFF&
         Caption         =   "15"
         Height          =   200
         Left            =   1170
         TabIndex        =   2
         Top             =   380
         Width           =   210
      End
      Begin VB.Label lblYourLevel 
         BackColor       =   &H00C0FFFF&
         Height          =   200
         Left            =   3600
         TabIndex        =   1
         Top             =   1635
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'enemy pokemon constants
Public ePokeNum, ePoke, eLVL, eBaseHP, eBaseATK, eBaseDef, eBaseSpAtk, eBaseSpDef, eBaseSpd, eHP, maxeHP, eAtk, eDef, eSpAtk, eSpDef, eSpd As Integer
'player pokemon constants
Public pPoke, pBaseHP, pBaseATK, pBaseDef, pBaseSpAtk, pBaseSpDef, pBaseSpd, pHP, pAtk, pDef, pSpAtk, pSpDef, pSpd As Integer
'battle constants
Public power, CH, BRN, STAB, damage, accuracy, pAccuracy, eAccuracy, atkNum, baseEXP, catchRate As Integer
Public atkType, debuff, buff, pokeOwner, atkName As String
'player/enemy attacks and types
Public eAtkName1, eAtkName2, eAtkName3, eAtkName4, eType1, eType2 As String
Private pAtkName(1 To 4, 1 To 3), pType1, pType2, compare As String 'max3
Public inBattle, fullParty As Boolean
'private arrays, because apparently YOU CAN'T PASS THEM BETWEEN FORMS tt
Private exp(1 To 3), pPokeNum(1 To 3), pLVL(1 To 3), typeBonus1, typeBonus2 As Integer 'max3
Public pokeNum, partyNum, index, passThis, arrayNum, tEXP As Integer
'variables to make this a trainer battle
Public canCatch As Boolean

Private Sub Form_Load()
    inBattle = True
    canCatch = True
    BRN = 1
    CH = 1
    tEXP = 1
    pAccuracy = 1
    'load starter pokemon variables
    pokeNum = 1
    partyNum = 1
    pPokeNum(1) = 4
    exp(1) = 64
    arrayNum = 1
    Call form1.GetLVL
    pLVL(1) = form1.ovwLVL
    fullParty = False
    frmBattle.Left = 0
    frmBattle.Top = 0
End Sub

Public Sub getEPokemon()
    Select Case ePokeNum
        Case 10
            'Caterpie
            index = 10
            ePoke = "CATERPIE"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 40
            eBaseATK = 30
            eBaseDef = 35
            eBaseSpAtk = 20
            eBaseSpDef = 20
            eBaseSpd = 45
            eAtkName1 = "TACKLE"
            eAtkName2 = "STRINGSHOT"
            eType1 = "BUG"
            eType2 = ""
            baseEXP = 39
            catchRate = 255
            Picture1.Picture = LoadPicture("caterpieFront.bmp") 'TODO
        Case 11
            'Metapod
            index = 11
            ePoke = "METAPOD"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 50
            eBaseATK = 20
            eBaseDef = 55
            eBaseSpAtk = 25
            eBaseSpDef = 25
            eBaseSpd = 30
            eAtkName1 = "HARDEN"
            eAtkName2 = ""
            eType1 = "BUG"
            eType2 = ""
            baseEXP = 72
            catchRate = 120
            Picture1.Picture = LoadPicture("metapodFront.bmp") 'TODO
        Case 13
            'Weedle
            index = 13
            ePoke = "WEEDLE"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 40
            eBaseATK = 35
            eBaseDef = 30
            eBaseSpAtk = 20
            eBaseSpDef = 20
            eBaseSpd = 50
            eAtkName1 = "POISONSTING"
            eAtkName2 = "STRINGSHOT"
            eType1 = "BUG"
            eType2 = "POISON"
            baseEXP = 39
            catchRate = 255
            Picture1.Picture = LoadPicture("weedleFront.bmp")
        Case 14
            'KAKUNA
            index = 14
            ePoke = "WEEDLE"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 45
            eBaseATK = 25
            eBaseDef = 50
            eBaseSpAtk = 25
            eBaseSpDef = 25
            eBaseSpd = 35
            eAtkName1 = "HARDEN"
            eAtkName2 = ""
            eType1 = "BUG"
            eType2 = "POISON"
            baseEXP = 72
            catchRate = 120
            Picture1.Picture = LoadPicture("kakunaFront.bmp")
        Case 16
            'Pidgey
            index = 16
            ePoke = "PIDGEY"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 40
            eBaseATK = 45
            eBaseDef = 40
            eBaseSpAtk = 35
            eBaseSpDef = 35
            eBaseSpd = 56
            eAtkName1 = "GUST"
            eType1 = "FLYING"
            eType2 = "NORMAL"
            baseEXP = 50
            catchRate = 255
            Select Case eLVL
                Case eLVL > 4
                    eAtkName2 = "SANDATTACK"
                Case eLVL > 11
                    eAtkName3 = "QUICKATTACK"
            End Select
            Picture1.Picture = LoadPicture("pidgeyFront.bmp")
        Case 18
            'Rattata
            index = 18
            ePoke = "RATTATA"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 30
            eBaseATK = 56
            eBaseDef = 35
            eBaseSpAtk = 25
            eBaseSpDef = 35
            eBaseSpd = 72
            eAtkName1 = "TACKLE"
            eAtkName2 = "TAILWHIP"
            eType1 = "NORMAL"
            eType2 = ""
            baseEXP = 57
            catchRate = 255
            Select Case eLVL
                Case eLVL > 3
                    eAtkName3 = "QUICKATTACK"
                Case eLVL > 9
                    eAtkName4 = "BITE"
            End Select
            Picture1.Picture = LoadPicture("rattataFront.bmp")
        Case 25
            'Pikachu
            index = 25
            ePoke = "PIKACHU"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 35
            eBaseATK = 55
            eBaseDef = 30
            eBaseSpAtk = 50
            eBaseSpDef = 40
            eBaseSpd = 90
            eAtkName1 = "GROWL"
            eAtkName2 = "THUNDERSHOCK"
            eType1 = "ELECTRIC"
            eType2 = ""
            baseEXP = 82
            catchRate = 190
            Select Case eLVL
                Case eLVL > 5
                    eAtkName3 = "TAILWHIP"
                Case eLVL > 10
                    eAtkName4 = "THUNDERWAVE"
            End Select
            Picture1.Picture = LoadPicture("pikachuFront.bmp")
        Case 143
            'Snorlax
            index = 143
            ePoke = "SNORLAX"
            lblEPoke.Caption = ePoke
            lblELvl.Caption = eLVL
            eBaseHP = 160
            eBaseATK = 110
            eBaseDef = 65
            eBaseSpAtk = 65
            eBaseSpDef = 110
            eBaseSpd = 30
            eAtkName1 = "TACKLE"
            eAtkName2 = "DEFENSECURL"
            eType1 = "NORMAL"
            eType2 = ""
            baseEXP = 154
            catchRate = 25
            Picture1.Picture = LoadPicture("snorlaxFront.bmp")
    End Select
End Sub

Public Sub getPPokemon()
    'all pokemon use medium fast exp formula
    Select Case pPokeNum(pokeNum)
    Case 4
        'Charmander
        pPoke = "CHARMANDER"
        lblYourLevel.Caption = pLVL(1)
        lblYourPoke.Caption = pPoke
        pBaseHP = 39
        pBaseATK = 52
        pBaseDef = 43
        pBaseSpAtk = 60
        pBaseSpDef = 50
        pBaseSpd = 65
        pType1 = "FIRE"
        pType2 = ""
        pAtkName(1, pokeNum) = "SCRATCH"
        pAtkName(2, pokeNum) = "GROWL"
        pAtkName(3, pokeNum) = ""
        pAtkName(4, pokeNum) = ""
        If pLVL(partyNum) > 6 Then
            pAtkName(3, pokeNum) = "EMBER"
        End If
        If pLVL(partyNum) > 9 Then
            pAtkName(4, pokeNum) = "SMOKESCREEN"
        End If
        Picture2.Picture = LoadPicture("charBack.bmp")
    Case 10
        'Caterpie
        pPoke = "CATERPIE"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 40
        pBaseATK = 35
        pBaseDef = 30
        pBaseSpAtk = 20
        pBaseSpDef = 20
        pBaseSpd = 50
        pType1 = "BUG"
        pType2 = ""
        pAtkName(1, partyNum) = "TACKLE"
        pAtkName(2, partyNum) = "STRINGSHOT"
        pAtkName(3, partyNum) = ""
        pAtkName(4, partyNum) = ""
        Picture2.Picture = LoadPicture("caterpieBack.bmp") 'TODO
    Case 11
        'Metapod
        pPoke = "METAPOD"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 50
        pBaseATK = 20
        pBaseDef = 55
        pBaseSpAtk = 25
        pBaseSpDef = 25
        pBaseSpd = 30
        pType1 = "BUG"
        pType2 = ""
        pAtkName(1, partyNum) = "HARDEN"
        pAtkName(2, partyNum) = ""
        pAtkName(3, partyNum) = ""
        pAtkName(4, partyNum) = ""
        Picture2.Picture = LoadPicture("metapodBack.bmp") 'TODO
    Case 13
        'Weedle
        pPoke = "WEEDLE"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 40
        pBaseATK = 35
        pBaseDef = 30
        pBaseSpAtk = 20
        pBaseSpDef = 20
        pBaseSpd = 50
        pType1 = "POISON"
        pType2 = "BUG"
        pAtkName(1, partyNum) = "POISONSTING"
        pAtkName(2, partyNum) = "STRINGSHOT"
        pAtkName(3, partyNum) = ""
        pAtkName(4, partyNum) = ""
        Picture2.Picture = LoadPicture("weedleBack.bmp")
    Case 14
        'Kakuna
        pPoke = "Kakuna"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 40
        pBaseATK = 35
        pBaseDef = 30
        pBaseSpAtk = 20
        pBaseSpDef = 20
        pBaseSpd = 50
        pType1 = "POISON"
        pType2 = "BUG"
        pAtkName(1, partyNum) = "HARDEN"
        pAtkName(2, partyNum) = ""
        pAtkName(3, partyNum) = ""
        pAtkName(4, partyNum) = ""
        Picture2.Picture = LoadPicture("kakunaBack.bmp")
    Case 16
        'Pidgey
        pPoke = "PIDGEY"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 40
        pBaseATK = 45
        pBaseDef = 40
        pBaseSpAtk = 35
        pBaseSpDef = 35
        pBaseSpd = 56
        pAtkName(1, partyNum) = "GUST"
        pType1 = "FLYING"
        pType2 = "NORMAL"
        Select Case pLVL(partyNum)
            Case pLVL(partyNum) > 4
                pAtkName(2, partyNum) = "SANDATTACK"
            Case pLVL(partyNum) > 11
                pAtkName(3, partyNum) = "QUICKATTACK"
        End Select
        Picture2.Picture = LoadPicture("pidgeyBack.bmp")
    Case 18
        'Rattata
        pPoke = "RATTATA"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 30
        pBaseATK = 56
        pBaseDef = 35
        pBaseSpAtk = 25
        pBaseSpDef = 35
        pBaseSpd = 72
        pAtkName(1, partyNum) = "TACKLE"
        pAtkName(2, partyNum) = "TAILWHIP"
        pType1 = "NORMAL"
        pType2 = ""
        Select Case pLVL(partyNum)
            Case pLVL(partyNum) > 3
                pAtkName(3, partyNum) = "QUICKATTACK"
            Case pLVL(partyNum) > 9
                pAtkName(4, partyNum) = "BITE"
        End Select
        Picture2.Picture = LoadPicture("rattataBack.bmp")
    Case 25
        'Pikachu
        pPoke = "PIKACHU"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 35
        pBaseATK = 55
        pBaseDef = 30
        pBaseSpAtk = 50
        pBaseSpDef = 40
        pBaseSpd = 90
        pAtkName(1, partyNum) = "GROWL"
        pAtkName(2, partyNum) = "THUNDERSHOCK"
        pType1 = "ELECTRIC"
        pType2 = ""
        Select Case pLVL(partyNum)
            Case pLVL(partyNum) > 3
                pAtkName(3, partyNum) = "TAILWHIP"
            Case pLVL(partyNum) > 9
                pAtkName(4, partyNum) = "THUNDERWAVE"
        End Select
        Picture2.Picture = LoadPicture("pikachuBack.bmp")
    Case 143
        'Snorlax
        pPoke = "SNORLAX"
        lblYourPoke.Caption = pPoke
        arrayNum = pokeNum
        Call form1.GetLVL
        lblYourLevel.Caption = form1.ovwLVL
        pLVL(pokeNum) = form1.ovwLVL
        pBaseHP = 160
        pBaseATK = 110
        pBaseDef = 65
        pBaseSpAtk = 65
        pBaseSpDef = 110
        pBaseSpd = 30
        pAtkName(1, partyNum) = "TACKLE"
        pAtkName(2, partyNum) = "DEFENSECURL"
        pType1 = "NORMAL"
        pType2 = ""
        Picture2.Picture = LoadPicture("snorlaxBack.bmp")
    End Select
End Sub
    
Public Sub getStats()
    'calculate the stats for a poke based on it's level and base stats
    If pokeOwner = "e" Then
            eHP = Int((((eBaseHP + 50) * eLVL) / 50) + 10)
            maxeHP = eHP
            eAtk = Int(((eBaseATK * eLVL) / 50) + 5)
            eDef = Int(((eBaseDef * eLVL) / 50) + 5)
            eSpAtk = Int(((eBaseSpAtk * eLVL) / 50) + 5)
            eSpDef = Int(((eBaseSpDef * eLVL) / 50) + 5)
            eSpd = Int(((eBaseSpd * eLVL) / 50) + 5)
    ElseIf pokeOwner = "p" Then
            If exp(pokeNum) > pLVL(pokeNum) ^ 3 Then
                'level up!
                pLVL(pokeNum) = pLVL(pokeNum) + 1
                passThis = pLVL(pokeNum)
                arrayNum = pokeNum
                Call form1.PassLVL
                lblYourLevel.Caption = pLVL(pokeNum)
            End If
            pHP = Int((((pBaseHP + 50) * pLVL(pokeNum)) / 50) + 10)
            lblMaxHP.Caption = pHP
            pAtk = Int(((pBaseATK * pLVL(pokeNum)) / 50) + 5)
            pDef = Int(((pBaseDef * pLVL(pokeNum)) / 50) + 5)
            pSpAtk = Int(((pBaseSpAtk * pLVL(pokeNum)) / 50) + 5)
            pSpDef = Int(((pBaseSpDef * pLVL(pokeNum)) / 50) + 5)
            pSpd = Int(((pBaseSpd * pLVL(pokeNum)) / 50) + 5)
    End If
End Sub

Public Sub getAttack()
    Select Case atkName
        Case "NOATTACK"
            power = 0
            atkType = "NORMAL"
            buff = ""
            debuff = ""
            accuracy = 0
            priority = 0
        Case "SCRATCH"
            power = 40
            atkType = "NORMAL"
            buff = ""
            debuff = ""
            accuracy = 100
            priority = 0
        Case "TACKLE"
            power = 35
            atkType = "NORMAL"
            buff = ""
            debuff = ""
            accuracy = 95
            priority = 1
        Case "TAILWHIP"
            power = 0
            atkType = "NORMAL"
            buff = ""
            debuff = "DEFENSE"
            accuracy = 100
            priority = 0
        Case "GROWL"
            power = 0
            atkType = "NORMAL"
            buff = ""
            debuff = "ATTACK1"
            accuracy = 100
            priority = 0
        Case "GUST"
            power = 40
            atkType = "FLYING"
            buff = ""
            debuff = ""
            accuracy = 100
            priority = 0
        Case "EMBER"
            power = 40
            atkType = "FIRE"
            buff = ""
            debuff = ""
            accuracy = 100
            priority = 0
        Case "SANDATTACK"
            power = 0
            atkType = "NORMAL"
            buff = ""
            debuff = "ACCURACY"
            accuracy = 100
            priority = 0
        Case "SMOKESCREEN"
            power = 0
            atkType = "NORMAL"
            buff = ""
            debuff = "ACCURACY"
            accuracy = 100
            priority = 0
        Case "QUICKATTACK"
            power = 0
            atkType = "NORMAL"
            buff = ""
            debuff = "ACCURACY"
            accuracy = 100
            priority = 1
        Case "POISONSTING"
            power = 40
            atkType = "POISON"
            buff = ""
            debuff = ""
            accuracy = 100
            priority = 0
        Case "STRINGSHOT"
            power = 0
            atkType = "NORMAL"
            buff = ""
            debuff = "SPEED"
            accuracy = 100
            priority = 0
        Case "HARDEN"
            power = 0
            atkType = "NORMAL"
            buff = "DEFENSE"
            debuff = ""
            accuracy = 100
            priority = 0
        Case "DEFENSECURL"
            power = 0
            atkType = "NORMAL"
            buff = "DEFENSE"
            debuff = ""
            accuracy = 100
            priority = 0
        Case "THUNDERSHOCK"
            power = 40
            atkType = "ELECTRIC"
            buff = ""
            debuff = "SPEED"
            accuracy = 100
            priority = 0
        Case "THUNDERWAVE"
            power = 40
            atkType = "ELECTRIC"
            buff = ""
            debuff = "" 'TODO
            accuracy = 100
            priority = 0
    End Select
End Sub

Private Sub compareTypes()
'compare typing to calculate supper effective damage
    Select Case compare
        Case "PVE1"
            Select Case atkType
                Case "FIRE"
                    If eType1 = "BUG" Or eType1 = "GRASS" Or eType1 = "ICE" Then
                        typeBonus1 = 2
                    End If
                Case "FLYING"
                    If eType1 = "BUG" Or eType1 = "GRASS" Or eType1 = "FIGHTING" Then
                        typeBonus1 = 2
                    End If
            End Select
        Case "PVE2"
            Select Case atkType
                Case "FIRE"
                    If eType2 = "BUG" Or eType2 = "GRASS" Or eType2 = "ICE" Then
                        typeBonus2 = 2
                    End If
                Case "FLYING"
                    If eType2 = "BUG" Or eType2 = "GRASS" Or eType2 = "FIGHTING" Then
                        typeBonus2 = 2
                    End If
            End Select
        Case "EVP1"
        Case "EVP2"
    End Select
End Sub

Public Sub battle()

    Randomize
    
    'speed check
    'TODO remember to work in priority
    'make sure we're attacking
    If atkName = "NOATTACK" Then
        tempSpd = pSpd
        pSpd = 999
    End If
    If pSpd > eSpd Or pSpd = eSpd Then
        'you attack first
        'check for STAB
        If atkType = pType1 Or atkType = pType2 Then
            STAB = 1.5
        Else:
            STAB = 1
        End If
        'YOU ATTACK
        'Accuracy Roll
        p = accuracy * pAccuracy
        'calculate super effective damage
        compare = "PVE1"
        Call compareTypes
        compare = "PVE2"
        Call compareTypes
        If p > (Rnd * 100) Then
            If power > 0 Then
                'deal non zero damage
                damage = (((((((pLVL(pokeNum) * 2 / 5) + 2) * power * pAtk / 50) / eDef) * BRN) + 2) * CH * (99 - (Rnd * 15)) / 100) * STAB * typeBonus1 * typeBonus2
                eHP = eHP - Int(damage)
                Else:
                damage = 0
            End If
            'debuff
            If Not debuff = "" Then
                Select Case debuff
                    Case "ATTACK1"
                        eAtk = eAtk * 0.8
                    Case "ACCURACY"
                        eAccuracy = eAccuracy * 0.8
                    Case "DEFENSE"
                        eDef = eDef * 0.8
                    Case "SPEED"
                        eDef = eDef * 0.8
                End Select
            End If
            If Not buff = "" Then
                Select Case buff
                    Case "DEFENSE"
                        pDef = pDef * 0.8
                End Select
            End If
            List1.AddItem pPoke & " used " & atkName
            If typeBonus1 * typeBonus2 > 1 Then
                List1.AddItem "It's super effective!"
            End If
            List1.AddItem ePoke & " has " & eHP & " HP"
            'reset values
            typeBonus1 = 1
            typeBonus2 = 1
            'is battle over?
            If eHP < 1 Then
                exp(pokeNum) = Int(exp(pokeNum) + ((baseEXP * eLVL * tEXP) / 7))
                cmdAtk1.Caption = "DONE"
                cmdAtk2.Caption = ""
                cmdAtk3.Caption = ""
                cmdAtk4.Caption = ""
                List1.AddItem ePoke & " has fainted."
                List1.AddItem pPoke & " has " & exp(pokeNum) & " exp!"
                If exp(pokeNum) > pLVL(pokeNum) ^ 3 Then
                    List1.AddItem pPoke & " has gained a level!"
                End If
                inBattle = False
            End If
        Else:
            List1.AddItem pPoke & " missed!"
        End If
        atkName = ""
        'make sure battle hasn't ended
        If inBattle = True Then
        'enemy attack
        Do Until Not atkName = ""
        rndAtk = Int((Rnd * 4) + 1)
        Select Case rndAtk
            Case 1
                atkName = eAtkName1
            Case 2
                atkName = eAtkName2
            Case 3
                atkName = eAtkName3
            Case 4
                atkName = eAtkName4
        End Select
        Loop
        Call getAttack
        'Enemy Attacks second
        'check for STAB
        If atkType = eType1 Or atkType = eType2 Then
            STAB = 1.5
        Else:
            STAB = 1
        End If
        'ENEMY ATTACK
        'Accuracy Roll
        p = accuracy * eAccuracy
        If p > (Rnd * 100) Then
            If power > 0 Then
                'deal damage
                damage = (((((((eLVL * 2 / 5) + 2) * power * eAtk / 50) / pDef) * BRN) + 2) * CH * (99 - (Rnd * 15)) / 100) * STAB
                         '* Type1 * Type2), for when I put in super effective/not very effective
                pHP = pHP - Int(damage)
                lblpHP.Caption = pHP
            Else: damage = 0
            End If
            'debuff
            If Not debuff = "" Then
                Select Case debuff
                    Case "ATTACK1"
                        pAtk = pAtk * 0.8
                    Case "ACCURACY"
                        pAccuracy = pAccuracy * 0.8
                    Case "DEFENSE"
                        pDef = pDef * 0.8
                    Case "SPEED"
                        pSpd = pSpd * 0.8
                End Select
            End If
            If Not buff = "" Then
                Select Case buff
                    Case "DEFENSE"
                        eDef = eDef * 0.8
                End Select
            End If
            List1.AddItem ePoke & " used " & atkName
            List1.AddItem pPoke & " has " & pHP & " HP"
            'is battle over?
            If pHP < 1 Then
                List1.AddItem pPoke & " has fainted."
                passThis = pHP
                arrayNum = pokeNum
                Call form1.ArrayFixer
                For i = 1 To 3
                    pokeNum = i
                    arrayNum = i
                    pokeOwner = "p"
                    Call getPPokemon
                    Call getStats
                    Call form1.GetHP
                    If form1.ovwHP > 0 Then
                        cmdAtk1.Caption = "Charmander"
                        cmdAtk2.Caption = "Poke2"
                        cmdAtk3.Caption = "Poke3"
                        cmdAtk4.Caption = "Done"
                    End If
                Next i
                If Not cmdAtk1.Caption = "Charmander" Then
                    List1.Clear
                    frmBattle.Hide
                    form1.Show
                    frmTxtBox.Show
                    frmTxtBox.List1.AddItem "You whited out! Your pokemon have healed."
                    tEXP = 1
                    canCatch = True
                    For x = 1 To frmBattle.partyNum
                        frmBattle.pokeNum = x
                        frmBattle.pokeOwner = "p"
                        Call getPPokemon
                        Call getStats
                        arrayNum = x
                        passThis = lblMaxHP
                        Call form1.ArrayFixer
                    Next x
                End If
            End If
        Else:
            List1.AddItem pPoke & " missed!"
        End If

        atkName = ""
        'back to menu options
        cmdAtk1.Caption = "Attack"
        cmdAtk2.Caption = "Items"
        cmdAtk3.Caption = "Pokemon"
        cmdAtk4.Caption = "Run"
    End If
    Else:
        atkName = ""
        'if opponent is faster
        'they attack first
        'choose enemy attack
        Do Until Not atkName = ""
        rndAtk = Int((Rnd * 4) + 1)
        Select Case rndAtk
            Case 1
                atkName = eAtkName1
            Case 2
                atkName = eAtkName2
            Case 3
                atkName = eAtkName3
            Case 4
                atkName = eAtkName4
        End Select
        Loop
        Call getAttack
        'check for STAB
        If atkType = eType1 Or etkType = eType2 Then
            STAB = 1.5
        Else:
            STAB = 1
        End If
        'Accuracy Roll
        p = accuracy * eAccuracy
        If p > (Rnd * 100) Then
            If power > 0 Then
                'deal damage
                damage = (((((((eLVL * 2 / 5) + 2) * power * eAtk / 50) / pDef) * BRN) + 2) * CH * (99 - (Rnd * 15)) / 100) * STAB
                         '* Type1 * Type2), for when I put in super effective/not very effective
                pHP = pHP - Int(damage)
                lblpHP.Caption = pHP
            Else:
                damage = 0
            End If
            'debuff
            If Not debuff = "" Then
                Select Case debuff
                    Case "ATTACK1"
                        pAtk = pAtk * 0.8
                    Case "ACCURACY"
                        pAccuracy = pAccuracy * 0.8
                    Case "DEFENSE"
                        pDef = pDef * 0.8
                    Case "SPEED"
                        pSpd = pSpd * 0.8
                End Select
            End If
            If Not buff = "" Then
                Select Case buff
                    Case "DEFENSE"
                        eDef = pDef * 0.8
                End Select
            End If
            List1.AddItem ePoke & " used " & atkName
            List1.AddItem pPoke & " has " & pHP & " HP!"
            'is battle over?
            If pHP < 1 Then
                List1.AddItem pPoke & " has fainted."
                passThis = pHP
                arrayNum = pokeNum
                Call form1.ArrayFixer
                For i = 1 To 3
                    pokeNum = i
                    arrayNum = i
                    pokeOwner = "p"
                    Call getPPokemon
                    Call getStats
                    Call form1.GetHP
                    If form1.ovwHP > 0 Then
                        cmdAtk1.Caption = "Charmander"
                        cmdAtk2.Caption = "Poke2"
                        cmdAtk3.Caption = "Poke3"
                        cmdAtk4.Caption = "Done"
                    End If
                Next i
                If Not cmdAtk1.Caption = "Charmander" Then
                    List1.Clear
                    frmBattle.Hide
                    form1.Show
                    frmTxtBox.Show
                    frmTxtBox.List1.AddItem "You whited out! Your pokemon have healed."
                    tEXP = 1
                    canCatch = True
                    For x = 1 To frmBattle.partyNum
                        frmBattle.pokeNum = x
                        frmBattle.pokeOwner = "p"
                        Call getPPokemon
                        Call getStats
                        arrayNum = x
                        passThis = lblMaxHP
                        Call form1.ArrayFixer
                    Next x
                End If
            End If
        Else:
            List1.AddItem ePoke & " missed!"
        End If
        atkName = ""
        
        'You Attack second
        'get attack
        atkName = pAtkName(atkNum, pokeNum)
        Call getAttack
        'check for STAB
        If atkType = pType1 Or atkType = pType2 Then
            STAB = 1.5
        Else:
            STAB = 1
        End If
        'check for super effective
        compare = "PVE1"
        Call compareTypes
        compare = "PVE2"
        Call compareTypes
        'YOU ATTACK
        'Accuracy Roll
        p = accuracy * pAccuracy
        If p > (Rnd * 100) Then
            If power > 0 Then
                'deal damage
                damage = (((((((pLVL(pokeNum) * 2 / 5) + 2) * power * pAtk / 50) / eDef) * BRN) + 2) * CH * (99 - (Rnd * 15)) / 100) * STAB * typeBonus1 * typeBonus2
                eHP = eHP - Int(damage)
            Else:
                damage = 0
            End If
            'debuff
            If Not debuff = "" Then
                Select Case debuff
                    Case "ATTACK1"
                        eAtk = eAtk * 0.8
                    Case "ACCURACY"
                        eAccuracy = eAccuracy * 0.8
                    Case "DEFENSE"
                        eDef = eDef * 0.8
                    Case "SPEED"
                        eSpd = eSpd * 0.8
                End Select
            End If
            If Not buff = "" Then
                Select Case buff
                    Case "DEFENSE"
                        pDef = pDef * 0.8
                End Select
            End If
            List1.AddItem pPoke & " used " & atkName
            If typeBonus1 * typeBonus2 > 1 Then
                List1.AddItem "It's super effective!"
            End If
            List1.AddItem ePoke & " has " & eHP & " HP"
            'set back to defaults
            typeBonus1 = 1
            typeBonus2 = 1
            'is battle over?
            If eHP < 1 Then
                exp(pokeNum) = Int(exp(pokeNum) + ((baseEXP * eLVL * tEXP) / 7))
                cmdAtk1.Caption = "DONE"
                cmdAtk2.Caption = ""
                cmdAtk3.Caption = ""
                cmdAtk4.Caption = ""
                List1.AddItem ePoke & " has fainted."
                List1.AddItem pPoke & " has " & exp(pokeNum) & " exp!"
                If exp(pokeNum) > pLVL(pokeNum) ^ 3 Then
                    List1.AddItem pPoke & " has gained a level!"
                End If
                inBattle = False
            End If
        Else:
            List1.AddItem pPoke & " missed!"
        End If
        atkName = ""
        If inBattle = True Then
            'back to menu options
            cmdAtk1.Caption = "Attack"
            cmdAtk2.Caption = "Items"
            cmdAtk3.Caption = "Pokemon"
            cmdAtk4.Caption = "Run"
        End If
    End If
    If pSpd = 999 Then
        pSpd = tempSpd
    End If
    
    lblpHP = pHP
End Sub

Private Sub cmdAtk1_Click()
    If cmdAtk1.Caption = "Attack" Then
        cmdAtk1.Caption = pAtkName(1, pokeNum)
        cmdAtk2.Caption = pAtkName(2, pokeNum)
        cmdAtk3.Caption = pAtkName(3, pokeNum)
        cmdAtk4.Caption = pAtkName(4, pokeNum)
        trainerBattle.Visible = False
    ElseIf cmdAtk1.Caption = "Start" Then
        trainerBattle.Visible = False
        cmdAtk1.Caption = "Attack"
        cmdAtk2.Caption = "Items"
        cmdAtk3.Caption = "Pokemon"
        cmdAtk4.Caption = "Run"
    ElseIf cmdAtk1.Caption = "Poke Ball" Then
        If fullParty = False And canCatch = True Then
            'calculate the probability of catching the pokemon
            a = ((3 * maxeHP - 2 * eHP) * catchRate) / (3 * maxeHP)
            probCapture = a / 255
            If a >= 255 Then
                'you capture the pokemon!
                partyNum = pokeNum + 1
                If partyNum = 2 Then 'MAX2
                    fullParty = True
                End If
                passThis = eLVL
                arrayNum = partyNum
                Call form1.PassLVL
                pPokeNum(partyNum) = index
                exp(pokeNum) = eLVL ^ 3
                passThis = maxeHP
                arrayNum = partyNum
                Call form1.ArrayFixer
                cmdAtk1.Caption = "DONE"
                cmdAtk2.Caption = ""
                cmdAtk3.Caption = ""
                cmdAtk4.Caption = ""
                List1.Clear
                List1.AddItem "You caught " & ePoke & " !"
                inBattle = False
            ElseIf a < 255 Then
                If probCapture > Rnd Then
                    'you capture the pokemon!
                    partyNum = partyNum + 1
                    If partyNum = 3 Then 'MAX3
                        fullParty = True
                    End If
                    passThis = eLVL
                    arrayNum = partyNum
                    Call form1.PassLVL
                    pPokeNum(partyNum) = index
                    exp(pokeNum) = eLVL ^ 3
                    passThis = maxeHP
                    arrayNum = partyNum
                    Call form1.ArrayFixer
                    inBattle = False
                    cmdAtk1.Caption = "DONE"
                    cmdAtk2.Caption = ""
                    cmdAtk3.Caption = ""
                    cmdAtk4.Caption = ""
                    List1.Clear
                    List1.AddItem "You caught " & ePoke & " !"
                Else:
                    'you fail to catch the pokemon
                    List1.Clear
                    List1.AddItem "Shoot! And it was so close too..."
                    atkName = "NOATTACK"
                    Call getAttack
                    Call battle
                End If
            End If
        Else
            'you can't catch any more pokemon, return to menu
            List1.AddItem "YOUR PARTY IS FULL"
            cmdAtk1.Caption = "Attack"
            cmdAtk2.Caption = "Items"
            cmdAtk3.Caption = "Pokemon"
            cmdAtk4.Caption = "Run"
        End If
    ElseIf cmdAtk1.Caption = "poke1" Then
            If form1.ovwHP > 0 Then
                'save pokemon's hp
                passThis = pHP
                arrayNum = pokeNum
                Call form1.ArrayFixer
                'change your pokemon
                pokeNum = 1
                pHP = form1.ovwHP
                pokeOwner = "p"
                Call getPPokemon
                Call getStats
                List1.Clear
                atkName = "NOATTACK"
                Call getAttack
                Call battle
            Else:
                cmdAtk1.Caption = "Attack"
                cmdAtk2.Caption = "Items"
                cmdAtk3.Caption = "Pokemon"
                cmdAtk4.Caption = "Run"
            End If
    Else:
        If inBattle = True Then
            List1.Clear
            atkName = pAtkName(1, pokeNum)
            atkNum = 1
            Call getAttack
            Call battle
        Else:
            'end the battle
            List1.Clear
            passThis = pHP
            arrayNum = pokeNum
            Call form1.ArrayFixer
            inBattle = True
            'set back to defaults
            cmdAtk1.Caption = "Attack"
            cmdAtk2.Caption = "Items"
            cmdAtk3.Caption = "Pokemon"
            cmdAtk4.Caption = "Run"
            frmBattle.Hide
            frmTxtBox.Show
            tEXP = 1
            canCatch = True
            Select Case form1.MAP
            Case "PALLET"
                form1.Show
                frmTxtBox.List1.AddItem "Pallet Town"
            Case "ROUTE1"
                frmRoute1.Show
                frmTxtBox.List1.AddItem "Route 1"
            Case "VIRIDIANFOREST"
                frmViridianForest.Show
                frmTxtBox.List1.AddItem "Viridian Forest"
                'call function to get the trainer's next pokemon
                Call frmViridianForest.get_Battle
            End Select
            'the next time you enter a battle, you want the first pokemon in your party, and form1.getLVL will only call from this frm
            arrayNum = 1
        End If
    End If
End Sub

Private Sub cmdAtk2_Click()
    If cmdAtk2.Caption = "Items" Then
        cmdAtk1.Caption = "Poke Ball"
        cmdAtk2.Caption = "Potion"
        cmdAtk3.Caption = ""
        cmdAtk4.Caption = "Done"
    ElseIf cmdAtk2.Caption = "Potion" Then
        pHP = pHP + 20
        If pHP > Int(lblMaxHP.Caption) Then
            pHP = Int(lblMaxHP.Caption)
        End If
        lblpHP.Caption = pHP
    ElseIf cmdAtk2.Caption = "Poke2" Then
        arrayNum = 2
        Call form1.GetHP
        If form1.ovwHP > 0 And partyNum > 1 Then
            'save pokemon's hp
            passThis = pHP
            arrayNum = pokeNum
            Call form1.ArrayFixer
            'change your pokemon
            pokeNum = 2
            arrayNum = 2
            pokeOwner = "p"
            Call getPPokemon
            Call getStats
            Call form1.GetHP
            pHP = form1.ovwHP
            List1.Clear
            atkName = "NOATTACK"
            Call getAttack
            Call battle
        Else:
            cmdAtk1.Caption = "Attack"
            cmdAtk2.Caption = "Items"
            cmdAtk3.Caption = "Pokemon"
            cmdAtk4.Caption = "Run"
        End If
    Else:
        If Not pAtkName(2, pokeNum) = "" Then
            List1.Clear
            atkName = pAtkName(2, pokeNum)
            atkNum = 2
            Call getAttack
            Call battle
        End If
    End If
End Sub

Private Sub cmdAtk3_Click()
    If cmdAtk3.Caption = "Pokemon" Then
        cmdAtk1.Caption = "Charmander"
        cmdAtk2.Caption = "Poke2"
        cmdAtk3.Caption = "Poke3"
        cmdAtk4.Caption = "Done"
    ElseIf cmdAtk3.Caption = "Poke3" Then
        arrayNum = 3
        Call form1.GetHP
        If form1.ovwHP > 0 And partyNum > 1 Then
            'save pokemon's hp
            passThis = pHP
            arrayNum = pokeNum
            Call form1.ArrayFixer
            'change your pokemon
            pokeNum = 3
            arrayNum = 3
            pokeOwner = "p"
            Call getPPokemon
            Call getStats
            Call form1.GetHP
            pHP = form1.ovwHP
            List1.Clear
            atkName = "NOATTACK"
            Call getAttack
            Call battle
        End If
    Else:
        If Not pAtkName(3, pokeNum) = "" Then
            List1.Clear
            atkName = pAtkName(3, pokeNum)
            atkNum = 3
            Call getAttack
            Call battle
        End If
    End If
End Sub

Private Sub cmdAtk4_Click()
    If cmdAtk4.Caption = "Run" Then
        cmdAtk1.Caption = "DONE"
        cmdAtk2.Caption = ""
        cmdAtk3.Caption = ""
        cmdAtk4.Caption = ""
        inBattle = False
    ElseIf cmdAtk4.Caption = "Done" Then
        cmdAtk1.Caption = "Attack"
        cmdAtk2.Caption = "Items"
        cmdAtk3.Caption = "Pokemon"
        cmdAtk4.Caption = "Run"
    Else:
        If Not pAtkName(4, pokeNum) = "" Then
            List1.Clear
            atkName = pAtkName(4, pokeNum)
            atkNum = 4
            Call getAttack
            Call battle
        End If
    End If
End Sub

Public Sub passpLVL()
    pLVL(arrayNum) = passThis
End Sub
