VERSION 5.00
Begin VB.Form frmTxtBox 
   Caption         =   "Form2"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   ScaleHeight     =   2070
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdButton6 
      Caption         =   "Button 6"
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdButton5 
      Caption         =   "Button 5"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdButton4 
      Caption         =   "Button 4"
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdButton3 
      Caption         =   "Button 3"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "Pokemon"
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdButton1 
      Caption         =   "Items"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "frmTxtBox.frx":0000
      Left            =   0
      List            =   "frmTxtBox.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmTxtBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdButton3_Click()
    If cmdButton3.Caption = "Pokemon 3" Then
        If frmBattle.partyNum > 2 Then
            List1.Clear
            frmBattle.pokeNum = 3
            frmBattle.pokeOwner = "p"
            Call frmBattle.getPPokemon
            Call frmBattle.getStats
            frmBattle.arrayNum = 3
            Call form1.GetHP
            List1.AddItem frmBattle.pPoke
            List1.AddItem "HP = " & form1.ovwHP & "/" & frmBattle.lblMaxHP
            List1.AddItem "Attack = " & frmBattle.pAtk
            List1.AddItem "Defense = " & frmBattle.pDef
            List1.AddItem "Special Attack = " & frmBattle.pSpAtk
            List1.AddItem "Special Defense = " & frmBattle.pSpDef
            List1.AddItem "Speed = " & frmBattle.pSpd
            cmdButton3.Visible = False
            cmdButton4.Visible = False
            cmdButton5.Visible = False
            cmdButton6.Visible = False
            cmdExit.Visible = False
            cmdButton1.Caption = "Make Party Lead"
            cmdButton2.Caption = "Done"
        End If
    End If
End Sub

Private Sub cmdButton5_Click()
    If cmdButton5.Caption = "Pokemon 1" Then
        List1.Clear
        frmBattle.pokeNum = 1
        frmBattle.pokeOwner = "p"
        Call frmBattle.getPPokemon
        Call frmBattle.getStats
        frmBattle.arrayNum = 1
        Call form1.GetHP
        List1.AddItem frmBattle.pPoke
        List1.AddItem "HP = " & form1.ovwHP & "/" & frmBattle.lblMaxHP
        List1.AddItem "Attack = " & frmBattle.pAtk
        List1.AddItem "Defense = " & frmBattle.pDef
        List1.AddItem "Special Attack = " & frmBattle.pSpAtk
        List1.AddItem "Special Defense = " & frmBattle.pSpDef
        List1.AddItem "Speed = " & frmBattle.pSpd
        cmdButton3.Visible = False
        cmdButton4.Visible = False
        cmdButton5.Visible = False
        cmdButton6.Visible = False
        cmdExit.Visible = False
        cmdButton1.Caption = "Make Party Lead"
        cmdButton2.Caption = "Done"
    End If
End Sub

Private Sub cmdButton6_Click()
    If cmdButton6.Caption = "Pokemon 2" Then
        If frmBattle.partyNum > 1 Then
            List1.Clear
            frmBattle.pokeNum = 2
            frmBattle.pokeOwner = "p"
            Call frmBattle.getPPokemon
            Call frmBattle.getStats
            frmBattle.arrayNum = 2
            Call form1.GetHP
            List1.AddItem frmBattle.pPoke
            List1.AddItem "HP = " & form1.ovwHP & "/" & frmBattle.lblMaxHP
            List1.AddItem "Attack = " & frmBattle.pAtk
            List1.AddItem "Defense = " & frmBattle.pDef
            List1.AddItem "Special Attack = " & frmBattle.pSpAtk
            List1.AddItem "Special Defense = " & frmBattle.pSpDef
            List1.AddItem "Speed = " & frmBattle.pSpd
            cmdButton3.Visible = False
            cmdButton4.Visible = False
            cmdButton5.Visible = False
            cmdButton6.Visible = False
            cmdExit.Visible = False
            cmdButton1.Caption = "Make Party Lead"
            cmdButton2.Caption = "Done"
        End If
    End If
End Sub

Private Sub Form_Load()
    frmTxtBox.Top = 8350
    frmTxtBox.Left = 2000
End Sub

Private Sub cmdButton1_Click()
    Select Case cmdButton1.Caption
    Case "Items"
        cmdExit.Visible = True
        cmdButton1.Caption = "Potions"
        cmdButton2.Caption = "Poke Balls"
    End Select
End Sub

Private Sub cmdButton2_Click()
    If cmdButton2.Caption = "Pokemon" Then
        cmdButton3.Visible = True
        cmdButton4.Visible = True
        cmdButton5.Visible = True
        cmdButton6.Visible = True
        cmdExit.Visible = True
        cmdButton5.Caption = "Pokemon 1"
        cmdButton6.Caption = "Pokemon 2"
        cmdButton3.Caption = "Pokemon 3"
        cmdButton4.Caption = "Pokemon 4"
        cmdButton1.Caption = "Pokemon 5"
        cmdButton2.Caption = "Pokemon 6"
    ElseIf cmdButton2.Caption = "Done" Then
        List1.Clear
        cmdButton3.Visible = False
        cmdButton4.Visible = False
        cmdButton5.Visible = False
        cmdButton6.Visible = False
        cmdExit.Visible = False
        cmdButton1.Caption = "Items"
        cmdButton2.Caption = "Pokemon"
    End If
End Sub

Private Sub cmdExit_Click()
    cmdButton3.Visible = False
    cmdButton4.Visible = False
    cmdButton5.Visible = False
    cmdButton6.Visible = False
    cmdExit.Visible = False
    cmdButton1.Caption = "Items"
    cmdButton2.Caption = "Pokemon"
End Sub

