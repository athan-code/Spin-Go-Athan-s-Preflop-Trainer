VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Joueur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1. Déclaration des variables de classe :
Private pNom As String
Private pPosition As String
Private pStack As Double
Private pMainDePoker As MainDePoker
'Private pEstActif As Boolean
'Private pEstAllIn As Boolean
'Private pderniereAction As String
'Private pmontantMise As Double
Private pisHero As Boolean

' 2. Méthodes publiques :

' // 2.1. Constructeur :
Public Sub Init(ByVal nomJ As String, ByVal pos As String, ByVal stackBB As Double, Optional ByVal hero As Boolean = False)
    pNom = nomJ
    pPosition = pos
    pStack = stackBB
    pisHero = hero
End Sub

' 3. Propriétés publiques :
Public Property Get Nom() As String
    Nom = pNom
End Property

Public Property Let Nom(ByVal v As String)
    pNom = v
End Property

Public Property Get Position() As String
    Position = pPosition
End Property

Public Property Let Position(ByVal v As String)
    pPosition = v
End Property

Public Property Get Stack() As Double
    Stack = pStack
End Property

Public Property Let Stack(ByVal v As Double)
    pStack = v
End Property

Public Property Get MainDePoker() As MainDePoker
    Set MainDePoker = pMainDePoker
End Property

Public Property Set MainDePoker(ByVal m As MainDePoker)
    Set pMainDePoker = m
End Property
