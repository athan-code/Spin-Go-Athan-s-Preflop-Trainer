VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1. Déclaration des variables de classe :

Private joueur1 As Joueur
Private joueur2 As Joueur
Private joueur3 As Joueur
Public table3Joueurs As Collection
Private deckDeTable As deck

' 2. Méthodes publiques :

' // 2.1. Constructeur :
Private Sub Class_Initialize()
    '// Initialisation de la collection des 3 joueurs :
    Set table3Joueurs = New Collection
    '// Initialisation du deck de carte de la table :
    Set deckDeTable = New deck
    deckDeTable.Melanger '// Le Deck est mélangé
End Sub
Public Sub Init(ByVal j1 As Joueur, ByVal j2 As Joueur, ByVal j3 As Joueur)
    Set joueur1 = j1
    Set joueur2 = j2
    Set joueur3 = j3
    
    '//Attribution des joueurs à la table si non déjà remplie :
    If table3Joueurs.Count = 0 Then
        table3Joueurs.Add joueur1
        table3Joueurs.Add joueur2
        table3Joueurs.Add joueur3
    Else:
        MsgBox "Joueurs déjà présents sur la table"
        AfficherTable
    End If
    
    MsgBox "Les 3 nouveaux joueurs sont assis à la nouvelle table."

End Sub

' 3. Propriétés publiques :
Public Sub DistribuerTable()
    '// Déclaration des mains :
    Dim m As MainDePoker
    Dim j As Joueur
    '// Pour chaque joueur de la table :
    For Each j In table3Joueurs
        Set m = New MainDePoker
        '// Pioche 2 cartes :
        m.Init deckDeTable.Piocher, deckDeTable.Piocher
        '// Let au joueur ces 2 cartes :
        Set j.MainDePoker = m
    Next j
End Sub


Public Function AfficherTable() As String
    Dim txt As String
    Dim j As Joueur
    For Each j In Me.table3Joueurs
        txt = txt & j.Nom() & " | " & j.Position & " | " & j.Stack & "  //  "
    Next j
    AfficherTable = txt
End Function
