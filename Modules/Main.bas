Attribute VB_Name = "Main"
Option Explicit

Sub Main()

    ' 1. TEST DU DECK
    
    '// Déclaration et initialisation :
    Dim jeu1 As deck
    Set jeu1 = New deck
    '// Test des fonctions
    jeu1.Melanger
    'Range("A1") = jeu1.Taille
    'Range("B1") = jeu1.Piocher.NomComplet
    
    ' 2. TEST DE LA MAINDEPOKER :
    '// Déclaration et initialisation :
    Dim table1 As Table
    Dim j, j1, j2, j3 As Joueur
    '// Initialisation des nouveaux joueurs :
    Set j1 = New Joueur
    Set j2 = New Joueur
    Set j3 = New Joueur

    j1.Init "Marc", "BTN", 25
    j2.Init "Benjamin", "SB", 25
    j3.Init "Maxime", "BB", 25
    
    '// Initialisation de la table de jeu :
    Set table1 = New Table
    table1.Init j1, j2, j3
    
    
    '// Distribution des cartes :
    table1.DistribuerTable
    
    '// Affichage de la table :
    Range("A1") = table1.AfficherTable
    Dim i As Integer
    i = 0
    For Each j In table1.table3Joueurs
        Range("A3").Offset(0, i) = j.Nom
        Range("A4").Offset(0, i) = j.Position
        Range("A5").Offset(0, i) = j.Stack
        Range("A6").Offset(0, i) = j.MainDePoker.AfficherMain()
        i = i + 1
    Next j
    
    
    
End Sub
