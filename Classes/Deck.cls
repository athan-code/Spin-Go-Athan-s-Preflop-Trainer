VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Deck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1. Déclaration des variables de classe :

Private CartesDuDeck As Collection

' 2. Méthodes publiques :

' // 2.1. Constructeur :
Private Sub Class_Initialize()
    Set CartesDuDeck = New Collection
    Call GenererDeck
End Sub

' // 2.2. Génération du deck :
Private Sub GenererDeck()
    Dim valeurs As Variant
    Dim couleurs As Variant
    Dim v As Variant, c As Variant
    Dim carte As carte
    
    valeurs = Array("2", "3", "4", "5", "6", "7", "8", "9", "T", "J", "Q", "K", "A")
    couleurs = Array(ChrW(9824), ChrW(9829), ChrW(9830), ChrW(9827))
    
    For Each v In valeurs
        For Each c In couleurs
            Set carte = New carte
            carte.Init v, c
            CartesDuDeck.Add carte
        Next c
    Next v
End Sub

' // 2.3. Mélanger le deck :
Public Sub Melanger()
    Dim temp As New Collection
    Dim i As Long, index As Long
    Dim carte As carte
    
    Randomize
    Do While CartesDuDeck.Count > 0
        index = Int(Rnd() * CartesDuDeck.Count) + 1
        Set carte = CartesDuDeck(index)
        temp.Add carte
        CartesDuDeck.Remove index
    Loop
    Set CartesDuDeck = temp
End Sub

' // 2.4. Distribuer une carte :
Public Function Piocher() As carte
    If CartesDuDeck.Count = 0 Then Exit Function
    Set Piocher = CartesDuDeck(1)
    CartesDuDeck.Remove 1 '// L'ensemble des index "descendent" vers
End Function


' // Nombre de cartes restantes
Public Property Get Taille() As Long
    Taille = CartesDuDeck.Count
End Property
