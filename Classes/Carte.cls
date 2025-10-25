VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Carte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1. Déclaration des variables de classe :

Private pValeur As String
Private pCouleur As String

' 2. Méthodes publiques :

' // 2.1. Constructeur :
Public Sub Init(ByVal v As String, ByVal c As String)
    pValeur = v
    pCouleur = c
End Sub

' 3. Propriétés publiques :

' // 3.1. Getter :
Public Property Get valeur() As String
    valeur = pValeur
End Property
Public Property Get couleur() As String
    couleur = pCouleur
End Property

' // 3.2. Letter :
Public Property Let valeur(ByVal v As String)
    valeur = pValeur
End Property
Public Property Let couleur(ByVal c As String)
    couleur = pCouleur
End Property

' // 3.3. Afficheur : nom complet de la carte pour affichage (ex: "As?")
Public Property Get NomComplet() As String
    NomComplet = pValeur & pCouleur
End Property

' // 3.4. Valeur numérique de la carte pour comparer les cartes (2–14)
Public Property Get ValeurNumerique() As Integer
    Select Case pValeur
        Case "2": ValeurNumerique = 2
        Case "3": ValeurNumerique = 3
        Case "4": ValeurNumerique = 4
        Case "5": ValeurNumerique = 5
        Case "6": ValeurNumerique = 6
        Case "7": ValeurNumerique = 7
        Case "8": ValeurNumerique = 8
        Case "9": ValeurNumerique = 9
        Case "T": ValeurNumerique = 10
        Case "J": ValeurNumerique = 11
        Case "Q": ValeurNumerique = 12
        Case "K": ValeurNumerique = 13
        Case "A": ValeurNumerique = 14
        Case Else: ValeurNumerique = 0
    End Select
End Property
