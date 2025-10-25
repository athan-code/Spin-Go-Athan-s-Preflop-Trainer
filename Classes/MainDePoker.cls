VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainDePoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 1. Déclaration des variables de classe :

Private carte1 As carte
Private carte2 As carte
Public Main2Cartes As Collection

' 2. Méthodes publiques :

' // 2.1. Constructeur :
Private Sub Class_Initialize()
    Set Main2Cartes = New Collection
End Sub
Public Sub Init(ByVal c1 As carte, ByVal c2 As carte)
    Set carte1 = c1
    Set carte2 = c2
    If Not carte1 Is Nothing Then Main2Cartes.Add carte1
    If Not carte2 Is Nothing Then Main2Cartes.Add carte2
End Sub

' 3. Propriétés publiques :
Public Function AfficherMain() As String
    Dim txt As String
    Dim c As carte
    For Each c In Me.Main2Cartes
        txt = txt & c.NomComplet & " "
    Next c
    AfficherMain = Trim(txt)
End Function

Public Function TypeMain() As String

    ' // Vérification que la main est complète :
    If Me.Main2Cartes.Count <> 2 Then Exit Function
    
    ' // Suited :
    Dim suited As String
    suited = IIf(Me.Main2Cartes(1).couleur = Me.Main2Cartes(2).couleur, "s", "o")
    
    Dim a, b As Integer
    
    '// Si carte 1 = carte 2
    If Me.Main2Cartes(1).ValeurNumerique() = Me.Main2Cartes(2).ValeurNumerique() Then
    a = 2: b = 1: suited = ""
    '// Si carte 1 > carte 2 :
    ElseIf Me.Main2Cartes(1).ValeurNumerique() > Me.Main2Cartes(2).ValeurNumerique() Then a = 1: b = 2
    '// Si carte 1 < carte 2 :
    Else: a = 2: b = 1
    End If

    TypeMain = Me.Main2Cartes(a).valeur & Me.Main2Cartes(b).valeur & suited
    
End Function
