VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FmTableArticles2 
   Caption         =   "Libelle de la table"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15720
   OleObjectBlob   =   "FmTableArticles2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FmTableArticles2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAjouter_Click()
AjoutModif = True
FmSaisieArticle.Show
Me.ListTable.Clear
Call UserForm_Initialize
End Sub

Private Sub CmdModifier_Click()
AjoutModif = False
LigneAModifier = -1
If Me.ListTable.ListIndex = -1 Then
    MsgBox "Veuillez sélectionner l'information à modifier", vbCritical
    Else
    Dim i As Long
    i = 1
    Do
    i = i + 1
        If CStr(Me.ListTable.Value) = CStr(FeuilArticles.Range("A" & i).Value) Then
             LigneAModifier = i
             
          Exit Do
        End If
    Loop Until FeuilArticles.Range("A" & i).Value = ""
End If
      If LigneAModifier = -1 Then
         MsgBox "L'information a modifier n'a été retrouvée dans la base de données !", vbCritical
         Else
         Dim j As Byte
         For j = 1 To 6
            FmSaisieArticle.Controls("Textbox" & j).Value = FeuilArticles.Cells(LigneAModifier, j).Value
         Next j
            FmSaisieArticle.Show
      End If
      Me.ListTable.Clear
      Call UserForm_Initialize
End Sub

Private Sub CmdSupprimer_Click()
LigneAModifier = -1
If Me.ListTable.ListIndex = -1 Then
    MsgBox "Veuillez sélectionner l'information à supprimer", vbCritical
    Else
    Dim i As Long
    i = 1
    Do
    i = i + 1
        If CStr(Me.ListTable.Value) = CStr(FeuilArticles.Range("A" & i).Value) Then
             LigneAModifier = i
             
          Exit Do
        End If
    Loop Until FeuilArticles.Range("A" & i).Value = ""
End If
If Not (LigneAModifier = -1) Then
   If MsgBox("Souhaitez-vous supprimer cet enregistrement ?", vbQuestion + vbYesNo) = vbYes Then
       FeuilArticles.Rows(LigneAModifier).Delete
   End If
  Else
    MsgBox "Aucun enregistrement sélectionné !", vbCritical
End If
'Actualisation de la table de données
      Me.ListTable.Clear
      Call UserForm_Initialize
End Sub

Private Sub TxRechercher_Change()
Me.ListTable.Clear
If Me.TxRechercher.Value = "" Then
    Call UserForm_Initialize
    Else
    'Afficher les données
        Dim j As Long
        j = 1
        Do
           j = j + 1
           If Not (FeuilArticles.Range("A" & j).Value = "") Then
              If CStr(UCase((FeuilArticles.Range("A" & j).Value & FeuilArticles.Range("B" & j).Value & FeuilArticles.Range("C" & j).Value & FeuilArticles.Range("D" & j).Value & FeuilArticles.Range("E" & j).Value))) Like "*" & CStr(UCase(Me.TxRechercher.Value)) & "*" Then
                    Me.ListTable.AddItem FeuilArticles.Range("A" & j).Value
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 1) = FeuilArticles.Range("B" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 2) = FeuilArticles.Range("C" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 3) = FeuilArticles.Range("D" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 4) = FeuilArticles.Range("F" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 5) = FeuilArticles.Range("G" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 6) = FeuilArticles.Range("I" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 7) = FeuilArticles.Range("J" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 8) = FeuilArticles.Range("K" & j)
               End If
           End If
           
        Loop Until FeuilArticles.Range("A" & j).Value = ""
        End If
End Sub

Private Sub UserForm_Initialize()

'Afficher les titre
Dim i As Byte
For i = 1 To 11
 If i = 5 Or i = 8 Then
    
    Else
    Me.Controls("LibTitre" & i).Caption = FeuilArticles.Cells(1, i).Text
 End If
   
Next
'Afficher les données
Dim j As Long
j = 1
Do
   j = j + 1
   If Not (FeuilArticles.Range("A" & j).Value = "") Then
       Me.ListTable.AddItem FeuilArticles.Range("A" & j).Value
       Me.ListTable.List(Me.ListTable.ListCount - 1, 1) = FeuilArticles.Range("B" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 2) = FeuilArticles.Range("C" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 3) = FeuilArticles.Range("D" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 4) = FeuilArticles.Range("F" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 5) = FeuilArticles.Range("G" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 6) = FeuilArticles.Range("I" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 7) = FeuilArticles.Range("J" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 8) = FeuilArticles.Range("K" & j)
   End If
   
Loop Until FeuilArticles.Range("A" & j).Value = ""
End Sub
