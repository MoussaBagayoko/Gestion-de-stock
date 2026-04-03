VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FmTableModeleFinal 
   Caption         =   "Libelle de la table"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12990
   OleObjectBlob   =   "FmTableModeleFinal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FmTableModeleFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAjouter_Click()
AjoutModif = True
FmSaisieFournisseur.Show
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
        If CStr(Me.ListTable.Value) = CStr(FeuilFournisseur.Range("A" & i).Value) Then
             LigneAModifier = i
             
          Exit Do
        End If
    Loop Until FeuilFournisseur.Range("A" & i).Value = ""
End If
      If LigneAModifier = -1 Then
         MsgBox "L'information a modifier n'a été retrouvée dans la base de données !", vbCritical
         Else
         Dim j As Byte
         For j = 1 To 6
            FmSaisieFournisseur.Controls("Textbox" & j).Value = FeuilFournisseur.Cells(LigneAModifier, j).Value
         Next j
            FmSaisieFournisseur.Show
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
        If CStr(Me.ListTable.Value) = CStr(FeuilFournisseur.Range("A" & i).Value) Then
             LigneAModifier = i
             
          Exit Do
        End If
    Loop Until FeuilFournisseur.Range("A" & i).Value = ""
End If
If Not (LigneAModifier = -1) Then
   If MsgBox("Souhaitez-vous supprimer cet enregistrement ?", vbQuestion + vbYesNo) = vbYes Then
       FeuilFournisseur.Rows(LigneAModifier).Delete
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
           If Not (FeuilFournisseur.Range("A" & j).Value = "") Then
              If CStr(UCase((FeuilFournisseur.Range("A" & j).Value & FeuilFournisseur.Range("B" & j).Value & FeuilFournisseur.Range("C" & j).Value & FeuilFournisseur.Range("D" & j).Value & FeuilFournisseur.Range("E" & j).Value))) Like "*" & CStr(UCase(Me.TxRechercher.Value)) & "*" Then
                    Me.ListTable.AddItem FeuilFournisseur.Range("A" & j).Value
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 1) = FeuilFournisseur.Range("B" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 2) = FeuilFournisseur.Range("C" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 3) = FeuilFournisseur.Range("D" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 4) = FeuilFournisseur.Range("E" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 5) = FeuilFournisseur.Range("F" & j)
                    Me.ListTable.List(Me.ListTable.ListCount - 1, 6) = FeuilFournisseur.Range("G" & j)
               End If
           End If
           
        Loop Until FeuilFournisseur.Range("A" & j).Value = ""
        End If
End Sub

Private Sub UserForm_Initialize()

'Afficher les titre
Dim i As Byte
For i = 1 To 6
    Me.Controls("LibTitre" & i).Caption = FeuilFournisseur.Cells(1, i).Text
Next
'Afficher les données
Dim j As Long
j = 1
Do
   j = j + 1
   If Not (FeuilFournisseur.Range("A" & j).Value = "") Then
       Me.ListTable.AddItem FeuilFournisseur.Range("A" & j).Value
       Me.ListTable.List(Me.ListTable.ListCount - 1, 1) = FeuilFournisseur.Range("B" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 2) = FeuilFournisseur.Range("C" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 3) = FeuilFournisseur.Range("D" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 4) = FeuilFournisseur.Range("E" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 5) = FeuilFournisseur.Range("F" & j)
       Me.ListTable.List(Me.ListTable.ListCount - 1, 6) = FeuilFournisseur.Range("G" & j)
   End If
   
Loop Until FeuilFournisseur.Range("A" & j).Value = ""
End Sub
