VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FmTableModele 
   Caption         =   "Libelle de la table"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12990
   OleObjectBlob   =   "FmTableModele.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FmTableModele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAjouter_Click()
AjoutModif = True

End Sub

Private Sub CmdModifier_Click()
AjoutModif = False
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
