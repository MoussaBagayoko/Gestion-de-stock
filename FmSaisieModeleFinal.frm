VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FmSaisieModeleFinal 
   Caption         =   "FOURNISSEUR"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8835
   OleObjectBlob   =   "FmSaisieModeleFinal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FmSaisieModeleFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnnuler_Click()
Unload Me
End Sub

Private Sub CmdEnregistre_Click()
    If Me.TextBox1.Value = "" Then
        MsgBox "Veuillez saisir obligatoirement le nom et le prénoms", vbCritical
        Me.TextBox1.SetFocus
      ElseIf Me.TextBox2.Value = "" Then
        MsgBox "Veuillez saisir obligatoirement le numéro de téléphone ", vbCritical
        Me.TextBox2.SetFocus
      ElseIf DetectDoublons(FeuilFournisseur.Name, AjoutModif, LigneAModifier, Me.TextBox1.Value) = True Then
        MsgBox "Ce nom de fournisseur est déjà présent dans la base de données !", vbCritical
        
      Else
       Dim i As Long
       Dim j As Byte
       If AjoutModif = True Then
            i = Derligne(FeuilFournisseur.Name)
               Else
            i = LigneAModifier
            
       End If
            For j = 1 To 6
               FeuilFournisseur.Cells(i, j).Value = Me.Controls("TextBox" & j).Value
            Next j
            
       Unload Me
    End If

End Sub
