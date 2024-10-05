VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RATIOS 
   Caption         =   "***MODULO BÁSICO DE EVALUACIÓN DE CRÉDITOS***"
   ClientHeight    =   9150.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15345
   OleObjectBlob   =   "RATIOS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "RATIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
If TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox16.Text = "" Or TextBox17.Text = "" Then
MsgBox "completar los Análisis de Ratios", , "MBEC v 1.2.0"
Else
Sheets("EEFF CONSOLIDADOS").Select
Cells(73, 16) = TextBox17.Value
Cells(67, 16) = TextBox16.Value
Cells(61, 16) = TextBox15.Value
Cells(55, 16) = TextBox14.Value
Application.ScreenUpdating = False
Me.Hide
EEFF.Show
End If
End Sub

Private Sub Label15_Click()

End Sub

Private Sub TextBox14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub Textbox15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If

End Sub

Private Sub TextBox16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If

End Sub

Private Sub TextBox17_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If
End Sub

Private Sub UserForm_ACTIVATE()
Sheets("EEFF CONSOLIDADOS").Select
TextBox1.Text = Format(Cells(56, 11), "#,###,###,##0.00")
TextBox2.Text = Format(Cells(58, 11), "#,###,###,##0.00")
TextBox3.Text = Format(Cells(60, 11), "#,###,###,##0.00")
TextBox4.Text = Format(Cells(62, 11), "#,###,###,##0.00")

TextBox6.Text = Format(Cells(66, 11), "0.00%")
TextBox7.Text = Format(Cells(68, 11), "0.00%")
TextBox8.Text = Format(Cells(70, 11), "0.00%")

TextBox5.Text = Format(Cells(74, 11), "0.00%")
TextBox9.Text = Format(Cells(76, 11), "0.00%")
TextBox10.Text = Format(Cells(78, 11), "0.00%")

TextBox12.Text = Format(Cells(82, 11), "#,###,###,##0.00")
TextBox13.Text = Format(Cells(84, 11), "#,###,###,##0.00")
TextBox11.Text = Format(Cells(86, 11), "#,###,###,##0.00")

TextBox14.Text = Cells(55, 16)
TextBox15.Text = Cells(61, 16)
TextBox16.Text = Cells(67, 16)
TextBox17.Text = Cells(73, 16)


End Sub

Private Sub UserForm_Terminate()
Application.ScreenUpdating = False
Me.Hide
EEFF.Show
End Sub
