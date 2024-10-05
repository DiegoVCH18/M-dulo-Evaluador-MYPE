VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ESTADOGYP 
   Caption         =   "***MODULO BÁSICO DE EVALUACIÓN DE CRÉDITOS***"
   ClientHeight    =   13440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10455
   OleObjectBlob   =   "ESTADOGYP.frx":0000
   StartUpPosition =   3  'Predeterminado de Widnows
End
Attribute VB_Name = "ESTADOGYP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()

Me.Hide
End Sub

Private Sub CommandButton10_Click()
Me.Hide
EEGYP3.Show
End Sub

Private Sub CommandButton11_Click()
Me.Hide
EEGYP3.Show
End Sub

Private Sub CommandButton12_Click()
Me.Hide
EEGYP2.Show
End Sub

Private Sub CommandButton13_Click()
Me.Hide
EEGYP3.Show
End Sub

Private Sub CommandButton14_Click()
Me.Hide
EEGYP3.Show
End Sub

Private Sub CommandButton15_Click()
Me.Hide
EEGYP3.Show
End Sub

Private Sub CommandButton16_Click()
Me.Hide
EEGYP1.Show
End Sub

Private Sub CommandButton17_Click()
Me.Hide
EEGYP3.Show
End Sub

Private Sub CommandButton18_Click()
Me.Hide
EEGYP3.Show
End Sub

Private Sub CommandButton2_Click()
Application.ScreenUpdating = False
If TextBox25.Text = "" Then
MsgBox "completar el Análisis de Esatado de Ganancias y Pérdidas", , "MBEC v 1.2.0"
Else
Sheets("EEFF CONSOLIDADOS").Select
Cells(48, 16) = TextBox25.Value
Application.ScreenUpdating = False
Me.Hide
EEFF.Show
End If
End Sub

Private Sub TextBox14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Hide
EEGYP3.Show
End Sub

Private Sub TextBox27_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Hide
EEGYP3.Show
End Sub

Private Sub TextBox29_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Hide
EEGYP1.Show
End Sub


Private Sub TextBox12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Hide
EEGYP3.Show
End Sub

Private Sub TextBox18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Hide
EEGYP2.Show
End Sub

Private Sub TextBox20_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Hide
EEGYP3.Show
End Sub

Private Sub TextBox22_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Hide
EEGYP3.Show
End Sub

Private Sub TextBox25_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
        KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
    End If

End Sub

Private Sub UserForm_ACTIVATE()
Application.ScreenUpdating = False
Sheets("DATOS G Y P").Select
TextBox4.Text = Format(Cells(42, 20), "#,###,###,##0.00")




Sheets("EEFF CONSOLIDADOS").Select
TextBox29.Text = Format(Cells(54, 6), "#,###,###,##0.00")
TextBox27.Text = Format(Cells(56, 6), "#,###,###,##0.00")
TextBox12.Text = Format(Cells(60, 6), "#,###,###,##0.00")
TextBox14.Text = Format(Cells(62, 6), "#,###,###,##0.00")
TextBox18.Text = Format(Cells(66, 6), "#,###,###,##0.00")
TextBox20.Text = Format(Cells(68, 6), "#,###,###,##0.00")
TextBox22.Text = Format(Cells(70, 6), "#,###,###,##0.00")
TextBox31.Text = Format(Cells(74, 6), "#,###,###,##0.00")
TextBox35.Text = Format(Cells(78, 6), "#,###,###,##0.00")
TextBox25.Text = Cells(48, 16)
TextBox10.Text = Format(Cells(58, 6), "#,###,###,##0.00")
If TextBox10.Text < 0 Then

TextBox10.BackColor = &HFFC0FF

Else

TextBox10.BackColor = &HFFC0C0

End If

TextBox16.Text = Format(Cells(64, 6), "#,###,###,##0.00")
If TextBox16.Text < 0 Then

TextBox16.BackColor = &HFFC0FF

Else

TextBox16.BackColor = &HFFC0C0

End If


TextBox24.Text = Format(Cells(72, 6), "#,###,###,##0.00")
If TextBox24.Text < 0 Then

TextBox24.BackColor = &HFFC0FF

Else

TextBox24.BackColor = &HFFC0C0

End If

TextBox33.Text = Format(Cells(76, 6), "#,###,###,##0.00")
If TextBox33.Text < 0 Then

TextBox33.BackColor = &HFFC0FF

Else

TextBox33.BackColor = &HFFC0C0

End If

TextBox37.Text = Format(Cells(80, 6), "#,###,###,##0.00")
If TextBox37.Text < 0 Then

TextBox37.BackColor = &HFFC0FF

Else

TextBox37.BackColor = &HFFC0C0

End If


TextBox28.Text = Format(Cells(54, 7), "0.00%")
TextBox26.Text = Format(Cells(56, 7), "0.00%")
TextBox9.Text = Format(Cells(58, 7), "0.00%")
TextBox11.Text = Format(Cells(60, 7), "0.00%")
TextBox13.Text = Format(Cells(62, 7), "0.00%")
TextBox15.Text = Format(Cells(64, 7), "0.00%")
TextBox17.Text = Format(Cells(66, 7), "0.00%")
TextBox19.Text = Format(Cells(68, 7), "0.00%")
TextBox21.Text = Format(Cells(70, 7), "0.00%")
TextBox23.Text = Format(Cells(72, 7), "0.00%")
TextBox30.Text = Format(Cells(74, 7), "0.00%")
TextBox32.Text = Format(Cells(76, 7), "0.00%")
TextBox34.Text = Format(Cells(78, 7), "0.00%")
TextBox36.Text = Format(Cells(80, 7), "0.00%")
End Sub

Private Sub UserForm_Terminate()
Application.ScreenUpdating = False
Me.Hide
EEFF.Show

End Sub
