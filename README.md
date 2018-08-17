# Code-f-r-Neuer-Benutzer

private sub cmd_close_Click()
  Me.undo
 DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmd_Save_Click()
If Not IsNull(Me.txtPass) And (Me.txtPass = Me.txtPass_Bestät) Then
    DoCmd.Close acForm, Me.Name, acSaveNo
     CurrentDb.Execute "INSERT INTO tblBenutzer(bzrName, bzrVorname, bzrLogin, bzrPass,bzrPass1)" _
     & "SELECT" & Nz(Me.Controls("txtName").Value, "") & "','" & Nz(Me.Controls("txtName").Value, "") & "','" & _
     Nz(Me.Controls("txtLogin").Value, "") & "', '" & Nz(Me.Controls("txtPass").Value, "") & "','" & _
     Nz(Me.Controls("txtPass_Bestät").Value, "") & "'"
ElseIf Me.txtPass <> Me.txtPass_Bestät Then    'IsNull (Me.txtPass) Or IsNull(Me.txtPass_Bestät) Or (Me.txtPass <> Me.txtPass_Bestät)
    MsgBox "Das Passwort bitte richtig eingeben!"
    'Nur Adden wenn, beide Passwörter übereinstimmen
'With Me
'  If Not IsNull(.txtPass) And (.txtPass = .txtPass_Bestät) Then
'    DoCmd.Close acForm, Me.Name, acSaveNo
'   .txtName = vbNullString
'   .txtVorname = vbNullString
'   .txtLogin = vbNullString
'   .txtPass = vbNullString
'   .txtPass_Bestät = vbNullString
' Else: IsNull (.txtPass) Or IsNull(.txtPass_Bestät) Or (.txtPass <> .txtPass_Bestät)
'    MsgBox "Das Passwort bitte richtig eingeben!"

' 'ElseIf IsNull(.txtPass) Or IsNull(.txtPass_Bestät) Or (.txtPass &amp &lt &amp &gt .txtPass_Bestät) Then
'                                                        '.txtPass &amp;lt;&amp;gt; .txtPass_Bestät
'    MsgBox "Das Passwort bitte richtig eingeben!"
 End If

End Sub

Private Sub Form_Load()
  DoCmd.OpenForm "Formular5"

    With Forms("Formular5")
    .txtPass.InputMask = "Password"
    .txtPass_Bestät.InputMask = "Password"
    
  End With
End Sub

private Sub txtPass_Click()
DoCmd.OpenForm "PasswortCheck"
End Sub

Private Sub txtPass_KeyUp(KeyCode As Integer, Shift As Integer)
DoCmd.OpenForm "PasswortCheck"
End Sub


###################################### M O D U L E ############################################################################

sub FormularMitTabelleVerknuepfen()
On Error Resume Next
 With Forms("Formular5")
  .DataEntry = 1
  .RecordSource = "tblBenutzer"
  .txtName.ControlSource = "bzrName"
  .txtVorname.ControlSource = "bzrVorname"
  .txtLogin.ControlSource = "bzrLogin"
  .txtPass.ControlSource = "bzrPass"
  .txtPass_Bestät.ControlSource = "bzrPass1"
 End With
End Sub

sub NeuenBenutzer()
Dim frm As Form
Dim ctlLabel_Text As Control
Const ctlBreite As Integer = 1450
Const ctlMargeWaagerecht As Integer = 1500
Const ctlMargeSenkrecht As Integer = 400
Const ctlHoehe As Integer = 350
Const waagerecht As Integer = 3000
Const senkrecht As Integer = 2000
DoCmd.OpenForm "LoginForm", acDesign
Set frm = Forms("LoginForm")
"", , waagerecht, senkrecht + ctlMargeSenkrecht * 5, ctlBreite * 1.5, ctlHoehe)
with ctlLabel_Text
  .Caption = "Formular5"
  .Name = "lblBenutzer"
  .ForeColor =vbBlue
  .FontUnderline = 1
  .Hyperlink.SubAddress = "Form Formular5"
End With
end Sub
