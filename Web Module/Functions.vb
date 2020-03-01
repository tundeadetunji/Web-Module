Imports Feedback.Feedback
Imports System.Windows.Forms
Imports System.IO
Public Class Functions
	'Private f As New Feedback.Feedback
#Region "Web-Bootstrap3, extended to 4"
	Public Function Success_B3() As String
		Return "success"
	End Function
	Public Function Danger_B3() As String
		Return "danger"
	End Function
	Public Function Well_B3() As String
		Return "well"
	End Function

	Public Shared Sub StandardAlert(str_ As String, div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional text_center As Boolean = False, Optional just_text As Boolean = True)
		If just_text = True Then
			If text_center = True Then
				div_.InnerHtml = "<span class=""text-center"">" & str_ & "</span>"
			Else
				div_.InnerHtml = "<span class=""text-left"">" & str_ & "</span>"
			End If
			Exit Sub
		End If
		If text_center = True Then
			div_.InnerHtml = "<div class=""alert alert-light text-center small"" role=""alert"">" & str_ & "</div>"
		Else
			div_.InnerHtml = "<div class=""alert alert-light text-left small"" role=""alert"">" & str_ & "</div>"
		End If
	End Sub

	''' <summary>
	''' Gives feedback to user with alert. Same as, and replaces Feedback() and Alert().
	''' </summary>
	''' <param name="str_">String to place inside</param>
	''' <param name="div_">Which control to place</param>
	''' <param name="WithClose">Should user be able to close it?</param>
	''' <param name="alert_OR_danger_OR_success_OR_warning">Format of the control</param>
	Public Shared Sub Toast(str_ As String, div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional WithClose As Boolean = True, Optional alert_OR_danger_OR_success_OR_warning As String = "warning")
		'		Dim h As System.Web.UI.HtmlControls.HtmlGenericControl = div_
		If IsNothing(div_) Then Exit Sub
		With div_
			.Visible = False
			If str_ Is Nothing Then Exit Sub
			If str_.Length < 1 Then Exit Sub
			.InnerHtml = Alert(str_, WithClose, alert_OR_danger_OR_success_OR_warning)
			.Visible = True
		End With
	End Sub

	''' <summary>
	''' Gives feedback to user. Same as Feedback.
	''' </summary>
	''' <param name="str_"></param>
	''' <param name="WithClose"></param>
	''' <param name="alert_OR_danger_OR_success_OR_warning"></param>
	''' <returns></returns>
	Public Shared Function Alert(str_ As String, Optional WithClose As Boolean = True, Optional alert_OR_danger_OR_success_OR_warning As String = "warning") As String
		Dim w_f As New Functions
		Return w_f.Feedback(str_, WithClose, alert_OR_danger_OR_success_OR_warning)
	End Function

	''' <summary>
	''' Gives feedback to user. Constructed as an alert DIV.
	''' </summary>
	''' <param name="str_"></param>
	''' <param name="WithClose"></param>
	''' <param name="alert_OR_danger_OR_success_OR_warning"></param>
	''' <returns></returns>
	Public Function Feedback(str_ As String, Optional WithClose As Boolean = True, Optional alert_OR_danger_OR_success_OR_warning As String = "warning") As String
		Select Case WithClose
			Case True
				Return FeedbackWithClose(str_, alert_OR_danger_OR_success_OR_warning)
			Case False
				Return FeedbackWithoutClose(str_, alert_OR_danger_OR_success_OR_warning)
		End Select



		'<span runat = "server" id="x"></span>
		'x.InnerHtml = f.Feedback(f.InvalidCredentialFeedback, True, "alert")
		'x.Visible = True

	End Function

	Public Function FeedbackWithoutClose(str_ As String, alert_OR_danger_OR_success_OR_warning As String) As String
		Dim header_ As String = "", footer_ As String = "</div>"
		Select Case alert_OR_danger_OR_success_OR_warning.ToLower
			Case "alert"
				header_ = "<div class=""alert alert-primary text-center"" role=""alert"">"
			Case "danger"
				header_ = "<div class=""alert alert-danger text-center"" role=""alert"">"
			Case "success"
				header_ = "<div class=""alert alert-success text-center"" role=""alert"">"
			Case "warning"
				header_ = "<div class=""alert alert-warning text-center"" role=""alert"">"
		End Select
		Return header_ & str_ & footer_
	End Function

	Public Function FeedbackWithClose(str_ As String, alert_OR_danger_OR_success_OR_warning As String) As String
		Dim header_ As String = "", footer_ As String = "</div>"
		Dim close_ As String = "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close""><span aria-hidden=""true"">&times;</span></button>"

		Select Case alert_OR_danger_OR_success_OR_warning.ToLower
			Case "alert"
				header_ = "<div class=""alert alert-primary alert-dismissible fade show"" role=""alert"">"
			Case "danger"
				header_ = "<div class=""alert alert-danger alert-dismissible fade show"" role=""alert"">"
			Case "success"
				header_ = "<div class=""alert alert-success alert-dismissible fade show"" role=""alert"">"
			Case "warning"
				header_ = "<div class=""alert alert-warning alert-dismissible fade show"" role=""alert"">"
		End Select

		Return header_ & str_ & close_ & footer_

	End Function

	Public Function Feedback_B3(str_ As String, Optional WithClose As Boolean = False, Optional well_OR_danger_OR_success As String = "danger", Optional size_in_px As Integer = 0, Optional TEXT_ALIGN_left_OR_center_OR_right As String = "center") As String
		If well_OR_danger_OR_success.Trim.Length < 1 Then well_OR_danger_OR_success = "well"
		Dim style_ As String = ""
		If Val(size_in_px) > 0 Then
			style_ = "style=""width:" & size_in_px & "px;"
		Else
			style_ = "style="""
		End If
		If TEXT_ALIGN_left_OR_center_OR_right.ToLower = "left" Then
			style_ &= "text-align:left"""
		ElseIf TEXT_ALIGN_left_OR_center_OR_right.ToLower = "right" Then
			style_ &= "text-align:right"""
		Else
			style_ &= "text-align:center"""
		End If

		Return "<div " & style_ & Xwell_OR_danger_OR_success_B3(well_OR_danger_OR_success) & ">" & XWithClose_B3(WithClose) & str_ & "</div>"
	End Function

	Public Function XWithClose_B3(WithClose As Boolean) As String
		Select Case WithClose
			Case True
				Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>"
			Case False
				Return ""
		End Select
	End Function

	Public Function Xwell_OR_danger_OR_success_B3(str_ As String) As String
		Select Case str_.ToLower
			Case "well"
				Return " class=""well well-sm fade in"" "
			Case "success"
				Return " class=""alert alert-success fade in"" "
			Case Else
				Return " class=""alert alert-danger fade in"" "
		End Select
	End Function
	Public Shared ReadOnly Property CommitRecordFeedback() As String = "Changes have been saved."

	Public Shared ReadOnly Property NewRecordFeedback() As String = "Record has been created."

	Public Shared ReadOnly Property InvalidDateFeedback() As String = "Please enter a valid date."

	Public Shared ReadOnly Property StandardFailureFeedback() As String = "Oops! An error occured. Please restart your browser or try again in a few minutes."

	Public Shared ReadOnly Property StandardFailure_LoggedIn_Feedback() As String = "Oops! An error occured. Please log out and log in again, restart your browser or try again in a few minutes."

	Public Shared ReadOnly Property JITFailureFeedback() As String = "Error occured while processing your request. Please verify that it was successful. If it persists, please make an enquiry."

	Public Shared ReadOnly Property DefaultFailureFeedback() As String = "Your request cannot be processed at this time. Please restart your browser or try again in a few minutes."

	Public Shared ReadOnly Property EnquiryFeedback(Optional email_ As String = "") As String
		Get
			email_ = email_.Trim
			If email_.Length < 1 Then
				Return "Enquiry has been received. Please check your email for response within 48 working hours."
			ElseIf email_.Length > 0 Then
				Return "Enquiry has been received. Please check " & email_ & " for response within 48 working hours."
			End If

		End Get
	End Property

	Public Shared ReadOnly Property SectionExistsFeedback() As String = "Please choose another name. That section already exists."

	Public Shared ReadOnly Property InvalidFieldFeedback() As String = "One or more values is invalid. Please check your input."

	Public Shared ReadOnly Property InvalidPasswordFeedback() As String = "Please enter your current password."

	Public Shared ReadOnly Property TermsFeedback() As String = "You must accept the terms and conditions."

	Public Shared ReadOnly Property RegistrationFailureFeedback() As String = "Oops! An error occured. Try <i><a href=""https://www.inovationware.com/Account/Login"">logging in here</a></i>, but if it doesn't, please restart your browser or try again in a few minutes."

	Public Shared ReadOnly Property EmailFeedback() As String = "Please enter an email."

	''' <summary>
	''' Feedback string.
	''' </summary>
	''' <returns>Please select a plan.</returns>
	Public Shared ReadOnly Property PlanFeedback() As String = "Please select a plan."

	''' <summary>
	''' Feedback string.
	''' </summary>
	''' <returns>Thank you. Your plan will be updated upon next payment.</returns>
	Public Shared ReadOnly Property PlanUpdateFeedback() As String = "Thank you. Your plan will be updated upon next payment."

	Public Shared ReadOnly Property CloseAccountFeedback() As String = "You will no longer have access to your account and records, effective immediately. You will be logged out. All your data will be deleted. This process cannot be undone."

	Public Shared ReadOnly Property IDTakenFeedback() As String = "Please choose another ID. That ID has been taken."

	Public Shared ReadOnly Property UsernameTakenFeedback() As String = "Please choose another username. That username has been taken."

	Public Shared ReadOnly Property RemovePostFeedback() As String = "Really remove post?"

	Public Shared ReadOnly Property InvalidAccountFeedback() As String = "Couldn't find matching record. The account either does not exist or is not activated."

	Public Shared ReadOnly Property InvalidCredentialFeedback() As String = "Couldn't find matching record. Username/Password combination is invalid."

	Public Shared ReadOnly Property InvalidPINFeedback() As String = "Couldn't find matching record. Admission Number/PIN combination is invalid."

	Public Shared ReadOnly Property EmailLinkSentFeedback() As String = "Email has been sent. Please check your mailbox."

	Public Shared ReadOnly Property PasswordResetExpiryFeedback() As String = "This page is no longer valid. Please make another request."

	Public Shared ReadOnly Property PasswordResetFeedback() As String = "Your password has been changed."


	Public Function FICFeedback() As String
		'failure invalid password
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>Wrong password."
	End Function

	Public Function LFFeedback() As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>We couldn't find matching ID and password."
	End Function

	Public Function FFeedback() As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>Your request cannot be processed at this time. Please log out and log in again, restart your browser or try again in a few minutes. If it persists, please make an enquiry."
	End Function
	Public Function FFFeedback(str As String) As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>" & str
	End Function

	Public Function SFeedback(str As String) As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>" & str
	End Function

	Public Function CFFeedback(str As String) As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>" & str
	End Function

	Public Function LoggedInFailureFeedback_B3() As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>Your request cannot be processed at this time. Please log out and log in again, restart your browser or try again in a few minutes."
	End Function

	Public Function RegistrationFailureFeedback_B3() As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>Your request cannot be processed at this time. Please restart your browser or try again in a few minutes."
	End Function

	Public Function RegistrationSuccessFeedback() As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>Account has been created. You can log in with your username and password."
	End Function

	Public Function UpdateSuccessFeedback() As String
		Return "<a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a>Changes have been saved."
	End Function


#End Region

#Region "Security"
	Private Function IsUpperCase(c As String) As Boolean
		Dim s As String = c.ToLower
		Dim IsNotAnAlphabet As Boolean = False
		If s <> "a" And s <> "b" And s <> "c" And s <> "d" And s <> "e" And s <> "f" And s <> "g" And s <> "h" And s <> "i" And s <> "j" And s <> "k" And s <> "l" And s <> "m" And s <> "n" And s <> "o" And s <> "p" And s <> "q" And s <> "r" And s <> "s" And s <> "t" And s <> "u" And s <> "v" And s <> "w" And s <> "x" And s <> "y" And s <> "z" Then
			IsNotAnAlphabet = True
		End If
		Dim isInUpperCase As Boolean = False
		If c = "A" Or c = "B" Or c = "C" Or c = "D" Or c = "E" Or c = "F" Or c = "G" Or c = "H" Or c = "I" Or c = "J" Or c = "K" Or c = "L" Or c = "M" Or c = "N" Or c = "O" Or c = "P" Or c = "Q" Or c = "R" Or c = "S" Or c = "T" Or c = "U" Or c = "V" Or c = "W" Or c = "X" Or c = "Y" Or c = "Z" Then
			isInUpperCase = True
		End If
		If IsNotAnAlphabet = True Or isInUpperCase = True Then
			Return True
		Else
			Return False
		End If

	End Function
	Private Function IsLowerCase(c As String) As Boolean
		Dim s As String = c.ToLower
		Dim IsNotAnAlphabet As Boolean = False
		If s <> "a" And s <> "b" And s <> "c" And s <> "d" And s <> "e" And s <> "f" And s <> "g" And s <> "h" And s <> "i" And s <> "j" And s <> "k" And s <> "l" And s <> "m" And s <> "n" And s <> "o" And s <> "p" And s <> "q" And s <> "r" And s <> "s" And s <> "t" And s <> "u" And s <> "v" And s <> "w" And s <> "x" And s <> "y" And s <> "z" Then
			IsNotAnAlphabet = True
		End If
		Dim isInLowerCase As Boolean = False
		If c = "a" Or c = "b" Or c = "c" Or c = "d" Or c = "e" Or c = "f" Or c = "g" Or c = "h" Or c = "i" Or c = "j" Or c = "k" Or c = "l" Or c = "m" Or c = "n" Or c = "o" Or c = "p" Or c = "q" Or c = "r" Or c = "s" Or c = "t" Or c = "u" Or c = "v" Or c = "w" Or c = "x" Or c = "y" Or c = "z" Then
			isInLowerCase = True
		End If
		If IsNotAnAlphabet = True Or isInLowerCase = True Then
			Return True
		Else
			Return False
		End If

	End Function

	Private Function HasCharacter(string_ As String, character_ As String) As Boolean

		If string_.Contains(character_) Then
			Return True
		Else
			Return False
		End If

	End Function

	Private Function ContainsSpecialCharacter(s As String) As Boolean
		If s.Contains("~") Or s.Contains("!") Or s.Contains("@") Or s.Contains("#") Or s.Contains("$") Or s.Contains("%") Or s.Contains("&") Or s.Contains("*") Or s.Contains("|") Or s.Contains("_") Then
			Return True
		Else
			Return False
		End If
	End Function
	Private Function ContainsNumber(string_ As String) As Boolean
		Dim val As Boolean = False
		If string_.Contains(0) Or string_.Contains(1) Or string_.Contains(2) Or string_.Contains(3) Or string_.Contains(4) Or string_.Contains(5) Or string_.Contains(6) Or string_.Contains(7) Or string_.Contains(8) Or string_.Contains(9) Then val = True
		Return val
	End Function
	''' <summary>
	''' Checks if password hasUpperCase, hasLowerCase, hasNumber, hasSpecialCharacter, hasMinimumLength. You can adjust respective functions and values to customize. Uncomment the lines to include hasUnderscore.
	''' </summary>
	''' <param name="str_"></param>
	''' <returns></returns>
	''' <example>
	''' Dim password_is_invalid As String = f.CheckPassword(Password.Text)
	''' If password_is_invalid.Length > 0 Then
	''' x.InnerHtml = password_is_invalid
	''' End If
	''' </example>
	Public Function CheckPassword(str_ As String) As String
		Dim p As String = str_
		Dim hasUpperCase As Boolean = False
		Dim hasLowerCase As Boolean = False
		'		Dim hasUnderscore As Boolean = False
		Dim hasNumber As Boolean = False
		Dim hasSpecialCharacter As Boolean = False
		Dim hasMinimumLength As Boolean = False

		For i As Integer = 1 To p.Length
			If IsUpperCase(Mid(p, i, 1)) = True Then hasUpperCase = True : Exit For
		Next

		For i As Integer = 1 To p.Length
			If IsLowerCase(Mid(p, i, 1)) = True Then hasLowerCase = True : Exit For
		Next

		'		If HasCharacter(p, "_") = True Then hasUnderscore = True

		If ContainsNumber(p) = True Then hasNumber = True

		If ContainsSpecialCharacter(p) = True Then hasSpecialCharacter = True

		If p.Length >= 8 Then hasMinimumLength = True

		Dim s As String = ""
		If hasUpperCase = False Then
			s = "Password must contain at least 1 upper case letter"
		End If

		If hasLowerCase = False Then
			If s.Length < 1 Then
				s &= "Password must contain at least 1 lower case letter"
			Else
				s &= "<br />Password must contain at least 1 lower case letter"
			End If
		End If

		'		If hasUnderscore = False Then
		'			If s.Length < 1 Then
		'				s &= "Password must contain at least 1 underscore"
		'			Else
		'				s &= "<br />Password must contain at least 1 underscore"
		'			End If
		'		End If

		If hasNumber = False Then
			If s.Length < 1 Then
				s &= "Password must contain at least 1 number"
			Else
				s &= "<br />Password must contain at least 1 number"
			End If
		End If

		If hasSpecialCharacter = False Then
			If s.Length < 1 Then
				s &= "Password must contain at least 1 of the following: ~!@#$%&*|_"
			Else
				s &= "<br />Password must contain at least 1 of the following: ~!@#$%&*|_"
			End If
		End If

		If hasMinimumLength = False Then
			If s.Length < 1 Then
				s &= "Password must be at least 8 characters"
			Else
				s &= "<br />Password must be at least 8 characters"
			End If
		End If
		Return s

	End Function


#End Region

#Region "Print"
	''' <summary>
	''' Gets the fields of a DataGridView as string. Customize as needed.
	''' </summary>
	''' <param name="grid_"></param>
	''' <param name="stat_"></param>
	''' <param name="app_"></param>
	''' <returns></returns>
	Public Function GridInfo(grid_ As DataGridView, stat_ As String, app_ As String) As String
		Dim str$ = "Printout from " & app_ & vbCrLf & "Date: " & Now.ToShortDateString & ",  Time: " & Now.ToLongTimeString & vbCrLf & vbCrLf & vbCrLf & stat_ & vbCrLf & vbCrLf

		With grid_
			For pg_r% = 0 To .Rows.Count - 1
				str &= "#" & pg_r + 1 & vbCrLf
				For pg_c% = 0 To .Columns.Count - 1
					If .Columns.Item(pg_c).Visible = True Then str &= .Columns.Item(pg_c).HeaderText & ":" & vbCrLf & .Rows(pg_r).Cells(pg_c).Value.ToString & vbCrLf
				Next
				str &= vbCrLf
			Next
		End With

		Return str
	End Function
	''' <summary>
	''' Saves the information on DataGridView to a file.
	''' </summary>
	''' <param name="grid_"></param>
	''' <param name="stat_"></param>
	''' <param name="app_"></param>

	Public Sub GridToFile(grid_ As DataGridView, stat_ As String, app_ As String)
		If grid_.Rows.Count < 1 Then Exit Sub
		Dim str_$ = GridInfo(grid_, stat_, app_)
		SaveToFile(str_, app_)
	End Sub

	Public Sub SaveToFile(str As String, app As String)
		Dim file_location_ As String = My.Application.Info.DirectoryPath & "\" & app & "\Saved Records"
		Try
			My.Computer.FileSystem.CreateDirectory(file_location_)
		Catch ex As Exception

		End Try
		Dim filename As String = file_location_ & "\" & Now.Year & "_" & Now.Month & "_" & Now.Day & "_" & Now.Hour & "_" & Now.Minute & "_" & Now.Second & "_" & Now.Millisecond & ".txt"
		Try
			My.Computer.FileSystem.WriteAllText(filename, str, False)
			If MsgBox("Information successfully saved to " & filename & vbCrLf & vbCrLf & "Open it?", MsgBoxStyle.YesNo + MsgBoxStyle.Information) = MsgBoxResult.Yes Then
				Try
					Process.Start(filename)
				Catch ex As Exception
					MsgBox("Could not open the file. One or more errors occured.", MsgBoxStyle.Information)
				End Try
			End If
		Catch ex_ As Exception
			returnFeedback("There was a problem while trying to process your request. Please veriy that the operation was successful.")
		End Try
	End Sub

#End Region
End Class
