Public Class Methods
	Public ReadOnly Property QuestionTypeList() As List(Of String)
		Get
			Dim l_ As New List(Of String)
			l_.Add("Multiple Choice")
			l_.Add("Essay")
			l_.Add("Multiple Choice and Essay")
			Return l_
		End Get
	End Property
	Public Function TaskTypeList() As List(Of String)
		Dim l_ As New List(Of String)
		l_.Add("Salary")
		l_.Add("Sundries")
		l_.Add("Specific")
		Return l_
	End Function

	Public Function ReminderTypeList() As List(Of String)
		Dim l_ As New List(Of String)
		l_.Add("Voice")
		l_.Add("Message Prompt")
		l_.Add("Message Prompt And Voice")
		Return l_
	End Function

	Public Function ReminderList() As List(Of String)
		Dim l_ As New List(Of String)
		l_.Add("Not At All")
		l_.Add("15 minutes")
		l_.Add("3 days")
		l_.Add("1 week")
		l_.Add("1 month")
		Return l_
	End Function

	Public Function RecurrenceList() As List(Of String)
		Dim l_ As New List(Of String)
		l_.Add("Not At All")
		l_.Add("Daily")
		l_.Add("Weekly")
		l_.Add("Every 2 weeks")
		l_.Add("Monthly")
		l_.Add("Every 3 months")
		l_.Add("Every 6 months")
		l_.Add("Yearly")
		Return l_
	End Function

	Public Function StatusList(Optional IsUpdate As Boolean = False) As List(Of String)
		Dim l_ As New List(Of String)
		If IsUpdate = False Then l_.Add("Started")
		l_.Add("Done")
		l_.Add("Canceled")
		Return l_
	End Function

	Public Function GenderList() As List(Of String)
		Dim l_ As New List(Of String)
		l_.Add("Male")
		l_.Add("Female")
		Return l_
	End Function

End Class
