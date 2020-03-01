Imports Feedback.Feedback
Imports NModule.D
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO
Public Class DataConnectionDesktop
	Private f_ As New Functions

	''' <summary>
	''' Populates Combobox with grid's columns' HeaderText.
	''' </summary>
	''' <param name="LocateBy_"></param>
	''' <param name="LoadThis_"></param>
	''' <param name="grid_"></param>
	Public Sub SetSearch(LocateBy_ As ComboBox, LoadThis_ As ComboBox, Optional grid_ As DataGridView = Nothing)
		Clear(LoadThis_)
		LocateBy_.Sorted = True
		'		If LocateBy_.Items.Count > 0 Then Exit Sub
		Clear(LocateBy_)
		If grid_ IsNot Nothing Then
			With grid_
				For i As Integer = 0 To .Columns.Count - 1
					If .Columns.Item(i).Visible = True Then LocateBy_.Items.Add(.Columns.Item(i).HeaderText)
				Next
			End With
		End If
		LocateBy_.Text = ""
	End Sub

	Public Sub ClearText(c_ As Control)

		Dim c As ComboBox
		Dim t As TextBox
		Dim p As PictureBox
		Dim h As CheckBox
		If TypeOf c_ Is CheckBox Then
			h = c_
			h.Checked = False
		End If
		If TypeOf c_ Is ComboBox Then
			c = c_
			c.Text = ""
		End If
		If TypeOf c_ Is TextBox Then
			t = c_
			t.Text = ""
		End If
		If TypeOf c_ Is PictureBox Then
			p = c_
			Try
				p.Image = Nothing
			Catch ex As Exception
			End Try
			Try
				p.BackgroundImage = Nothing
			Catch ex As Exception
			End Try
		End If
	End Sub

	Public Sub ClearFields(dialog As Control, Optional clearData As Boolean = False)
		If TypeOf (dialog) IsNot Form And TypeOf (dialog) IsNot Panel Then Exit Sub

		For Each c As Control In dialog.Controls
			If clearData = True Then
				Clear(c)
			Else
				ClearText(c)
			End If
		Next
	End Sub

	''' <summary>
	''' Commits record to SQL Server database by default, or to MS Access database if DB_Is_SQL_ is set to false.
	''' </summary>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="parameters_keys_values_">Values to put in table.</param>
	''' <returns>True if successful, False if not.</returns>
	''' <example>
	''' Dim Insert_String As String = "UPDATE [Table_Name] SET Key1=@Key1Value, Key2=@Key2Value WHERE (PK=@PK)"
	''' Dim parameters_() = {}
	''' parameters_ = {"Key1Value", Key1Value, "Key2Value", Key2Value, "PK", PK}
	''' d.CommitRecord(Insert_String, con_string_, parameters_)
	''' </example>
	Public Function CommitRecord(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional DB_Is_SQL_ As Boolean = True) As Boolean
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = parameters_keys_values_

		If DB_Is_SQL_ = True Then
			CommitSQLRecord(query, connection_string, select_parameter_keys_values)
			Return True
			Exit Function
		End If

		Try
			Dim insert_query As String = query
			Using insert_conn As New OleDbConnection(connection_string)
				Using insert_comm As New OleDbCommand()
					With insert_comm
						.Connection = insert_conn
						'						.CommandTimeout = 0
						'						.CommandType = CommandType.Text
						.CommandText = insert_query
						If select_parameter_keys_values IsNot Nothing Then
							For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
								.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
							Next
						End If
					End With
					Try
						insert_conn.Open()
						insert_comm.ExecuteNonQuery()
					Catch ex As Exception
					End Try
				End Using
			End Using
			Return True
		Catch ex As Exception
		End Try

	End Function

	''' <summary>
	''' Commits record to SQL Server database by default, or to MS Access database if DB_Is_SQL_ is set to false. Same as CommitRecord.
	''' </summary>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="parameters_keys_values_">Values to put in table.</param>
	''' <returns>True if successful, False if not.</returns>
	Public Shared Function CommitSequel(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional DB_Is_SQL_ As Boolean = True) As Boolean
		Dim d_c As New DataConnectionDesktop
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = parameters_keys_values_

		If DB_Is_SQL_ = True Then
			d_c.CommitSQLRecord(query, connection_string, select_parameter_keys_values)
			Return True
			Exit Function
		End If

		Try
			Dim insert_query As String = query
			Using insert_conn As New OleDbConnection(connection_string)
				Using insert_comm As New OleDbCommand()
					With insert_comm
						.Connection = insert_conn
						'						.CommandTimeout = 0
						'						.CommandType = CommandType.Text
						.CommandText = insert_query
						If select_parameter_keys_values IsNot Nothing Then
							For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
								.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
							Next
						End If
					End With
					Try
						insert_conn.Open()
						insert_comm.ExecuteNonQuery()
					Catch ex As Exception
					End Try
				End Using
			End Using
			Return True
		Catch ex As Exception
		End Try


	End Function

	Public Function CommitSQLRecord(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As Boolean
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			Dim insert_query As String = query
			Using insert_conn As New SqlConnection(connection_string)
				Using insert_comm As New SqlCommand()
					With insert_comm
						.Connection = insert_conn
						.CommandTimeout = 0
						.CommandType = CommandType.Text
						.CommandText = insert_query
						If select_parameter_keys_values IsNot Nothing Then
							For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
								.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
							Next
						End If
					End With
					Try
						insert_conn.Open()
						insert_comm.ExecuteNonQuery()
					Catch ex As Exception
					End Try
				End Using
			End Using
			Return True
		Catch ex As Exception
		End Try

	End Function
	''' <summary>
	''' Checks if PictureBox has Image or BackgroundImage
	''' </summary>
	''' <param name="control_"></param>
	''' <param name="UseImage"></param>
	''' <returns></returns>
	Public Function ControlImageIsNull(control_ As PictureBox, Optional UseImage As Boolean = False) As Boolean
		If UseImage = True Then
			Return control_.Image Is Nothing
		Else
			Return control_.BackgroundImage Is Nothing
		End If
	End Function
	''' <summary>
	''' Checks if control has text.
	''' </summary>
	''' <param name="control_"></param>
	''' <param name="UseTrim"></param>
	''' <returns></returns>
	Public Function ControlTextIsNull(control_ As Control, Optional UseTrim As Boolean = True) As Boolean
		If UseTrim = True Then
			Return control_.Text.Trim.Length < 1
		Else
			Return control_.Text.Length < 1
		End If
	End Function
	Public Function ControlIsEmpty(control_ As Control, Optional VoiceFeedbackString As String = "", Optional UseTrim As Boolean = True, Optional UseVoiceFeedback As Boolean = True, Optional GetImageButton As Button = Nothing, Optional UseImage As Boolean = False) As Boolean
		If TypeOf control_ Is PictureBox Then
			If ControlImageIsNull(control_, UseImage) Then
				If GetImageButton IsNot Nothing Then GetImageButton.Focus()
				If UseVoiceFeedback = True And VoiceFeedbackString.Length > 0 Then returnFeedback(VoiceFeedbackString)
				Return True
				Exit Function
			Else
				Return False
				Exit Function
			End If
		End If

		If ControlTextIsNull(control_, UseTrim) Then
			control_.Focus()
			If UseVoiceFeedback = True And VoiceFeedbackString.Length > 0 Then ReturnFeedback(VoiceFeedbackString)
			Return True
		End If
	End Function

	''' <summary>
	''' Gets the stream of an image or image/backgroundImage of PictureBox or path to image.
	''' </summary>
	''' <param name="picture_">PictureBox or Image or Path to image</param>
	''' <param name="file_extension">File extension to save it with.</param>
	''' <param name="UseImage">Checks for PictureBox.Image instead of PictureBox.BackgroundImage.</param>
	''' <returns></returns>
	Public Shared Function PictureFromStream(picture_ As Object, Optional file_extension As String = ".jpg", Optional UseImage As Boolean = False) As Byte() ' IO.MemoryStream
		Dim photo_ As Image ' = Picture.BackgroundImage
		Dim stream_ As New IO.MemoryStream

		If TypeOf picture_ Is PictureBox Then
			Select Case UseImage
				Case True
					photo_ = picture_.Image
				Case False
					photo_ = picture_.BackgroundImage
			End Select
		ElseIf TypeOf picture_ Is Image Then
			photo_ = picture_
		ElseIf TypeOf picture_ Is String Then
			photo_ = Image.FromFile(picture_)
		End If


		If photo_ IsNot Nothing Then
			Select Case LCase(file_extension)
				Case ".jpg"
					photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
				Case ".jpeg"
					photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
				Case ".png"
					photo_.Save(stream_, Imaging.ImageFormat.Png)
				Case ".gif"
					photo_.Save(stream_, Imaging.ImageFormat.Gif)
				Case ".bmp"
					photo_.Save(stream_, Imaging.ImageFormat.Bmp)
				Case ".tif"
					photo_.Save(stream_, Imaging.ImageFormat.Tiff)
				Case ".ico"
					photo_.Save(stream_, Imaging.ImageFormat.Icon)
				Case Else
					photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
			End Select
		End If
		Return stream_.GetBuffer
	End Function

#Region "Bindings"
	''' <summary>
	''' Displays data on DataGridView.
	''' </summary>
	''' <param name="g_">DataGridView to bind to</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <example>Display(DataGridView, SQL_Query, Connection_String, Select_Parameters)</example>
	''' <returns>g_</returns>

	Public Shared Function Display(g_ As DataGridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As DataGridView
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			g_.DataSource = Nothing
		Catch ex As Exception
		End Try

		Try

			Dim connection As New SqlConnection(connection_string)
			Dim sql As String = query

			Dim Command = New SqlCommand(sql, connection)
			'		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
			'		Catch
			'		End Try

			Dim da As New SqlDataAdapter(Command)
			Dim dt As New DataTable
			da.Fill(dt)

			g_.DataSource = dt
			'			g_.DataBind()
		Catch
		End Try

		Return g_
		'		d.GData(gPayment, Payment_, g_con)

		'		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
		'		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

	End Function

	''' <summary>
	''' Binds ComboBox to database column.
	''' </summary>
	''' <param name="d_">ComboBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Column</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <param name="First_Element_Is_Empty">Should first element of ComboBox appear empty?</param>
	''' <returns></returns>
	Public Shared Function DData(d_ As ComboBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional First_Element_Is_Empty As Boolean = True) As ComboBox

		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			d_.DataSource = Nothing
		Catch ex As Exception

		End Try


		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		d_.DataSource = dt
		d_.DisplayMember = data_text_field

		If First_Element_Is_Empty Then d_.SelectedIndex = -1
		Return d_
	End Function

	''' <summary>
	''' Binds ComboBox Text property to database field.
	''' </summary>
	''' <param name="d_">ComboBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>d_</returns>
	Public Shared Function DText(d_ As ComboBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As ComboBox
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			d_.DataSource = Nothing
		Catch ex As Exception

		End Try


		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Text", dt, data_text_field)
		d_.DataBindings.Add(b)

		Return d_
	End Function

	''' <summary>
	''' Binds CheckBox Checked property to database field.
	''' </summary>
	''' <param name="h_">CheckBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>h_</returns>
	Public Shared Function HData(h_ As CheckBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As CheckBox

		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_

		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Checked", dt, data_text_field)
		h_.DataBindings.Add(b)

		Return h_
	End Function

	''' <summary>
	''' Binds CheckBox Text property to database field.
	''' </summary>
	''' <param name="h_">CheckBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>h_</returns>
	Public Shared Function HText(h_ As CheckBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As CheckBox

		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_

		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Text", dt, data_text_field)
		h_.DataBindings.Add(b)

		Return h_

	End Function

	''' <summary>
	''' Binds ListBox to database column.
	''' </summary>
	''' <param name="l_">ListBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Column</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>l_</returns>
	Public Shared Function LData(l_ As ListBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As ListBox
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			l_.DataSource = Nothing
		Catch ex As Exception

		End Try


		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		l_.DataSource = dt
		l_.DisplayMember = data_text_field

		Return l_
	End Function

	''' <summary>
	''' Binds PictureBox Image property to database field.
	''' </summary>
	''' <param name="p_">PictureBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>p_</returns>
	Public Shared Function PImage(p_ As PictureBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As PictureBox
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			p_.Image = Nothing
		Catch ex As Exception
		End Try


		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Image", dt, data_text_field, True)
		p_.DataBindings.Add(b)

		Return p_
	End Function

	''' <summary>
	''' Binds PictureBox BackgroundImage property to database field.
	''' </summary>
	''' <param name="p_">PictureBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>p_</returns>
	Public Shared Function PBackgroundImage(p_ As PictureBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As PictureBox
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			p_.BackgroundImage = Nothing
		Catch ex As Exception
		End Try


		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("BackgroundImage", dt, data_text_field, True)
		p_.DataBindings.Add(b)

		Return p_
	End Function

	''' <summary>
	''' Binds Button Text property to database field.
	''' </summary>
	''' <param name="b_">Button</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>b_</returns>
	Public Shared Function BData(b_ As Button, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As Button
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_


		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Text", dt, data_text_field)
		b_.DataBindings.Add(b)

		Return b_
	End Function

	''' <summary>
	''' Binds DateTimePicker Value property to database field.
	''' </summary>
	''' <param name="date_">DateTimePicker</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>date_</returns>
	Public Shared Function DATEData(date_ As DateTimePicker, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As DateTimePicker
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_

		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Value", dt, data_text_field)
		date_.DataBindings.Add(b)

		Return date_
	End Function

	''' <summary>
	''' Binds Label Text property to database field.
	''' </summary>
	''' <param name="label_">Label</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>label_</returns>
	Public Shared Function LABELData(label_ As Label, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As Label

		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_

		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Text", dt, data_text_field)
		label_.DataBindings.Add(b)

		Return label_

	End Function

	''' <summary>
	''' Binds TextBox Text property to database field.
	''' </summary>
	''' <param name="t_">TextBox</param>
	''' <param name="query">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="select_parameter_keys_values_">Select Parameters</param>
	''' <returns>t_</returns>
	Public Shared Function TData(t_ As TextBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As TextBox

		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_

		Dim connection As New SqlConnection(connection_string)
		Dim sql As String = query

		Dim Command = New SqlCommand(sql, connection)

		Try
			If select_parameter_keys_values IsNot Nothing Then
				For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
					Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
				Next
			End If
		Catch
		End Try

		Dim da As New SqlDataAdapter(Command)
		Dim dt As New DataTable
		da.Fill(dt)

		Dim b As New Binding("Text", dt, data_text_field)
		t_.DataBindings.Add(b)

		Return t_

	End Function

	''' <summary>
	''' Binds Control property to database column/field.
	''' </summary>
	''' <param name="control_">Control</param>
	''' <param name="property_">Property to bind to (Text, Value, Image, BackgroundImage or empty to bind to column)</param>
	''' <param name="query_">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="select_parameter_keys_values_">Select Parameters (Nothing, if not needed)</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="First_Element_Is_Empty">Should first element of ComboBox appear empty?</param>
	''' <returns>control_</returns>
	Public Shared Function BindProperty(control_ As Control, property_ As String, query_ As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional data_text_field As String = "", Optional First_Element_Is_Empty As Boolean = True) As Control
		'c, text
		'c, checked
		If TypeOf control_ Is CheckBox Then
			If property_.ToLower = "text" Then
				Return HText(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
			ElseIf property_.ToLower = "checked" Then
				Return HData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
			End If
		End If
		'g
		If TypeOf control_ Is DataGridView Then
			Return Display(control_, query_, connection_string, select_parameter_keys_values_)
		End If
		'd, text
		'd, data
		If TypeOf control_ Is ComboBox Then
			If property_.ToLower = "text" Then
				Return DText(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
			Else
				Return DData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_, First_Element_Is_Empty)
			End If
		End If
		'l
		If TypeOf control_ Is ListBox Then
			Return LData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
		End If
		'p, image
		'p, backgroundImage
		If TypeOf control_ Is PictureBox Then
			If property_.ToLower = "image" Then
				Return PImage(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
			Else
				Return PBackgroundImage(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
			End If
		End If
		'b, text
		If TypeOf control_ Is Button Then
			Return BData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
		End If
		'date, value
		If TypeOf control_ Is DateTimePicker Then
			Return DATEData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
		End If
		'l, text
		If TypeOf control_ Is Label Then
			Return LABELData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
		End If
		't, text
		If TypeOf control_ Is TextBox Then
			Return TData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
		End If
	End Function

#End Region

End Class
