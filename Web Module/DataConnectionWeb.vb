Imports NModule.W
Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Web.UI.WebControls
Public Class DataConnectionWeb


	''' <summary>
	''' Same as NModule.W.GenderDrop.
	''' </summary>
	''' <param name="d_"></param>
	''' <param name="FirstIsEmpty"></param>
	Public Sub PopulateGenderDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Dim web_methods_ As New Methods

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.GenderList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
	End Sub
	''' <summary>
	''' Populates DropDownList with user-friendly version of True/False, depending on the pattern.
	''' </summary>
	''' <param name="d_"></param>
	''' <param name="pattern_"></param>
	''' <param name="FirstIsEmpty"></param>
	Public Sub DropTextBoolean(d_ As DropDownList, Optional pattern_ As String = "always/never", Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		With d_.Items
			Select Case pattern_.Trim.ToLower
				Case ""
					.Add("Always")
					.Add("Never")
				Case "yes/no"
					.Add("Yes")
					.Add("No")
				Case "always/never"
					.Add("Always")
					.Add("Never")
				Case "on/off"
					.Add("On")
					.Add("Off")
				Case "1/0"
					.Add("1")
					.Add("0")
				Case "true/false"
					.Add("True")
					.Add("False")
			End Select

		End With

	End Sub
	''' <summary>
	''' Gets the content of the first cell in the first row of GridView.
	''' </summary>
	''' <param name="g_"></param>
	''' <returns></returns>
	Public Shared Function GetData(g_ As GridView)
		If g_.Rows.Count > 0 Then Return g_.Rows(0).Cells(0).Text
	End Function

	''' <summary>
	''' Checks if control is empty. You might want to use NModule.W.IsEmpty instead.
	''' </summary>
	''' <param name="c_"></param>
	''' <param name="use_trim_"></param>
	''' <returns></returns>
	Public Function WebControlIsEmpty(c_ As WebControl, Optional use_trim_ As Boolean = False) As Boolean
		Dim t_ As TextBox
		Dim d_ As DropDownList
		Dim l_ As ListBox

		If TypeOf c_ Is TextBox Then
			t_ = c_
			If use_trim_ = True Then
				Return t_.Text.Trim.Length < 1
			Else
				Return t_.Text.Length < 1
			End If
		ElseIf TypeOf c_ Is DropDownList Then
			d_ = c_
			Return d_.Items.Count < 1
		ElseIf TypeOf c_ Is ListBox Then
			l_ = c_
			Return l_.Items.Count < 1
		End If
	End Function

	''' <summary>
	''' Populates DropDownList with chart types. You can adjust the list.
	''' </summary>
	''' <param name="d_"></param>
	''' <param name="supports_pie"></param>
	Public Sub ChartTypeDrop(d_ As DropDownList, Optional supports_pie As Boolean = False)
		ClearDropDown(d_)
		With d_
			With .Items
				.Add("Column")
				.Add("Line")
				If supports_pie = True Then .Add("Pie")
			End With
		End With
	End Sub

	''' <summary>
	''' Populates DropDownList with chart styles. You can adjust the list.
	''' </summary>
	''' <param name="d_"></param>
	Public Sub ChartStyleDrop(d_ As DropDownList)
		ClearDropDown(d_)
		With d_
			With .Items
				.Add("3-D")
				.Add("Flat")
			End With
		End With
	End Sub

	''' <summary>
	''' Binds CheckBoxList to database column.
	''' </summary>
	''' <param name="c_"></param>
	''' <param name="query"></param>
	''' <param name="connection_string"></param>
	''' <param name="data_text_field"></param>
	''' <param name="data_value_field"></param>
	''' <param name="select_parameter_keys_values_"></param>
	Public Sub CData(c_ As CheckBoxList, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing)
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			c_.DataSource = Nothing
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

			Dim da As New SqlDataAdapter(Command)
			Dim dt As New DataTable
			da.Fill(dt)

			c_.DataSource = dt
			c_.DataTextField = data_text_field
			If data_value_field.Length > 0 Then c_.DataValueField = data_value_field
			c_.DataBind()
		Catch
		End Try

	End Sub

	''' <summary>
	''' Clears ListBox of items and DataSource.
	''' </summary>
	''' <param name="l_"></param>
	Public Sub ClearList(l_ As ListBox)
		Try
			l_.DataSource = Nothing
		Catch ex As Exception
		End Try
		l_.Items.Clear()
	End Sub

	''' <summary>
	''' Clears DropDownList of items and DataSource.
	''' </summary>
	Public Sub ClearDropDown(d_ As DropDownList)
		Try
			d_.DataSource = Nothing
		Catch ex As Exception
		End Try
		d_.Items.Clear()
	End Sub

	''' <summary>
	''' Clears ListBox or DropDownList of items and DataSource.
	''' </summary>
	''' <param name="l_"></param>
	''' <param name="d_"></param>
	Public Sub Clear_(Optional l_ As ListBox = Nothing, Optional d_ As DropDownList = Nothing)
		Try
			If l_ IsNot Nothing Then
				ClearList(l_)
			End If
		Catch
		End Try
		Try
			If d_ IsNot Nothing Then
				ClearDropDown(d_)
			End If
		Catch
		End Try
	End Sub

	''' <summary>
	''' Binds SQL Server database table to GridView. You might want to use Display instead.
	''' </summary>
	''' <example>
	''' <code>
	''' Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
	''' d.GData(gPayment, School_, m_con, select_parameter_keys_values)
	''' </code>
	''' <code>
	''' d.GData(gPayment, Payment_, g_con)
	''' </code>
	''' </example>
	''' <param name="g_">GridView to bind to.</param>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="select_parameter_keys_values_">The SQL select parameters.</param>
	Public Sub GData(g_ As GridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing)
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

			Dim da As New SqlDataAdapter(Command)
			Dim dt As New DataTable
			da.Fill(dt)

			g_.DataSource = dt
			g_.DataBind()
		Catch
		End Try

	End Sub

	''' <summary>
	''' Binds SQL Server database table to GridView. Returns the GridView. Same as GData.
	''' </summary>
	''' <param name="g_">GridView to bind to.</param>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="select_parameter_keys_values_">The SQL select parameters.</param>
	''' <return>g_</return>
	Public Shared Function Display(g_ As GridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As GridView
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

			Dim da As New SqlDataAdapter(Command)
			Dim dt As New DataTable
			da.Fill(dt)

			g_.DataSource = dt
			g_.DataBind()
		Catch
		End Try

		Return g_
	End Function

	''' <summary>
	''' Binds DropDownList to SQL database column.
	''' </summary>
	''' <param name="d_"></param>
	''' <param name="query"></param>
	''' <param name="connection_string"></param>
	''' <param name="data_text_field"></param>
	''' <param name="data_value_field"></param>
	''' <param name="select_parameter_keys_values_"></param>
	''' <param name="DontIfFull">Ignores the function if d_ already has items</param>
	''' <returns>d_</returns>
	Public Shared Function DData(d_ As DropDownList, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing, Optional DontIfFull As Boolean = False) As DropDownList
		If DontIfFull = True Then
			If d_.Items.Count > 0 Then Return d_
		End If
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
		d_.DataTextField = data_text_field
		If data_value_field.Length > 0 Then d_.DataValueField = data_value_field
		d_.DataBind()
		Return d_
	End Function

	''' <summary>
	''' Binds DropDownList to List.
	''' </summary>
	''' <param name="d_"></param>
	''' <param name="dont_if_full">Ignores the function if d_ already has items.</param>
	''' <param name="object_">List.</param>
	''' <returns>d_</returns>
	Public Shared Function DData(d_ As DropDownList, object_ As Object, Optional dont_if_full As Boolean = False) As DropDownList
		If dont_if_full And d_.Items.Count > 0 Then Return d_
		Try
			Dim l As New List(Of String)
			l = CType(object_, List(Of String))
			With d_
				.DataSource = l
				.DataBind()
			End With
		Catch
		End Try
		Return d_
	End Function

	''' <summary>
	''' Binds ListBox to database column.
	''' </summary>
	''' <param name="l_"></param>
	''' <param name="query"></param>
	''' <param name="connection_string"></param>
	''' <param name="data_text_field"></param>
	''' <param name="data_value_field"></param>
	''' <param name="select_parameter_keys_values_"></param>

	Public Sub LData(l_ As ListBox, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing)
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
		l_.DataTextField = data_text_field
		If data_value_field.Length > 0 Then l_.DataValueField = data_value_field
		l_.DataBind()

	End Sub

	''' <summary>
	''' Commits record to SQL Server database. You might want to use CommitSequel instead.
	''' </summary>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="parameters_keys_values_">Values to put in table.</param>
	''' <returns>True if successful, False if not.</returns>
	Public Function CommitRecord(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As Boolean
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = parameters_keys_values_
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
			Return False
		End Try

	End Function

	''' <summary>
	''' Commits record to SQL Server database. Same as CommitRecord.
	''' </summary>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="parameters_keys_values_">Values to put in table.</param>
	''' <returns>True if successful, False if not.</returns>
	''' <example>
	''' Dim Entries_Insert As String = "INSERT INTO ENTRIES (EntryBy, ID, Category) VALUES (@EntryBy, @ID, @Category)"
	''' Dim entries_parameters_() = {"EntryBy", TitleBy.Text.Trim, "ID", EntryID.Text.Trim, "Category", Category.Text.Trim}
	''' d.CommitRecord(Entries_Insert, connection_string, entries_parameters_)
	''' </example>
	Public Shared Function CommitSequel(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As Boolean
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = parameters_keys_values_
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
			Return False
		End Try

	End Function

	''' <summary>
	''' Gets the stream of image from Image control.
	''' </summary>
	''' <param name="picture_">Image control.</param>
	''' <param name="file_extension">File extension to save it with.</param>
	''' <returns></returns>

	Public Function PictureFromStream(picture_ As System.Drawing.Image, Optional file_extension As String = ".jpg") As Byte()
		Dim photo_ As System.Drawing.Image = picture_
		Dim stream_ As New IO.MemoryStream

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

End Class
