Imports System.Data.SqlClient
Public Class DW

	''' <summary>
	''' Builds SQL Update Query String.
	''' </summary>
	''' <param name="t_">Table to update.</param>
	''' <param name="update_keys">Fields to replace.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to update or not, followed by the operator to apply on the key (empty means equal to).</param>
	''' <returns>String.</returns>
	''' <example>
	''' BuildUpdateString_CONDITIONAL("Table_Name", {"FieldToUpdate1"}, {"Field_To_Check_Condition_On_1", ">"})
	''' </example>
	Public Shared Function BuildUpdateString_CONDITIONAL(t_ As String, update_keys As Array, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "UPDATE " & t_ & " SET "

		For j As Integer = 0 To update_keys.Length - 1
			v &= update_keys(j) & "=@" & update_keys(j)
			If update_keys.Length > 1 And j <> update_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("

				For k As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(k) & " " & operator_(where_keys(k + 1)) & " @" & where_keys(k)
					If where_keys.Length > 1 And k <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
				v &= ")"
			End If
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Count Query String.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to count or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	''' <example>
	''' BuildCountString_CONDITIONAL("Table_Name", {"Field_To_Check_Condition_On_1", "=", "Field_To_Check_Condition_On_2", ">"})
	''' </example>
	Public Shared Function BuildCountString_CONDITIONAL(t_ As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT COUNT(*) "
		'For i As Integer = 0 To select_params.Length - 1
		'	v &= select_params(i)
		'	If select_params.Length > 1 And i <> select_params.Length - 1 Then
		'		v &= ", "
		'	End If
		'Next
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function

	Private Shared Function operator_(operator__ As String) As String
		If operator__ = "" Then
			Return "="
		Else
			Return operator__
		End If
	End Function

	''' <summary>
	''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Fields to select.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to select or not, followed by the operator to apply on the key.</param>
	''' <returns>String.</returns>
	''' <example>
	''' BuildSelectString_CONDITIONAL("Table_Name", {"Field_To_Select_1"}, {"Field_To_Check_Condition_On_1", "", "Field_To_Check_Condition_On_2", ">"})
	''' </example>
	Public Shared Function BuildSelectString_CONDITIONAL(t_ As String, Optional select_params As Array = Nothing, Optional where_key_operator As Array = Nothing, Optional OrderByField As String = "PRIMARY_KEY_OR_OTHER") As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField.Length > 0 Then v &= " ORDER BY " & OrderByField

		Return v
	End Function

	''' <summary>
	''' Builds SQL Select Query String with DISTINCT. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Fields to select.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
	''' <returns>String.</returns>
	''' <example>
	''' BuildSelectString("Table_Name", {"Field_To_Select_1"}, {"Field_To_Check_Condition_On_1"})
	''' </example>

	Public Shared Function BuildSelectString_DISTINCT(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT DISTINCT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Insert Query String.
	''' </summary>
	''' <param name="t_">Table to insert into.</param>
	''' <param name="insert_keys">Columns to insert into.</param>
	''' <returns>String.</returns>
	''' <example>
	''' BuildInsertString("Table_Name", {"Field_To_Insert_1", "Field_To_Insert_2"}) 
	''' </example>
	Public Shared Function BuildInsertString(t_ As String, insert_keys As Array) As String
		Dim v As String = "INSERT INTO " & t_ & " ("

		For i As Integer = 0 To insert_keys.Length - 1
			v &= insert_keys(i)
			If insert_keys.Length > 1 And i <> insert_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		v &= ") VALUES ("

		For j As Integer = 0 To insert_keys.Length - 1
			v &= "@" & insert_keys(j)
			If insert_keys.Length > 1 And j <> insert_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		v &= ")"
		Return v
	End Function

	''' <summary>
	''' Builds SQL Update Query String.
	''' </summary>
	''' <param name="t_">Table to update.</param>
	''' <param name="update_keys">Fields to replace.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to update or not.</param>
	''' <returns>String.</returns>
	''' <example>
	''' BuildUpdateString("Table_Name", {"Field_To_Update_1"}, {"Field_To_Check_Condition_On_1"})
	''' </example>
	Public Shared Function BuildUpdateString(t_ As String, update_keys As Array, Optional where_keys As Array = Nothing) As String

		Dim v As String = "UPDATE " & t_ & " SET "

		For j As Integer = 0 To update_keys.Length - 1
			v &= update_keys(j) & "=@" & update_keys(j)
			If update_keys.Length > 1 And j <> update_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("

				For k As Integer = 0 To where_keys.Length - 1
					v &= where_keys(k) & "=@" & where_keys(k)
					If where_keys.Length > 1 And k <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
				v &= ")"
			End If
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Fields to select.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
	''' <returns>String.</returns>
	''' <example>
	''' BuildSelectString("Table_Name", {"Field_To_Select_1"}, {"Field_To_Check_Condition_On_1"})
	''' </example>
	Public Shared Function BuildSelectString(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional OrderByField As String = "PRIMARY_KEY_OR_OTHER") As String
		Dim v As String = "SELECT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField.Length > 0 Then v &= " ORDER BY " & OrderByField
		Return v
	End Function

	''' <summary>
	''' Builds SQL Count Query String.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
	''' <returns>String</returns>
	''' <example>
	''' BuildSelectString("Table_Name", {"Field_To_Check_Condition_On_1"})
	''' </example>
	Public Shared Function BuildCountString(t_ As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT COUNT(*) "
		'For i As Integer = 0 To select_params.Length - 1
		'	v &= select_params(i)
		'	If select_params.Length > 1 And i <> select_params.Length - 1 Then
		'		v &= ", "
		'	End If
		'Next
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function
	''' <summary>
	''' Builds SQL Select Top Query String.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
	''' <returns>String.</returns>
	''' <example>
	''' BuildTopString("Table_Name", {"Field_To_Check_Condition_On_1"})
	''' </example>
	Public Shared Function BuildTopString(t_ As String, Optional where_keys As Array = Nothing, Optional rows_to_select As Long = 1, Optional OrderByField As String = "PRIMARY_KEY_OR_OTHER") As String
		'		Dim v As String = "SELECT TOP 1 * "
		Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
		'For i As Integer = 0 To select_params.Length - 1
		'	v &= select_params(i)
		'	If select_params.Length > 1 And i <> select_params.Length - 1 Then
		'		v &= ", "
		'	End If
		'Next
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField.Length > 0 Then v &= " ORDER BY " & OrderByField

		Return v
	End Function
	''' <summary>
	''' Builds Select Max string.
	''' </summary>
	''' <param name="t_"></param>
	''' <param name="where_keys"></param>
	''' <param name="Max_Field"></param>
	''' <returns></returns>
	Public Shared Function BuildMaxString(t_ As String, Optional where_keys As Array = Nothing, Optional Max_Field As String = "PRIMARY_KEY_OR_OTHER") As String
		Dim v As String = "SELECT MAX (" & Max_Field & ")"
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function


#Region "Retrieval"
	''' <summary>
	''' Executes SQL statement for a single value.
	''' </summary>
	''' <param name="query_">SQL Query. You can use BuildSelectString, BuildUpdateString, BuildInsertString, BuildCountString, BuildTopString instead.</param>
	''' <param name="connection_string">Connection String.</param>
	''' <param name="parameters_keys_values_">Parameters.</param>
	''' <param name="return_type_is_string">Always return string.</param>
	''' <returns>Scalar value.</returns>
	Public Shared Function QData(query_ As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional return_type_is_string As Boolean = True)
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = parameters_keys_values_

		Dim con As SqlConnection = New SqlConnection(connection_string)
		Dim cmd As SqlCommand = New SqlCommand(query_, con)

		If select_parameter_keys_values IsNot Nothing Then
			For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
				cmd.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
			Next
		End If

		Try
			Using con
				con.Open()
				If return_type_is_string Then
					Return CType(cmd.ExecuteScalar(), String)
				End If
			End Using
		Catch
		End Try

	End Function

	''' <summary>
	''' Checks if a field exists w/ or w/o specified condition.
	''' </summary>
	''' <param name="t_">Table to perform operation on.</param>
	''' <param name="connection_string">Connection String.</param>
	''' <param name="where_keys">List of parameters.</param>
	''' <param name="where_keys_values">List of parameters and their values.</param>
	''' <returns>True if it exists, False otherwise.</returns>

	Public Shared Function QExists(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing) As Boolean
		Return QData(BuildCountString(t_, where_keys), connection_string, where_keys_values) > 0
	End Function


	''' <summary>
	''' Highest value of Max_Field for given choices w/ or w/o specified condition.
	''' </summary>
	''' <param name="t_">Table to perform operation on.</param>
	''' <param name="connection_string">Connection String.</param>
	''' <param name="where_keys">List of parameters.</param>
	''' <param name="where_keys_values">List of parameters and their values.</param>
	''' <param name="Max_Field">Field to use as maximum.</param>
	''' <returns>Value of Max_Field.</returns>
	Public Shared Function QMax(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing, Optional Max_Field As String = "PRIMARY_KEY_OR_OTHER")
		Return QData(BuildMaxString(t_, where_keys, Max_Field), connection_string, where_keys_values)
	End Function


#End Region

End Class
