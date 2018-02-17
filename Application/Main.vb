'	MAINTENANCE SYSTEM

'	Created by Trevor Xander

Imports System.IO

Public Class Main

    Structure Theme
        Dim background As System.Drawing.Bitmap
        Dim buttonActiveColor As Color
        Dim buttonInactiveColor As Color
        Dim cellBackColor As Color
        Dim cellForeColor As Color
    End Structure

    Structure Cell
        Dim row As Integer
        Dim column As Integer
    End Structure

    Dim userID As String
    Dim userName As String
    Dim password As String
    Dim fullName As String
    Dim isAdmin As Boolean

    Dim themes(4) As Theme
    Dim inactiveColor As Color
    Dim activeColor As Color

    Dim curPage As String

    Dim button As New Hashtable

    Dim adminCommand As String
    Dim userCommand As String
    Dim adminSearchCommand As String
    Dim userSearchCommand As String

    Dim adapter As OleDb.OleDbDataAdapter
    Dim dataSet As New DataSet

    Dim hasEdited As Boolean

    Dim editedCells(0) As Cell
    Dim newRows As New Hashtable

    Dim confirm As Boolean
    Dim deny As Boolean

    Sub dialog(ByVal actionText As String)

        'To prevent user from navigating to another screen/page
        name_button.Enabled = False
        nav_button.Enabled = False

        centerControl(dialog_panel)
        dialog_text.Text = actionText 'Displays text describing the action
        dialog_text.Location = New Point(((dialog_panel.Width - dialog_text.Width) \ 2), _
                 (dialog_text.Location.Y))
        dialog_panel.BringToFront()
        dialog_panel.Visible = True

        'Initializing      
        confirm = False

        'Halts program till user clicks on the button
        While confirm = False
            Application.DoEvents() 'To keep the program from being unresponsive
        End While

        dialog_panel.Visible = False

        'Re-enabling controls for free navigation
        name_button.Enabled = True
        nav_button.Enabled = True

    End Sub


    'To confirm with the user if an action needs to be done
    Function confirmAction(ByVal actionText As String) As Boolean

        'To prevent user from navigating to another screen/page
        resetPage()

        centerControl(confirmAction_panel)
        actionText_text.Text = actionText 'Displays text describing the action
        actionText_text.Location = New Point(((confirmAction_panel.Width - actionText_text.Width) \ 2), _
                  (actionText_text.Location.Y))
        confirmAction_panel.Visible = True

        'Initializing      
        confirm = False
        deny = False

        'Halts program till user clicks on a button
        While confirm = False And deny = False
            Application.DoEvents() 'To keep the program from being unresponsive
        End While

        confirmAction_panel.Visible = False

        'Re-enabling controls for free navigation
        name_button.Enabled = True
        nav_button.Enabled = True

        'Returns value depending on the button clicked
        If confirm = True Then
            Return True
        Else
            Return False
        End If

    End Function

    Sub logsInit()

        Dim type_combo As New DataGridViewComboBoxColumn


        logs_dataGrid.Location = New Point(0, header_panel.Height + settings_button.Height + 14)
        logs_dataGrid.Width = Me.Width - 17
        logs_dataGrid.Height = Me.Height - 125

        logs_dataGrid.Columns(0).Visible = False

        logs_dataGrid.Columns(5).HeaderText = "Completed"


        If userID = 0 Then
            logs_dataGrid.Columns(7).HeaderText = "User Added"
            logs_dataGrid.Columns(7).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            logs_dataGrid.Columns(7).ReadOnly = True
        End If

        logs_dataGrid.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        logs_dataGrid.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        logs_dataGrid.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        logs_dataGrid.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        logs_dataGrid.Columns(6).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells

        logs_dataGrid.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        logs_dataGrid.Columns(6).ReadOnly = True


        logs_dataGrid.RowHeadersWidthSizeMode = _
          DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders


        logs_dataGrid.Refresh()


    End Sub

    'Centers controls relative to the form
    Sub centerControl(ByVal ctrl As Control)

        If ctrl.Location.Y > header_panel.Height Then
            ctrl.Location = New Point((ClientSize.Width - ctrl.Width) \ 2, _
                   (ClientSize.Height - ctrl.Height) \ 2)
        End If

    End Sub

    'Loads and displays the database on a dataGrid
    Sub dbLoad()

        dataSet = New DataSet

        'Clearing the dataGrid
        logs_dataGrid.DataSource = Nothing

        'Selects all logs from all users
        adminCommand = ("SELECT refID, Title, Description, Type, Cost, isComplete, Date, Users.Name FROM Logs " & _
            "INNER JOIN Users ON Logs.user_id=Users.user_id " & _
            "ORDER BY isComplete DESC, Date DESC, type ASC, title ASC;")

        'Selects only logs added by the current user
        userCommand = ("SELECT refID, Title, Description, Type, Cost, isComplete, Date FROM Logs " & _
            "WHERE user_id ='" & userID & "' " & _
            "ORDER BY isComplete DESC, Date DESC, type ASC, title ASC ;")

        If isAdmin Then
            adapter = New OleDb.OleDbDataAdapter(adminCommand, conn)
        Else
            adapter = New OleDb.OleDbDataAdapter(userCommand, conn)
        End If

        adapter.Fill(dataSet)
        logs_dataGrid.DataSource = dataSet.Tables(0)

        logsInit()
        logs_dataGrid.Visible = True

    End Sub

    Sub changePage()

        'Displays name of current page on button
        nav_button.Text = curPage

        resetPage()

        'To confirm whether user wants to save changes
        If hasEdited Then
            Dim obj As New Object
            save_button_Click(obj, New System.EventArgs)
        End If

        'Shows controls for the selected page
        If curPage = "Home" Then

            nav_button.Visible = True
            name_button.Visible = True
            home_button.ForeColor = Color.White
            home_button.Cursor = Cursors.Default

            statsInit()

            centerControl(stats_panel)
            stats_panel.Visible = True
            '
        ElseIf curPage = "Login" Then
            centerControl(login_panel)
            login_panel.Visible = True

        ElseIf curPage = "Logs" Then
            nav_button.Visible = True
            name_button.Visible = True

            logs_button.ForeColor = Color.White
            logs_button.Cursor = Cursors.Default

            dbLoad()

            search_text.Visible = True
            search_text.Location = New Point(((ClientSize.Width - search_text.Width) \ 2), _
                     (logs_dataGrid.Location.Y - search_text.Height))

            save_button.Visible = True
            save_button.Location = New Point((search_text.Location.X + search_text.Width), _
                     search_text.Location.Y)

        ElseIf curPage = "Settings" Then
            nav_button.Visible = True
            name_button.Visible = True

            settings_button.ForeColor = Color.White
            settings_button.Cursor = Cursors.Default

            centerControl(settings_panel)
            settings_panel.Visible = True

            If userID = 0 Then
                changeadmin_button.Visible = True
                newUser_button.Visible = True
            Else
                changeadmin_button.Visible = False
                newUser_button.Visible = False
            End If
        End If

    End Sub

    'Displays panel containing details of the log
    Sub statsInit()

        Dim ds As New DataSet
        Dim adapter As OleDb.OleDbDataAdapter

        If userID = 0 Then
            'Admin is able to see details of all logs             
            adapter = New OleDb.OleDbDataAdapter("SELECT refID FROM Logs " & _
                      "WHERE True", conn)
        Else
            'User is able to see details of the logs added by him
            adapter = New OleDb.OleDbDataAdapter("SELECT refID FROM Logs " & _
                      "WHERE user_Id='" & userID & "'", conn)
        End If

        adapter.Fill(ds)
        totalRecords_label.Text = ds.Tables(0).Rows.Count
        ds.Clear()

        If userID = 0 Then
            adapter = New OleDb.OleDbDataAdapter("SELECT refID FROM Logs " & _
                      "WHERE isComplete = True", conn)
        Else
            adapter = New OleDb.OleDbDataAdapter("SELECT refID FROM Logs " & _
                      "WHERE isComplete = True AND user_ID='" & userID & "'", conn)
        End If

        adapter.Fill(ds)
        CompletedRecords_label.Text = ds.Tables(0).Rows.Count
        IncompleteRecords_label.Text = Val(totalRecords_label.Text) - _
                  Val(CompletedRecords_label.Text)

    End Sub

    Sub resetPage()

        'Clears the page of controls
        nav_button.Visible = False
        name_button.Visible = False
        save_button.Visible = False
        add_panel.Visible = False
        search_text.Visible = False
        logs_dataGrid.Visible = False
        settings_panel.Visible = False
        changePassword_panel.Visible = False
        changeAdmin_panel.Visible = False
        login_panel.Visible = False
        stats_panel.Visible = False

        'Resets the color of buttons
        home_button.ForeColor = inactiveColor
        settings_button.ForeColor = inactiveColor
        logs_button.ForeColor = inactiveColor

        'Resets cursor types of buttons
        settings_button.Cursor = Cursors.Hand
        home_button.Cursor = Cursors.Hand
        logs_button.Cursor = Cursors.Hand


    End Sub

    Sub controlsInit()

        Me.MinimumSize = New Point(800, 640)


        nav_button.Location = New Point(125, 4)

        name_button.Location = New Point((Me.Size.Width.ToString) - (name_button.Size.Width + 15), 4)

        header_panel.Location = New Point(0, 0)
        header_panel.BackColor = Color.FromArgb(36, 36, 36)
        header_panel.Size = New Size(Me.Size.Width, 40)

        logout_panel.BackColor = header_panel.BackColor
        logout_panel.Location = New Point((Me.Size.Width) - (110), name_button.Location.Y + 40)

        nav_panel.Location = New Point(header_panel.Location.X, logout_panel.Location.Y)
        nav_panel.BackColor = header_panel.BackColor
        nav_panel.Size = New Size((home_button.Width + logs_button.Width + settings_button.Width) + 5, _
                home_button.Height + 5)

        logs_dataGrid.Location = New Point(0, header_panel.Height + settings_button.Height + 14)

        settings_panel.BackColor = Color.FromArgb(150, 0, 0, 0)

        changePassword_panel.BackColor = Color.FromArgb(150, 0, 0, 0)

        changeAdmin_panel.BackColor = Color.FromArgb(150, 0, 0, 0)

        login_panel.BackColor = Color.FromArgb(150, 0, 0, 0)

        confirmAction_panel.BackColor = Color.FromArgb(150, 0, 0, 0)

        dialog_panel.BackColor = Color.FromArgb(200, 0, 0, 0)

        stats_panel.BackColor = Color.FromArgb(150, 0, 0, 0)

        newUser_panel.BackColor = Color.FromArgb(150, 0, 0, 0)

        'Storing all buttons in a hashtable
        button.Add(nav_button.Name.ToString, nav_button)
        button.Add(name_button.Name, name_button)
        button.Add(logout_button.Name, logout_button)
        button.Add(logs_button.Name, logs_button)
        button.Add(home_button.Name, home_button)
        button.Add(settings_button.Name, settings_button)
        button.Add(changePassword_button.Name, changePassword_button)
        button.Add(changeadmin_button.Name, changeadmin_button)
        button.Add(close_add.Name, close_add)
        button.Add(save_button.Name, save_button)
        button.Add(backChangePassword_button.Name, backChangePassword_button)
        button.Add(doneChangePassword_button.Name, doneChangePassword_button)
        button.Add(doneChangeAdmin_button.Name, doneChangeAdmin_button)
        button.Add(backChangeAdmin_button.Name, backChangeAdmin_button)
        button.Add(close_button.Name, close_button)
        button.Add(login_button.Name, login_button)
        button.Add(confirmAction_button.Name, confirmAction_button)
        button.Add(denyAction_button.Name, denyAction_button)
        button.Add(backNewUser_button.Name, backNewUser_button)
        button.Add(doneNewUser_button.Name, doneNewUser_button)
        button.Add(newUser_button.Name, newUser_button)
        button.Add(ok_button.Name, ok_button)

        'Iterates through 'button' hashtable and
        'adds MouseHover and MouseLeave event handlers
        'to sub routines
        Dim entry As DictionaryEntry
        For Each entry In button

            Dim btn As New Label
            btn = entry.Value

            'To change color of button on these events
            AddHandler btn.MouseHover, AddressOf button_MouseHovers
            AddHandler btn.MouseLeave, AddressOf button_MouseLeave

            btn.Cursor = Cursors.Hand
            btn.ForeColor = inactiveColor
            btn.BackColor = Color.Transparent
        Next

    End Sub

    'Chooses the wallpaper and button colors for the form
    Sub themeInit()

        'Storing themes into an array    
        Dim blueOrange As New Theme
        blueOrange.background = My.Resources.blueOrange
        blueOrange.buttonActiveColor = Color.FromArgb(133, 82, 19)
        blueOrange.buttonInactiveColor = Color.DarkOrange
        blueOrange.cellBackColor = Color.White
        blueOrange.cellForeColor = Color.Black
        themes(0) = (blueOrange)

        Dim bluePink As Theme
        bluePink.background = My.Resources.bluePink
        bluePink.buttonActiveColor = Color.FromArgb(139, 17, 11)
        bluePink.buttonInactiveColor = Color.FromArgb(249, 46, 49)
        themes(1) = (bluePink)

        Dim green As Theme
        green.background = My.Resources.green
        green.buttonActiveColor = Color.FromArgb(0, 94, 53)
        green.buttonInactiveColor = Color.FromArgb(21, 183, 113)
        themes(2) = (green)

        Dim greenBlue As Theme
        greenBlue.background = My.Resources.greenBlue
        greenBlue.buttonActiveColor = Color.FromArgb(8, 55, 106)
        greenBlue.buttonInactiveColor = Color.FromArgb(3, 114, 237)
        themes(3) = (greenBlue)

        Dim yellowBlue As Theme
        yellowBlue.background = My.Resources.yellowBlue
        yellowBlue.buttonActiveColor = Color.FromArgb(192, 121, 4)
        yellowBlue.buttonInactiveColor = Color.FromArgb(251, 199, 1)
        themes(4) = (yellowBlue)

        Dim curTheme As Theme
        curTheme = themes(nextTheme)

        Me.BackgroundImage = curTheme.background
        inactiveColor = curTheme.buttonInactiveColor
        activeColor = curTheme.buttonActiveColor

        'Frees memory
        ReDim themes(0)

    End Sub

    Function nextTheme() As Short

        'Reads number from file that determines
        'the previous theme used
        Dim StreamReader As StreamReader

        StreamReader = _
        New StreamReader((Application.StartupPath.ToLower _
            .Replace("\bin\debug", "") _
            .Replace("\bin\release", "")) _
             + "\theme.txt")

        Dim themeNumber As Short
        themeNumber = CShort(StreamReader.ReadLine)
        StreamReader.Close()

        'Obtains the value of the next theme      
        themeNumber += 1

        'Rounds back to 0 if the number exceeds
        'the total number of themes
        If themeNumber > themes.length - 1 Then
            themeNumber = 0
        End If

        'Writes the new number into the file    
        Dim StreamWriter = _
        New StreamWriter((Application.StartupPath.ToLower _
            .Replace("\bin\debug", "") _
            .Replace("\bin\release", "")) _
             + "\theme.txt")

        StreamWriter.WriteLine(themeNumber)
        StreamWriter.Close()

        Return themeNumber

    End Function

    Private Sub Main_Load _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles MyBase.Load

        Me.Size = New Size(1040, 640)

        'Sets the theme
        themeInit()

        'Initializes controls
        controlsInit()

        'Sets current page to login
        curPage = "Login"
        changePage()

    End Sub

    Private Sub Admin_Resize _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles Me.Resize

        header_panel.Size = New Size(Me.Size.Width, 40)
        name_button.Location = New Point((Me.Size.Width.ToString) - _
         (name_button.Size.Width + 15), _
          name_button.Location.Y)

        If logout_panel.Visible Then
            logout_panel.Visible = False
        End If

        If nav_panel.Visible Then
            nav_panel.Visible = False
        End If

        If logs_dataGrid.Visible Then
            logs_dataGrid.Width = Me.Width - 17
            logs_dataGrid.Height = Me.Height - 125
        End If

        If add_panel.Visible Then
            centerControl(add_panel)
        End If

        If search_text.Visible Then
            search_text.Location = _
             New Point(((ClientSize.Width - search_text.Width) \ 2), _
               (logs_dataGrid.Location.Y - search_text.Height))
        End If

        If save_button.Visible Then
            save_button.Location = _
             New Point((search_text.Location.X + search_text.Width), _
               search_text.Location.Y)
        End If

        If settings_panel.Visible Then
            centerControl(settings_panel)
        End If

        If changePassword_panel.Visible Then
            centerControl(changePassword_panel)
        End If

        If changeAdmin_panel.Visible Then
            centerControl(changeAdmin_panel)
        End If

        If login_panel.Visible Then
            centerControl(login_panel)
        End If

        If confirmAction_panel.Visible Then
            centerControl(confirmAction_panel)
        End If

        If stats_panel.Visible Then
            centerControl(stats_panel)
        End If

        If dialog_panel.Visible Then
            centerControl(dialog_panel)
        End If

        If newUser_panel.Visible Then
            centerControl(newUser_panel)
        End If

    End Sub

    'Button changes color on MouseHover
    Private Sub button_MouseHovers _
    (ByVal sender As Object, ByVal e As System.EventArgs)

        Dim btn As Label = DirectCast(sender, Label) 'Casts sender object to type 'label'

        If Not (btn.Text = curPage And btn.Name.ToString <> ("nav_button")) Then
            button(btn.Name).ForeColor = activeColor
        End If

    End Sub

    'Button changes color on MouseLeave
    Private Sub button_MouseLeave _
    (ByVal sender As Object, ByVal e As System.EventArgs)

        Dim btn As Label = DirectCast(sender, Label)
        If Not (btn.Text = curPage And btn.Name.ToString <> ("nav_button")) Then
            button(btn.Name).ForeColor = inactiveColor
        End If

    End Sub

    Private Sub logs_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles logs_button.Click

        If curPage <> "Logs" Then
            curPage = "Logs"
            changePage()
        End If

    End Sub

    Private Sub home_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles home_button.Click

        If curPage <> "Home" Then
            curPage = "Home"
            changePage()
        End If

    End Sub

    Private Sub settings_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles settings_button.Click

        If curPage <> "Settings" Then
            curPage = "Settings"
            changePage()
        End If

    End Sub

    Private Sub add_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs)

        logs_dataGrid.Visible = False
        add_panel.Visible = True

        add_panel.BringToFront()
        centerControl(add_panel)

        Dim dataSet As New DataSet
        Dim adapter As New OleDb.OleDbDataAdapter("SELECT DISTINCT Type from Logs ORDER BY Type ASC ", conn)

        adapter.Fill(dataSet)

        type_combo.Items.Clear()

        Dim count As Integer

        For count = 0 To dataSet.Tables(0).Rows.Count - 1

            type_combo.Items.Add(dataSet.Tables(0).Rows(count).Item(0).ToString)
        Next

        type_combo.Items.Add("New Category...")

    End Sub

    Private Sub search_text_TextChanged _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles search_text.TextChanged

        Dim adapter As OleDb.OleDbDataAdapter
        dataSet = New DataSet

        adminSearchCommand = ("SELECT refID, Title, Description, Type, Cost, isComplete, Date, Users.Name FROM Logs " & _
               "INNER JOIN Users ON Logs.user_id=Users.user_id " & _
               "WHERE (refID='" & search_text.Text & "' OR " & _
                   "Title Like '%" & search_text.Text & "%' OR " & _
                   "Type Like '%" & search_text.Text & "%' OR " & _
                   "description LIKE '%" & search_text.Text & "%' OR " & _
                   "Name LIKE '%" & search_text.Text & "%' OR " & _
                   "date LIKE '%" & search_text.Text & "%') " & _
                "ORDER BY isComplete DESC, Date DESC, type ASC, title ASC;")

        userSearchCommand = ("SELECT refID, Title, Description, Type, Cost, isComplete, Date FROM Logs " & _
              "WHERE user_id ='" & userID & "'AND " & _
                "(refID='" & search_text.Text & "' " & _
                "OR Title Like '%" & search_text.Text & "%' " & _
                "OR Type Like '%" & search_text.Text & "%' " & _
                "OR description LIKE '%" & search_text.Text & "%' " & _
                "OR date LIKE '%" & search_text.Text & "%' ) " & _
              "ORDER BY isComplete DESC, Date DESC, type ASC, title ASC ;")

        If search_text.Text <> "" Then
            'Filters data by the search text
            If isAdmin Then
                adapter = New OleDb.OleDbDataAdapter(adminSearchCommand, conn)
            Else
                adapter = New OleDb.OleDbDataAdapter(userSearchCommand, conn)
            End If
        Else
            'Unfiltered data is displayed when search text is blank
            If isAdmin Then
                adapter = New OleDb.OleDbDataAdapter(adminCommand, conn)
            Else
                adapter = New OleDb.OleDbDataAdapter(userCommand, conn)
            End If
        End If

        adapter.Fill(dataSet)
        logs_dataGrid.DataSource = Nothing
        logs_dataGrid.DataSource = dataSet.Tables(0)

        logsInit()

    End Sub

    'Gets the last reference ID
    Function getRefID() As String

        Dim refAdapter As OleDb.OleDbDataAdapter
        Dim refID As New DataSet

        refAdapter = New OleDb.OleDbDataAdapter("SELECT refID FROM Logs", conn)
        refAdapter.Fill(refID)

        Return (refID.Tables(0).Rows(refID.Tables(0).Rows.Count - 1).Item(0) + 1)

    End Function

    Sub saveChanges()

        logs_dataGrid.EndEdit()

        'Updates cells 
        If editedCells.Length > 0 Then
            Dim columnName As String
            Dim columnNumber As Short
            Dim rowNumber As Integer

            Dim refID As String
            Dim isComplete As Boolean
            Dim cost As Double
            Dim editedValue As String

            Dim updAdapter As New OleDb.OleDbCommand
            updAdapter.Connection = conn

            Dim count As Integer
            For count = 0 To editedCells.Length - 1
                columnNumber = editedCells(count).column
                rowNumber = editedCells(count).row
                refID = logs_dataGrid.Rows(rowNumber).Cells(0).Value
                columnName = dataSet.Tables(0).Columns(columnNumber).ColumnName

                'Using variables of appropriate type depening on the column that needs to updated

                If columnNumber = 4 Then 'For numeric
                    cost = logs_dataGrid.Rows(rowNumber).Cells(columnNumber).Value
                    updAdapter.CommandText = ("UPDATE Logs SET [" & columnName & "] = " & cost & " " & _
                            "WHERE refID = '" & refID & "' ")

                ElseIf columnNumber = 5 Then 'For boolean
                    isComplete = logs_dataGrid.Rows(rowNumber).Cells(columnNumber).Value
                    updAdapter.CommandText = ("UPDATE Logs SET [" & columnName & "] = " & isComplete & " " & _
                            "WHERE refID = '" & refID & "' ")

                Else 'For string
                    editedValue = logs_dataGrid.Rows(rowNumber).Cells(columnNumber).Value
                    updAdapter.CommandText = ("UPDATE Logs SET [" & columnName & "] = '" & editedValue & "' " & _
                            "WHERE refID = '" & refID & "' ")
                End If

                updAdapter.ExecuteNonQuery()
            Next

        End If

        'Adds new logs
        If newRows.Count > 0 Then
            Dim insrtAdapter As New OleDb.OleDbCommand
            insrtAdapter.Connection = conn

            'Iterates through newRows adding each one into database
            Dim entry As New DictionaryEntry
            For Each entry In newRows

                Dim row As Windows.Forms.DataGridViewRow
                row = entry.Value

                'Sets default values if user left the field blank
                If row.Cells(1).Value.ToString = DBNull.Value.ToString Then
                    row.Cells(1).Value = "Untitled"
                End If
                If row.Cells(2).Value.ToString = DBNull.Value.ToString Then
                    row.Cells(2).Value = "Not Included"
                End If
                If row.Cells(3).Value.ToString = DBNull.Value.ToString Then
                    row.Cells(3).Value = "General"
                End If
                If row.Cells(4).Value.ToString = DBNull.Value.ToString Then
                    row.Cells(4).Value = 0
                End If
                If row.Cells(5).Value.ToString = DBNull.Value.ToString Then
                    row.Cells(5).Value = False
                End If



                insrtAdapter.CommandText = ("INSERT INTO Logs " & _
                          "([refID], " & _
                          "[title], " & _
                          "[description], " & _
                          "[type], " & _
                          "[Cost], " & _
                          "[isComplete], " & _
                          "[Date], " & _
                          "[user_id]) " & _
                       "VALUES ('" & getRefID() + 1 & "', " & _
                          "'" & row.Cells(1).Value & "', " & _
                          "'" & row.Cells(2).Value & "', " & _
                          "'" & row.Cells(3).Value & "', " & _
                          "'" & row.Cells(4).Value & "', " & _
                          "" & row.Cells(5).Value & ", " & _
                          "'" & Now.Date.ToShortDateString & "', " & _
                          "'" & userID & "' );")
                MessageBox.Show(insrtAdapter.CommandText)
                insrtAdapter.ExecuteNonQuery()

            Next

        End If

    End Sub

    Private Sub save_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles save_button.Click

        'Verify whether user intends to save changes
        If confirmAction("Do you want to save changes?") Then
            saveChanges()
        End If

        hasEdited = False
        changePage()
        ReDim editedCells(0)
        newRows.Clear()

    End Sub

    Private Sub logs_dataGrid_CellContentClick _
    (ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) _
     Handles logs_dataGrid.CellContentClick

        If e.ColumnIndex = 5 Then
            logs_dataGrid_CellEndEdit(logs_dataGrid, e)
        End If

    End Sub

    'Keeps track of the edits the user has made
    Private Sub logs_dataGrid_CellEndEdit _
    (ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) _
     Handles logs_dataGrid.CellEndEdit

        hasEdited = True

        'If edited cell was in a new row
        If newRows.Contains(e.RowIndex) Then
            Dim row As Windows.Forms.DataGridViewRow
            row = newRows(e.RowIndex)
            newRows.Remove(e.RowIndex)
            newRows.Add(e.RowIndex, row)

            'If edited cell was from an existing row
        Else
            Dim editedCell As New Cell
            editedCell.column = e.ColumnIndex
            editedCell.row = e.RowIndex

            editedCells(editedCells.Length - 1) = editedCell
            ReDim Preserve editedCells(editedCells.Length)
        End If

    End Sub

    Private Sub logs_data_UserAddedRow _
    (ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) _
     Handles logs_dataGrid.UserAddedRow

        hasEdited = True

        'Keeps track of the new rows added
        newRows.Add(e.Row.Index - 1, logs_dataGrid.Rows(e.Row.Index - 1))

        'Fills in the data for 'date' and the 'userAdded' columns 
        logs_dataGrid.Rows(e.Row.Index - 1).Cells(6).Value = Now.Date.ToShortDateString
        If userID = 0 Then
            logs_dataGrid.Rows(e.Row.Index - 1).Cells(7).Value = fullName
        End If

    End Sub

    Private Sub changePassword_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles changePassword_button.Click

        changePassword_panel.Location = settings_panel.Location
        settings_panel.Visible = False
        changePassword_panel.Visible = True

    End Sub

    Private Sub back_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles backChangePassword_button.Click

        settings_panel.Location = changePassword_panel.Location
        changePassword_panel.Visible = False
        settings_panel.Visible = True

    End Sub

    Private Sub backChangeAdmin_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles backChangeAdmin_button.Click

        settings_panel.Location = changeAdmin_panel.Location
        changeAdmin_panel.Visible = False
        settings_panel.Visible = True

    End Sub

    Private Sub changeadmin_button_Click _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles changeadmin_button.Click

        changeAdmin_panel.Location = settings_panel.Location
        settings_panel.Visible = False
        changeAdmin_panel.Visible = True

        Dim allUsers As New DataSet
        Dim adapter As New OleDb.OleDbDataAdapter("SELECT name FROM Users " & _
                    "WHERE user_id <> '0' " & _
                    "ORDER BY name ASC ;", conn)

        adapter.Fill(allUsers)

        newAdmin_combo.Items.Clear()
        newAdmin_combo.Text = "Choose New Admin"

        Dim count As Integer

        For count = 0 To allUsers.Tables(0).Rows.Count - 1

            newAdmin_combo.Items.Add(allUsers.Tables(0).Rows(count).Item(0).ToString)

        Next

    End Sub

    Sub changeAdmin()

        Dim updAdapter As New OleDb.OleDbCommand
        updAdapter.Connection = conn

        Dim newAdminUserID As String
        Dim newAdminName As String

        Dim user As New DataSet
        Dim adapter As New OleDb.OleDbDataAdapter("SELECT user_id FROM Users " & _
                    "WHERE name = '" & newAdmin_combo.Text & "' ;", conn)
        adapter.Fill(user)

        If newAdmin_combo.Text <> "Choose New Admin" Then
            newAdminName = newAdmin_combo.Text
            newAdminUserID = user.Tables(0).Rows(0).Item(0)

            'UserIDs are swapper with the admin and the chosen user
            'so that the new user is recognized as admin
            updAdapter.CommandText = ("UPDATE Users SET [user_id] = 'temp' " & _
                    "WHERE user_id = '0' ")
            updAdapter.ExecuteNonQuery()
            updAdapter.CommandText = ("UPDATE Users SET [user_id] = '0' " & _
                    "WHERE name = '" & newAdminName & "' ")
            updAdapter.ExecuteNonQuery()
            updAdapter.CommandText = ("UPDATE Users SET [user_id] = '" & newAdminUserID & "' " & _
                    "WHERE user_id = 'temp' ")
            updAdapter.ExecuteNonQuery()

            userID = newAdminUserID

            curPage = "Settings"
            changePage()
        End If

    End Sub

    Private Sub doneChangeAdmin_button_Click _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles doneChangeAdmin_button.Click

        changeAdmin_panel.Visible = False

        'Verify whether user intends to change admin
        If confirmAction("Do you want to change the admin?") Then
            changeAdmin()
            changePage()
        Else
            'Returns to the 'change admin' page
            changePage()
            Dim obj As New Object
            changeadmin_button_Click(obj, New System.EventArgs)
        End If

    End Sub

    Private Sub doneChangePassword_button_Click _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles doneChangePassword_button.Click

        If oldPassword_text.Text = password And _
           newPassword_text.Text = reenterPassword_text.Text Then

            If confirmAction("Do you want to change the password?") Then
                Dim updAdapter As New OleDb.OleDbCommand
                updAdapter.Connection = conn
                updAdapter.CommandText = ("UPDATE Users SET [password]='" & newPassword_text.Text & "' " & _
                       "WHERE user_id='" & userID & "'")
                updAdapter.ExecuteNonQuery()
                changePage()
            End If

        ElseIf oldPassword_text.Text <> password Then
            dialog("Incorrect Password")
        Else
            dialog("Passwords do not match")
        End If

    End Sub

    'Closes the form
    Private Sub close_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles close_button.Click

        Me.Close()

    End Sub

    Private Sub login_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles login_button.Click

        If loginUserName_text.Text = "" Or loginPassword_text.Text = "" Then
            dialog("Fill all the fields")
            Exit Sub
        End If

        Dim user As DataSet
        Dim loginAdapter = _
        New OleDb.OleDbDataAdapter("SELECT user_id, username, password, name FROM Users " & _
                  "WHERE username= '" & loginUserName_text.Text & "' AND " & _
                        "password = '" & loginPassword_text.Text & "' ;", conn)

        user = New DataSet
        loginAdapter.Fill(user)

        'Check if username password combination exists
        If user.Tables(0).Rows.Count > 0 Then

            userID = user.Tables(0).Rows(0).Item(0)
            userName = user.Tables(0).Rows(0).Item(1)
            password = user.Tables(0).Rows(0).Item(2)
            fullName = user.Tables(0).Rows(0).Item(3)

            If userID = 0 Then
                isAdmin = True
            Else
                isAdmin = False
            End If


            'Extracts first name from full name
            Dim firstName As String
            firstName = fullName.Remove(fullName.IndexOf(" "), _
               (fullName.Length) - fullName.IndexOf(" "))

            name_button.Text = firstName

            displayedFullName_label.Text = fullName
            displayedUserName_label.Text = userName

            'Opens connection to database
            conn.Open()

            curPage = "Home"
            changePage()
        Else
            dialog("Incorrect username password combination")
            loginPassword_text.Text = ""
        End If


    End Sub

    Private Sub logout_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles logout_button.Click

        'Closes connection to database 
        conn.Close()

        curPage = "Login"
        changePage()

    End Sub

    Private Sub confirmAction_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles confirmAction_button.Click

        confirm = True

    End Sub

    Private Sub denyAction_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles denyAction_button.Click

        deny = True

    End Sub

    Private Sub nav_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles nav_button.Click

        If nav_panel.Visible Then
            nav_panel.Visible = False
        Else
            nav_panel.Visible = True
        End If

    End Sub

    Private Sub name_button_Click _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles name_button.Click

        If logout_panel.Visible Then
            logout_panel.Visible = False
        Else
            logout_panel.Visible = True
        End If

    End Sub

    Private Sub ok_button_Click _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles ok_button.Click

        confirm = True

    End Sub

    Private Sub newUser_button_Click _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles newUser_button.Click

        newUser_panel.Location = settings_panel.Location
        settings_panel.Visible = False
        newUser_panel.Visible = True

    End Sub

    Private Sub backNewUser_button_Click _
    (ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles backNewUser_button.Click

        settings_panel.Location = newUser_panel.Location
        settings_panel.Visible = True
        newUser_panel.Visible = False

    End Sub

    Function validateNewUser() As Boolean



    End Function

    Private Sub doneNewUser_button_Click _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles doneNewUser_button.Click

        If validateNewUser() Then
            Dim insrtAdapter As New OleDb.OleDbCommand
            insrtAdapter.Connection = conn
        End If

    End Sub

End Class
