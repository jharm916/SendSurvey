Imports System
Imports System.ComponentModel
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports Microsoft.SharePoint
Imports Microsoft.SharePoint.WebControls
Imports System.Text.RegularExpressions
Imports Microsoft.SharePoint.Utilities
Imports System.Collections.Specialized
Imports System.Net.Mail
Imports Microsoft.Office.Interop.Outlook

<ToolboxItemAttribute(false)> _
Public Class create_survey_3
    Inherits WebPart

    Public tb1, tb2 As DataGrid
    Public dt, dt2 As DataTable
    Public list As CheckBoxList
    Public allMaps As List(Of Map)
    Public allGroupMaps As List(Of Map)
    Public contr As ControlCollection
    Public _sourceDropDownList As DropDownList
    Public emailSubject As TextBox = Nothing
    Public emailBody As TextBox = Nothing
    Public addRecipients As TextBox
    Public addedEmails As New List(Of String)
    Public mailingList As New List(Of SPUser)

    Protected Overrides Sub OnLoad(e As EventArgs)
        If Not Page.IsPostBack Then
            BindData()
        End If
    End Sub
    Protected Overrides Sub CreateChildControls()
        Dim output_one As New Label()
        output_one.Text = "Select a survey to issue:<br /><br />"
        Me.Controls.Add(output_one)

        'Create a DropDownList object
        _sourceDropDownList = New DropDownList()
        _sourceDropDownList.ID = "_sourceDropDownList"

        'Sets up aspects of the dropdown list
        LoadBoxData()
        Me.Controls.Add(_sourceDropDownList)

        Dim output_four As New Label()
        output_four.Text = "<br /><br />Select the users and groups to send the survey to:<br /><br />"
        Me.Controls.Add(output_four)

        tb1 = New DataGrid()
        tb1.AutoGenerateColumns = False

        list = New CheckBoxList()
        Dim userList As SPUserCollection = SPContext.Current.Web.AllUsers
        Dim groupList As SPGroupCollection = SPContext.Current.Web.Groups

        'Create a new TemplateColumn object.
        Dim tcol As New TemplateColumn()
        With tcol
            .HeaderText = "Select"
            .ItemTemplate = New DynamicItemTemplate_2()
        End With

        Dim copy As DynamicItemTemplate_2
        copy = tcol.ItemTemplate
        copy.SetControls(Me.Controls)
        copy.SetCheckBoxes(list)
        allMaps = New List(Of Map)
        copy.SetMap(allMaps)

        tb1.Columns.Add(tcol)
        tb1.Columns.Add(CreateBoundColumn("StringValue", "User"))
        tb1.Columns.Add(CreateBoundColumn("EmailValue", "Email"))

        If Not Page.IsPostBack Then
            tb1.DataBind()
        End If

        'Set up data
        dt = New DataTable()

        dt.Columns.Add(New DataColumn("IntegerValue", GetType(CheckBoxField)))
        dt.Columns.Add(New DataColumn("StringValue", GetType(String)))
        dt.Columns.Add(New DataColumn("EmailValue", GetType(String)))

        Dim dr As DataRow
        Dim k As Integer
        If (userList.Count > 0) Then
            For k = 0 To (userList.Count - 1)
                dr = dt.NewRow()
                'This needs to pull from user selection on page reload. 
                Dim field As CheckBoxField = New CheckBoxField()
                dr(0) = field
                dr(1) = userList(k).Name
                dr(2) = userList(k).Email
                dt.Rows.Add(dr)
                'Creates a new Map entry
            Next k
        Else
            'This is bad
        End If
        Dim dv As New DataView(dt)
        tb1.DataSource = dv
        tb1.DataBind()
        tb1.Width = 500
        Me.Controls.Add(tb1)

        Dim output_five As New Label()
        output_five.Text = "<br /><br />"
        Me.Controls.Add(output_five)

        'Create grid for groups
        tb2 = New DataGrid()
        tb2.AutoGenerateColumns = False
        list = New CheckBoxList()

        'Create a new TemplateColumn object.
        Dim tcol2 As New TemplateColumn()
        With tcol2
            .HeaderText = "Select"
            .ItemTemplate = New DynamicItemTemplate_3()
        End With

        Dim copy2 As DynamicItemTemplate_3
        copy2 = tcol2.ItemTemplate
        copy2.SetControls(Me.Controls)
        copy2.SetCheckBoxes(list)
        allGroupMaps = New List(Of Map)
        copy2.SetMap(allGroupMaps)

        tb2.Columns.Add(tcol2)
        tb2.Columns.Add(CreateBoundColumn("StringValue", "Group"))
        'tb2.Columns.Add(CreateBoundColumn("EmailValue", "Email"))

        If Not Page.IsPostBack Then
            tb2.DataBind()
        End If

        'Set up data
        dt2 = New DataTable()
        dt2.Columns.Add(New DataColumn("IntegerValue", GetType(CheckBoxField)))
        dt2.Columns.Add(New DataColumn("StringValue", GetType(String)))
        'dt2.Columns.Add(New DataColumn("EmailValue", GetType(String)))

        If (groupList.Count > 0) Then
            For k = 0 To (groupList.Count - 1)
                dr = dt2.NewRow()
                'This needs to pull from user selection on page reload. 
                Dim field As CheckBoxField = New CheckBoxField()
                dr(0) = field
                dr(1) = groupList(k).Name
                'dr(2) = ""
                dt2.Rows.Add(dr)
                'Creates a new Map entry
            Next k
        Else
            'This is bad
        End If

        Dim dv2 As New DataView(dt2)
        tb2.DataSource = dv2
        tb2.DataBind()
        tb2.Width = 500
        Me.Controls.Add(tb2)

        Dim output_six As New Label()
        output_six.Text = "<br /><br />Enter a message to send:<br /><br />"
        Me.Controls.Add(output_six)

        'Create controls for sending the email
        emailSubject = New TextBox()
        emailBody = New TextBox()
        emailSubject.Text = "- Email Subject -"
        emailSubject.TextMode = TextBoxMode.MultiLine
        emailSubject.Width = 500

        emailBody.Text = "- Email Body -"
        emailBody.TextMode = TextBoxMode.MultiLine
        emailBody.Width = 500
        emailBody.Rows = 6
        Me.Controls.Add(emailSubject)
        Dim output_nine As New Label()
        output_nine.Text = "<br />"
        Me.Controls.Add(output_nine)
        Me.Controls.Add(emailBody)

        Dim output_seven As New Label()
        output_seven.Text = "<br /><br />Enter the emails of additional recipients for the survey separated by spaces."
        output_seven.Text &= "<br /><b>Please Note: Entering anything other than email addresses and spaces will cause "
        output_seven.Text &= "an<br />error. </b>Example: ""user1@example.com user2@example.com user3@example.com"""
        output_seven.Text = output_seven.Text & "<br /><br />"
        Me.Controls.Add(output_seven)

        'Create control to add additional recipients to the email
        addRecipients = New TextBox()
        addRecipients.Text = "- Additional Recipients -"
        addRecipients.TextMode = TextBoxMode.MultiLine
        addRecipients.Width = 500
        addRecipients.Rows = 6
        Me.Controls.Add(addRecipients)

        Dim finalizeForm As Button = New Button()
        finalizeForm.Text = "Send Surveys"
        AddHandler finalizeForm.Click, AddressOf sendSurveys_Click

        Dim output_eight As New Label()
        output_eight.Text = "<br /><br />"
        Me.Controls.Add(output_eight)
        Me.Controls.Add(finalizeForm)

        MyBase.CreateChildControls()

    End Sub
    Public Sub BindData()

    End Sub

    'Fills the drop down box with items containing "Survey"
    Private Sub LoadBoxData()
        'If Me.Page.IsPostBack = False Then
        'If (SourceList IsNot Nothing And SourceList.Items IsNot Nothing) Then
        Dim index As Integer
        Dim collList_temp As SPListCollection = SPContext.Current.Web.Lists
        'Loop through elements of the collection, adding them all to the drop-down
        Dim newItem As ListItem = New ListItem()
        newItem.Text = "- Select Survey -"
        newItem.Value = 0.ToString()
        _sourceDropDownList.Items.Add(newItem)

        If (collList_temp.Count > 0) Then
            For index = 0 To (collList_temp.Count - 1)
                Dim currList As SPList = collList_temp.Item(index)
                newItem = New ListItem()
                'Set the new item's Text and Value
                If (collList_temp.Item(index).Title.Contains("Survey")) Then
                    'Text is the Title, Value is the index in the collection
                    newItem.Text = collList_temp.Item(index).Title
                    newItem.Value = index.ToString
                    _sourceDropDownList.Items.Add(newItem)
                End If
            Next index
        End If
    End Sub

    'Accesses user-selected data and submits the email.
    Public Sub sendSurveys_Click()
        Dim output As New Label()
        'Dim selectedSurvey As ListItem = Nothing
        Dim finalList As New List(Of SPUser)
        Dim finalize As New Label()
        'Dim surveyURL As String = "" 'The selected survey's URL
        'Creates the "to" field of the email
        Dim allEmails As New List(Of String)
        Dim msgTo As String = "Emails: "
        Dim selectedSurvey As ListItem = Nothing
        Dim surveyURL As String = ""

        Dim errorLabel As New Label
        errorLabel.Text = ""
        Dim send = True 'Change upon error
        'Checks if the user has provided valid input
        selectedSurvey = _sourceDropDownList.Items(_sourceDropDownList.SelectedIndex)
        If (String.Compare(selectedSurvey.Text, "- Select Survey -") = 0) Then
            errorLabel.Text = "<br />Please choose valid selections"
            send = False
        End If
        If (String.Compare(emailSubject.Text, "- Email Subject -") = 0) Then
            errorLabel.Text = "<br />Please choose valid selections"
            send = False
        End If
        If (String.Compare(emailBody.Text, "- Email Body -") = 0) Then
            errorLabel.Text = "<br />Please choose valid selections"
            send = False
        End If
        If (VerifyFormat() = False) Then
            errorLabel.Text = "<br />Please provide a valid format for additional recipients."
            send = False
        End If

        'Check if the input provided is correct
        If send = True Then
            'The format is correct, proceed
            'Construct the "allEmails" list, adding email addresses from selected users and groups, 
            'checking for duplicates. Then, add the "addedEmails" list to this list.
            For i = 0 To (allMaps.Count - 1)
                If allMaps(i).GetCheckBox().Checked = True Then
                    'Check if already added
                    If (Not allEmails.Contains(allMaps(i).GetUser().Email)) Then
                        allEmails.Add(allMaps(i).GetUser().Email)
                    End If
                End If
            Next
            For i = 0 To (allGroupMaps.Count - 1)
                If (allGroupMaps(i).GetCheckBox().Checked = True) Then
                    For j = 0 To (allGroupMaps(i).GetGroup().Users.Count - 1)
                        'Check if the user was added
                        If (Not allEmails.Contains(allGroupMaps(i).GetGroup().Users(j).Email)) Then
                            allEmails.Add(allGroupMaps(i).GetGroup().Users(j).Email)
                        End If
                    Next
                End If
            Next
            For i = 0 To (addedEmails.Count - 1)
                'Check if the user was added
                If (Not allEmails.Contains(addedEmails(i))) Then
                    allEmails.Add(addedEmails(i))
                End If
            Next
            'Construct the "msgTo" String which will be the "to" field of the email
            For i = 0 To (allEmails.Count - 1)
                msgTo &= allEmails(i)
                If (Not i = allEmails.Count - 1) Then
                    msgTo &= ", "
                End If
            Next
            'Provides a notice regarding the "to" field.
            Dim toField As New Label()
            toField.Text = "<br /><br />Note: The ""to"" field of the email is currently the following:<br />"
            toField.Text &= msgTo & "</b>"

            Me.Controls.Add(toField)
            'Finds the URL of the selected survey
            Dim showTest As New Label()
            surveyURL &= SPContext.Current.Site.HostName
            surveyURL &= SPContext.Current.Web.Lists(selectedSurvey.ToString).DefaultViewUrl
            'Adds the URL of the survey to the bottom of the email body
            emailBody.Text &= Environment.NewLine() & Environment.NewLine() & "Link to survey: " & surveyURL
            'Send the email containing a link to the selected survey
            SendEmail(emailSubject.Text, allEmails, emailBody.Text, allEmails.Count - 1)
        End If
        Me.Controls.Add(errorLabel)
    End Sub

    'Sends an email using Outlook
    Public Sub SendEmail(subject As String, sendTo As List(Of String), body As String, numUsers As Integer)
        'Open an Outlook email
        Try
            Dim outlookApp As New Application()
            Dim mailItem As MailItem = CType(outlookApp.CreateItem(OlItemType.olMailItem), MailItem)
            Dim currentUser As AddressEntry = outlookApp.Session.CurrentUser.AddressEntry
            mailItem.Subject = subject
            For i = 0 To numUsers
                mailItem.Recipients.Add(sendTo(i))
            Next
            mailItem.Body = body
            'mailItem.Display(True) <-- controls whether or not the email message being sent pops up
            mailItem.Send()
            Dim emailSuccess As New Label()
            emailSuccess.Text = "<br /><b>Email sent successfully. It is now visible in Outlook.</b>"
            Me.Controls.Add(emailSuccess)
        Catch ex As SystemException
            Dim emailOpenError As New Label()
            emailOpenError.Text = "<br />Error, cannot send email. Please verify that you are logged into Microsoft Outlook."
            Dim selectionMade As ListItem = _sourceDropDownList.Items(_sourceDropDownList.SelectedIndex)
            Dim surveyAlt As String = ""
            surveyAlt &= SPContext.Current.Site.HostName
            surveyAlt &= SPContext.Current.Web.Lists(selectionMade.ToString).DefaultViewUrl
            emailOpenError.Text &= "<br /><br /><b>Note: </b>Alternatively, you can paste the following URL to the survey into an email, "
            emailOpenError.Text &= "using the emails listed above.<br />"
            emailOpenError.Text &= surveyAlt
            Me.Controls.Add(emailOpenError)
        End Try
    End Sub

    'Checks the format of the additional recipients TextBox
    Public Function VerifyFormat()
        If (String.Compare(addRecipients.Text, "- Additional Recipients -") = 0) Then
            Return True
        End If
        If (String.Compare(addRecipients.Text, "") = 0) Then
            Return True
        End If

        'We have to manipulate the String and create a list of Strings
        Dim emailsString As String = addRecipients.Text
        Dim emails As New List(Of String)
        Dim flag As Boolean = False

        'Format: separated by spaces
        Dim done As Boolean = False
        'Remove leading and trailing spaces
        Dim arrStrings As String()
        Dim subString As String = (addRecipients.Text.Trim())
        arrStrings = Split(subString)

        'Remove extra white space
        For i = 0 To (arrStrings.Length - 1)
            arrStrings(i) = arrStrings(i).Replace(" ", String.Empty)
        Next

        For k = 0 To (arrStrings.Length - 1)
            'Get 'name@ex.com'
            Dim emailRegEx As New Regex("(\S+)@([^\.\s]+)(?:\.([^\.\s]+))+")
            Dim m As Match = emailRegEx.Match(arrStrings(k))
            If m.Success Then
                Dim output As String = ""
                output &= m.Groups(1).Value '& vbCrLf
                Dim lock = 0
                For n As Integer = 2 To m.Groups.Count - 1
                    Dim g As Group = m.Groups(n)
                    For Each c As Capture In g.Captures
                        If lock = 0 Then
                            output &= "@" & c.Value '& vbCrLf
                            lock = 1
                        Else
                            output &= "." & c.Value '& vbCrLf
                        End If
                    Next
                Next
                output = output.Trim()
                'Add the email address to the list
                emails.Add(output)
            Else
                Return False
            End If
        Next
        SetEmailList(emails)
        Return True
    End Function
    Public Sub SetEmailList(emails As List(Of String))
        addedEmails = emails
    End Sub

    Function CreateBoundColumn(DataFieldValue As String, HeaderTextValue As String) As BoundColumn
        ' This version of CreateBoundColumn method sets only the 
        ' DataField and HeaderText properties.
        ' Create a BoundColumn.
        Dim column As BoundColumn = New BoundColumn()

        ' Set the properties of the BoundColumn.
        column.DataField = DataFieldValue
        column.HeaderText = HeaderTextValue

        Return column
    End Function
    Function CreateCheckBoxColumn(DataFieldValue As Boolean, HeaderTextValue As String) As BoundColumn
        ' This version of CreateBoundColumn method sets only the 
        ' DataField and HeaderText properties.
        ' Create a BoundColumn.

        Dim column As BoundColumn = New BoundColumn()

        ' Set the properties of the BoundColumn.
        column.DataField = DataFieldValue.ToString
        column.HeaderText = HeaderTextValue

        Return column
    End Function

End Class

'Column template for a user
Public Class DynamicItemTemplate_2
    ' ITemplate - When implemented by a class, defines the Control object
    ' to which child controls and templates belong. These child controls 
    ' are in turn defined within an inline template.
    Implements ITemplate
    Public controls As ControlCollection
    Public checkBoxes As CheckBoxList
    Public maps As List(Of Map) = Nothing
    Public index = 0

    Public Overridable Overloads Sub InstantiateIn(ByVal container As Control) Implements ITemplate.InstantiateIn
        Dim userList As SPUserCollection = SPContext.Current.Web.AllUsers
        Dim groupList As SPGroupCollection = SPContext.Current.Web.Groups

        Dim newMap As New Map()
        'If (lock = 0) Then
        newMap.AddUser(userList(index))
        'Else
        'newMap.AddGroup(groupList(index))
        'End If

        ' InstantiateIn - When implemented by a class, defines the Control 
        ' object to which child controls and templates belong. These child 
        ' controls are, in turn, defined within an inline template.
        ' 
        ' Create an instance of a CheckBox object.
        Dim oCheckBox As CheckBox = New CheckBox()
        oCheckBox.ID = "box_" & index.ToString()
        index = index + 1
        'Initialize all check boxes to False
        newMap.SetChoice(False)
        newMap.AddCheckBox(oCheckBox)
        maps.Add(newMap)
        ' When the DataBinding event of the CheckBox fires, call the sub 
        ' BindCheckBox to properly bind.  
        AddHandler oCheckBox.CheckedChanged, AddressOf BindCheckBox
        'Add the CheckBox to the controls collection.
        container.Controls.Add(oCheckBox)
    End Sub
    Public Sub SetControls(contr As ControlCollection)
        controls = contr
    End Sub
    Public Sub SetCheckBoxes(list As CheckBoxList)
        checkBoxes = list
    End Sub
    Public Sub IncreaseCounter()
        index = index + 1
    End Sub
    Public Sub DecreaseCounter()
        index = index - 1
    End Sub
    Public Sub BindCheckBox(ByVal sender As Object, ByVal e As EventArgs)
        'Create a new instance of a CheckBox. 
        Dim oCheckBox As CheckBox = CType(sender, CheckBox)
        Dim container As DataGridItem = CType(oCheckBox.NamingContainer, DataGridItem)
        'Evaluate the data from the Grid item and set the Checked property 
        ' appropriatly
        Dim lab As New Label()
        For i = 0 To (maps.Count - 1)
            If String.Compare(maps(i).GetCheckBoxID(), oCheckBox.ID) = 0 Then
                maps(i).SetChoice(oCheckBox.Checked)
            End If
        Next

        controls.Add(lab)

    End Sub

    Public Sub SetMap(listPointer As List(Of Map))
        maps = listPointer
    End Sub
End Class

'Column template for a group
Public Class DynamicItemTemplate_3
    ' ITemplate - When implemented by a class, defines the Control object
    ' to which child controls and templates belong. These child controls 
    ' are in turn defined within an inline template.
    Implements ITemplate
    Public controls As ControlCollection
    Public checkBoxes As CheckBoxList
    Public maps As List(Of Map) = Nothing
    Public index = 0

    Public Overridable Overloads Sub InstantiateIn(ByVal container As Control) Implements ITemplate.InstantiateIn
        Dim groupList As SPGroupCollection = SPContext.Current.Web.Groups

        Dim newMap As New Map()
        newMap.AddGroup(groupList(index))

        ' InstantiateIn - When implemented by a class, defines the Control 
        ' object to which child controls and templates belong. These child 
        ' controls are, in turn, defined within an inline template.
        ' 
        ' Create an instance of a CheckBox object.
        Dim oCheckBox As CheckBox = New CheckBox()
        oCheckBox.ID = "box_" & index.ToString()
        index = index + 1
        'Initialize all check boxes to False
        newMap.SetChoice(False)
        newMap.AddCheckBox(oCheckBox)
        maps.Add(newMap)

        ' When the DataBinding event of the CheckBox fires, call the sub 
        ' BindCheckBox to properly bind.  
        AddHandler oCheckBox.CheckedChanged, AddressOf BindCheckBox
        'Add the CheckBox to the controls collection.
        container.Controls.Add(oCheckBox)
    End Sub
    Public Sub SetControls(contr As ControlCollection)
        controls = contr
    End Sub
    Public Sub SetCheckBoxes(list As CheckBoxList)
        checkBoxes = list
    End Sub
    Public Sub IncreaseCounter()
        index = index + 1
    End Sub
    Public Sub DecreaseCounter()
        index = index - 1
    End Sub
    Public Sub BindCheckBox(ByVal sender As Object, ByVal e As EventArgs)
        'Create a new instance of a CheckBox. 
        Dim oCheckBox As CheckBox = CType(sender, CheckBox)
        Dim container As DataGridItem = CType(oCheckBox.NamingContainer, DataGridItem)

        Dim lab As New Label()

        For i = 0 To (maps.Count - 1)
            If String.Compare(maps(i).GetCheckBoxID(), oCheckBox.ID) = 0 Then
                maps(i).SetChoice(oCheckBox.Checked)
            End If
        Next

        controls.Add(lab)

    End Sub

    Public Sub SetMap(listPointer As List(Of Map))
        maps = listPointer
    End Sub

End Class

Public Class Map
    'This class maps a User to a CheckBox Object
    Public _user As SPUser = Nothing
    Public _group As SPGroup = Nothing
    Public _checkBox As CheckBox = Nothing
    Public Shared index = 0
    Public pos As Integer
    Public val As Boolean

    Public Sub Link(userName As SPUser, checkBox As CheckBox)

    End Sub
    Public Sub AddUser(user As SPUser)
        If (_user Is Nothing) Then
            _user = user
            pos = index
            index = index + 1
        End If
    End Sub
    Public Sub AddGroup(group As SPGroup)
        If (_group Is Nothing) Then
            _group = group
            pos = index
            index = index + 1
        End If
    End Sub
    Public Sub SetChoice(choice As Boolean)
        val = choice
    End Sub
    Public Sub AddCheckBox(checkBox As CheckBox)
        If (_checkBox Is Nothing) Then
            _checkBox = checkBox
        End If
    End Sub
    Public Sub ResetIndex()
        index = 0
    End Sub
    Public Function GetUsername()
        Return _user.Name
    End Function
    Public Function GetUserEmail()
        Return _user.Email
    End Function
    Public Function GetUser()
        Return _user
    End Function
    Public Function GetGroup()
        Return _group
    End Function
    Public Function GetGroupname()
        Return _group.Name
    End Function
    Public Function GetCheckBox()
        Return _checkBox
    End Function
    Public Function GetCheckBoxID()
        Return _checkBox.ID
    End Function
    Public Function GetChoice()
        Return val
    End Function
End Class
