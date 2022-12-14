Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports ComboBox = System.Windows.Forms.ComboBox

<System.ComponentModel.ToolboxItem(True)>
Public Class ucComboboxAutoComplete
    Inherits ComboBox
    Implements IDataGridViewEditingControl

    Private Const CLICK_HERE = "<Click here to add value>"
    Private WithEvents m_dropDown As ucComboboxAutoComplete_DropDownControl
    Private WithEvents m_suggestionList As ListBox
    Private m_boldFont As Font
    Private m_fromKeyboard As Boolean
    Private m_matchingMethod As ucComboboxAutoComplete_StringMatchingMethod
    Private components As IContainer
    Private _bSelectAllow As Boolean = True
    Private _bAllowForItemsOnly As Boolean = True
    Private _bAllowUserToAdd As Boolean = True

    Public Overrides Property SelectedIndex As Integer
        Get
            Return Math.Max(m_suggestionList.SelectedIndex - DefaultItemCount(), -1)
        End Get
        Set(value As Integer)
            If (value < m_suggestionList.Items.Count - DefaultItemCount()) Then
                m_suggestionList.SelectedIndex = value + DefaultItemCount()
            End If
        End Set
    End Property

    <Browsable(True), EditorBrowsable(EditorBrowsableState.Always)>
    Public Event ItemFromSuggestionListSelected As EventHandler
    <Browsable(True), EditorBrowsable(EditorBrowsableState.Always)>
    Public Event NewItemAdded(ByVal value As String, ByVal tag As String, ByRef sender As ucComboboxAutoComplete)

    Public synonymMode As Boolean
    Public synonymMode_TextID As New List(Of String)
    Public synonymMode_DropDownDisplay As New List(Of String)
    Public synonymMode_ActualText As New List(Of String)

    Public Sub New()
        m_matchingMethod = ucComboboxAutoComplete_StringMatchingMethod.NoWildcards
        DropDownStyle = ComboBoxStyle.DropDown
        AutoCompleteMode = AutoCompleteMode.None

        m_suggestionList = New ListBox With {
                .DisplayMember = "Text",
                .TabStop = False,
                .Dock = DockStyle.Fill,
                .DrawMode = DrawMode.OwnerDrawFixed,
                .IntegralHeight = True,
                .Sorted = Sorted
            }

        AddHandler m_suggestionList.Click, AddressOf onSuggestionListClick
        AddHandler m_suggestionList.DrawItem, AddressOf onSuggestionListDrawItem
        AddHandler FontChanged, AddressOf onFontChanged2
        AddHandler m_suggestionList.MouseMove, AddressOf onSuggestionListMouseMove

        m_dropDown = New ucComboboxAutoComplete_DropDownControl(m_suggestionList)
        onFontChanged2(Nothing, Nothing)
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If m_boldFont IsNot Nothing Then
                m_boldFont.Dispose()
            End If

            m_dropDown.Dispose()

            If m_timerAutoFocus IsNot Nothing Then
                m_timerAutoFocus.Dispose()
                m_timerAutoFocus = Nothing
            End If

        End If

        MyBase.Dispose(disposing)
    End Sub

    Public Property SelectAllAutomatically As Boolean
        Get
            Return _bSelectAllow
        End Get
        Set(value As Boolean)
            _bSelectAllow = value
        End Set
    End Property

    <DefaultValue(True)>
    Public Property AllowForItemsOnly As Boolean
        Get
            Return _bAllowForItemsOnly
        End Get
        Set(value As Boolean)
            _bAllowForItemsOnly = value
        End Set
    End Property

    <DefaultValue(True)>
    Public Property AllowUserToAdd As Boolean
        Get
            Return _bAllowUserToAdd
        End Get
        Set(value As Boolean)
            _bAllowUserToAdd = value
        End Set
    End Property

    Protected Overrides Sub OnLocationChanged(ByVal e As EventArgs)
        MyBase.OnLocationChanged(e)
        hideDropDown()
    End Sub

    Protected Overrides Sub OnSizeChanged(ByVal e As EventArgs)
        MyBase.OnSizeChanged(e)
        m_dropDown.Width = Width
    End Sub

    Dim m_timerAutoFocus As Timer
    Public Sub showDropDown()
        If DesignMode Then
            Return
        End If
        UpdateDropdownList()

        'If MyBase.DroppedDown Then
        'BeginUpdate()
        'Dim oText As String = Text
        'Dim selStart As Integer = SelectionStart
        'Dim selLen As Integer = SelectionLength
        ''MyBase.DroppedDown = False
        'Text = showActualText(oText)
        'If _bSelectAllow Then [Select](selStart, selLen)
        'EndUpdate()
        'End If

        Dim h As Integer = Math.Min(MaxDropDownItems, m_suggestionList.Items.Count) * m_suggestionList.ItemHeight
        m_dropDown.Show(Me, New Size(DropDownWidth, h))

        If m_timerAutoFocus Is Nothing Then
            m_timerAutoFocus = New Timer()
            m_timerAutoFocus.Interval = 10
            AddHandler m_timerAutoFocus.Tick, AddressOf timerAutoFocus_Tick
        End If

        m_timerAutoFocus.Enabled = True
        m_sShowTime = DateTime.Now
    End Sub

    Private Sub timerAutoFocus_Tick(ByVal sender As Object, ByVal e As EventArgs)
        If m_dropDown.Visible Then
            m_timerAutoFocus.Enabled = False
        End If

        If MyBase.DroppedDown Then MyBase.DroppedDown = False
    End Sub


    Private m_lastHideTime As DateTime = DateTime.Now
    Private m_sShowTime As DateTime = DateTime.Now

    Public Sub hideDropDown()
        If m_dropDown.Visible Then
            m_lastHideTime = DateTime.Now
            m_dropDown.Close()
        End If
    End Sub

    Private Sub OnDropDownClosing(sender As Object, e As EventArgs) Handles m_dropDown.Closing
        ConfirmSuggestion(sender, e)
    End Sub

    Protected Overrides Sub OnLostFocus(ByVal e As EventArgs)
        If Not m_dropDown.Focused AndAlso Not m_suggestionList.Focused Then
            If m_suggestionList.Items.Count < 1 + DefaultItemCount() And AllowForItemsOnly Then
                Text = ""
            End If
            hideDropDown()
        End If
        MyBase.OnLostFocus(e)
    End Sub

    Protected Overrides Sub OnDropDown(ByVal e As EventArgs)
        hideDropDown()
        MyBase.OnDropDown(e)
    End Sub

    Private mAddValueDlg As Form

    Private Sub onSuggestionListClick(ByVal sender As Object, ByVal e As EventArgs)
        If m_suggestionList.SelectedIndex = 0 And AllowUserToAdd Then
            createAddValueDlg()
            Return
        End If
        Focus()
        If m_suggestionList.SelectedIndex > DefaultItemCount() - 1 Then
            RaiseEvent ItemFromSuggestionListSelected(Me, e)
        End If
    End Sub


    Private Sub createAddValueDlg()
        mAddValueDlg = New Form
        With mAddValueDlg
            .Text = "Add Value"
            .Size = New Size(290, 160)
        End With

        Dim ValueLabel As New Label
        With ValueLabel
            .Text = "Value:"
            .Size = New Size(40, 20)
            .Location = New Point(20, 20)
        End With
        mAddValueDlg.Controls.Add(ValueLabel)

        Dim ValueTextbox As New TextBox
        With ValueTextbox
            .Name = "dlgValue"
            .Text = Text
            .Size = New Size(80, 20)
            .Location = New Point(70, 20)
        End With
        mAddValueDlg.Controls.Add(ValueTextbox)

        Dim IdLabel As New Label
        With IdLabel
            .Text = "ID:"
            .Size = New Size(40, 20)
            .Location = New Point(20, 50)
        End With
        mAddValueDlg.Controls.Add(IdLabel)

        Dim IdTextbox As New TextBox
        With IdTextbox
            .Name = "dlgId"
            .Text = Text
            .Size = New Size(80, 20)
            .Location = New Point(70, 50)
        End With
        mAddValueDlg.Controls.Add(IdTextbox)


        Dim TagLabel As New Label
        With TagLabel
            .Text = "Tag:"
            .Size = New Size(40, 20)
            .Location = New Point(20, 80)
        End With
        mAddValueDlg.Controls.Add(TagLabel)

        Dim TagTextbox As New TextBox
        With TagTextbox
            .Name = "dlgTag"
            .Text = ""
            .Size = New Size(80, 20)
            .Location = New Point(70, 80)
        End With
        mAddValueDlg.Controls.Add(TagTextbox)

        Dim ConfrimButton As New Button
        With ConfrimButton
            .Text = "Add"
            .Size = New Size(50, 30)
            .Location = New Point(200, 50)
        End With
        mAddValueDlg.Controls.Add(ConfrimButton)
        AddHandler ConfrimButton.Click, AddressOf AddNewItem
        mAddValueDlg.ShowDialog(Me)
    End Sub
    Private Sub AddNewItem(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim val As String = "", id As String = "", tag As String = ""
        For Each ctrl As Control In mAddValueDlg.Controls
            Select Case ctrl.Name
                Case "dlgValue"
                    val = ctrl.Text
                Case "dlgId"
                    id = ctrl.Text
                Case "dlgTag"
                    tag = ctrl.Text
            End Select
        Next

        Dim dataGridView As DataGridView = CType(Me, IDataGridViewEditingControl).EditingControlDataGridView
        If dataGridView IsNot Nothing Then
            Dim idx As Integer = dataGridView.CurrentCell.ColumnIndex
            Dim column As DataGridViewComboBoxColumn = dataGridView.Columns(idx)
            column.Items.Add(val)
            dataGridView.CurrentCell.Value = val
        End If

        synonymMode_AddItem(val, id, "")
        Me.Text = val
        UpdateDropdownList()
        mAddValueDlg.Close()
        'Focus()
        RaiseEvent NewItemAdded(val, tag, Me)
    End Sub
    Private Sub ConfirmSuggestion(ByVal sender As Object, ByVal e As EventArgs)
        If m_suggestionList.SelectedIndex < 0 Then
            Return
        End If
        Dim sel As StringMatch = CType(m_suggestionList.Items(m_suggestionList.SelectedIndex), StringMatch)

        If AllowUserToAdd And sel.AppendOption Then
            'MessageBox.Show("Add new item")
        Else
            If m_suggestionList.Items.Count < 1 + DefaultItemCount() Then
                Return
            End If

            m_fromKeyboard = False

            Dim actualText As String = showActualText(sel.Text)
            If actualText.CompareTo(Text) <> 0 Then
                Text = actualText
                Dim dataGridView As DataGridView = CType(Me, IDataGridViewEditingControl).EditingControlDataGridView
                If dataGridView IsNot Nothing Then
                    dataGridView.CurrentCell.Value = actualText
                End If

                If IsNothing(sender) AndAlso IsNothing(e) Then
                    RaiseEvent ItemFromSuggestionListSelected(sender, e)
                End If
            End If
        End If
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If (keyData = Keys.Tab) AndAlso (m_dropDown.Visible) Then
            hideDropDown()
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
        m_fromKeyboard = True

        If Not m_dropDown.Visible Then
            MyBase.OnKeyDown(e)
            Return
        End If

        Select Case e.KeyCode
            Case Keys.Down

                If m_suggestionList.SelectedIndex < DefaultItemCount() Then
                    m_suggestionList.SelectedIndex = DefaultItemCount()
                ElseIf m_suggestionList.SelectedIndex < m_suggestionList.Items.Count - 1 Then
                    m_suggestionList.SelectedIndex += 1
                End If

            Case Keys.Up

                If m_suggestionList.SelectedIndex > DefaultItemCount() Then
                    m_suggestionList.SelectedIndex -= 1
                ElseIf m_suggestionList.SelectedIndex < DefaultItemCount() Then
                    m_suggestionList.SelectedIndex = m_suggestionList.Items.Count - 1
                End If

            Case Keys.Enter
                If m_suggestionList.Items.Count < 1 + DefaultItemCount() And AllowForItemsOnly Then
                    Text = ""
                End If
                hideDropDown()
            Case Keys.Escape
                hideDropDown()
            Case Else
                MyBase.OnKeyDown(e)
                Return
        End Select

        e.Handled = True
        e.SuppressKeyPress = True
    End Sub

    Protected Overrides Sub OnDropDownClosed(ByVal e As EventArgs)
        m_fromKeyboard = False
        MyBase.OnDropDownClosed(e)
    End Sub

    Private Const WM_COMMAND As UInteger = &H111
    Private Const WM_USER As UInteger = &H400
    Private Const WM_REFLECT As UInteger = WM_USER + &H1C00
    Private Const WM_LBUTTONDOWN As UInteger = &H201
    Private Const CBN_DROPDOWN As UInteger = 7
    Private Const CBN_CLOSEUP As UInteger = 8
    Private Function HIWORD(ByVal n As UInteger) As UInteger
        Return CUInt(n >> 16) And &HFFFF
    End Function

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_LBUTTONDOWN Then
            If m_dropDown.Visible Then
                hideDropDown()
            ElseIf (DateTime.Now - m_lastHideTime).Milliseconds > 1 Then
                showDropDown()
            End If
            Return
        End If
        If m.Msg = (WM_REFLECT + WM_COMMAND) Then

            Select Case HIWORD(CInt(m.WParam))
                Case CBN_DROPDOWN
                    If m_dropDown.Visible Then
                        hideDropDown()
                    ElseIf (DateTime.Now - m_lastHideTime).Milliseconds > 50 Then
                        showDropDown()
                    End If
                    Return
                Case CBN_CLOSEUP
                    If (DateTime.Now - m_sShowTime).Seconds > 1 Then hideDropDown()
                    Return
            End Select
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub UpdateDropdownList()
        m_suggestionList.BeginUpdate()
        m_suggestionList.Items.Clear()
        If AllowUserToAdd Then
            Dim appendSm As StringMatch = New StringMatch With {
                .AppendOption = True
            }
            m_suggestionList.Items.Add(appendSm)
        End If
        Dim matcher As ucComboboxAutoComplete_StringMatcher = New ucComboboxAutoComplete_StringMatcher(MatchingMethod, Text)

        If synonymMode Then
            For Each item As Object In synonymMode_DropDownDisplay
                Dim sm As StringMatch = matcher.Match(GetItemText(item))
                If sm IsNot Nothing Then
                    m_suggestionList.Items.Add(sm)
                End If
            Next
        Else
            For Each item As Object In Items
                Dim sm As StringMatch = matcher.Match(GetItemText(item))
                If sm IsNot Nothing Then
                    m_suggestionList.Items.Add(sm)
                End If
            Next
        End If

        If Text.Length > 0 Then
            If (m_suggestionList.SelectedIndex < 0 Or m_suggestionList.SelectedIndex > m_suggestionList.Items.Count - 1) And m_suggestionList.Items.Count > DefaultItemCount() Then
                m_suggestionList.SelectedIndex = DefaultItemCount()
            End If
        End If

        m_suggestionList.EndUpdate()
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As EventArgs)
        MyBase.OnTextChanged(e)

        If Not m_fromKeyboard OrElse Not Focused Then
            Return
        End If
        showDropDown()
        'Dim visible As Boolean = m_suggestionList.Items.Count <> 0

        'If m_suggestionList.Items.Count = 1 + DefaultItemCount() AndAlso (CType(m_suggestionList.Items(0), StringMatch)).Text.Length = Text.Trim().Length Then
        '    visible = False
        'End If

        'If visible Then
        '    showDropDown()
        'Else
        '    hideDropDown()
        'End If

        m_fromKeyboard = False
    End Sub

    Private Function DefaultItemCount() As Integer
        If AllowUserToAdd Then
            Return 1
        End If
        Return 0
    End Function

    Private Sub onSuggestionListMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        Dim idx As Integer = m_suggestionList.IndexFromPoint(e.Location)

        If (idx >= 0) AndAlso (idx <> m_suggestionList.SelectedIndex) Then
            m_suggestionList.SelectedIndex = idx
        End If
    End Sub

    Private Sub onFontChanged2(ByVal sender As Object, ByVal e As EventArgs)
        If m_boldFont IsNot Nothing Then
            m_boldFont.Dispose()
        End If

        m_suggestionList.Font = Font
        m_boldFont = New Font(Font, FontStyle.Bold)
        m_suggestionList.ItemHeight = m_boldFont.Height + 2
    End Sub

    Private Shared Sub DrawString(ByVal g As Graphics, ByVal color As Color, ByRef rect As Rectangle, ByVal text As String, ByVal font As Font)
        Dim proposedSize As Size = New Size(Integer.MaxValue, Integer.MaxValue)
        Dim sz As Size = TextRenderer.MeasureText(g, text, font, proposedSize, TextFormatFlags.NoPadding)
        TextRenderer.DrawText(g, text, font, rect, color, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding Or TextFormatFlags.NoPrefix)
        rect.X += sz.Width
        rect.Width -= sz.Width
    End Sub

    Private Sub onSuggestionListDrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        Dim sm As StringMatch = TryCast(m_suggestionList.Items(e.Index), StringMatch)
        If sm Is Nothing Then
            Return
        End If

        e.DrawBackground()
        Dim rBounds As Rectangle = e.Bounds
        Dim isBold As Boolean = sm.StartsOnMatch
        If (sm.AppendOption) Then
            DrawString(e.Graphics, e.ForeColor, rBounds, CLICK_HERE, Font)
        Else
            For Each s As String In sm.Segments
                Dim f As Font = If(isBold, m_boldFont, Font)
                DrawString(e.Graphics, e.ForeColor, rBounds, s, f)
                isBold = Not isBold
            Next
        End If
        e.DrawFocusRectangle()
    End Sub

    Private Sub ucListSelectedIndexChanged(sender As Object, e As EventArgs) Handles m_suggestionList.SelectedIndexChanged
        If SelectedIndex > DefaultItemCount() - 1 Then NotifyDataGridViewOfValueChange()
    End Sub

    <Category("Behavior"), DefaultValue(False), Description("Specifies whether items in the list portion of the combobo are sorted.")>
    Public Overloads Property Sorted As Boolean
        Get
            Return MyBase.Sorted
        End Get
        Set(ByVal value As Boolean)
            m_suggestionList.Sorted = value
            MyBase.Sorted = value
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overloads Property DroppedDown As Boolean
        Get
            Return MyBase.DroppedDown OrElse m_dropDown.Visible
        End Get
        Set(ByVal value As Boolean)
            m_dropDown.Visible = False
            MyBase.DroppedDown = value
        End Set
    End Property

    <DefaultValue(ucComboboxAutoComplete_StringMatchingMethod.NoWildcards), Description("How strings are matched against the user input"), Browsable(True), EditorBrowsable(EditorBrowsableState.Always), Category("Behavior")>
    Public Property MatchingMethod As ucComboboxAutoComplete_StringMatchingMethod
        Get
            Return m_matchingMethod
        End Get
        Set(ByVal value As ucComboboxAutoComplete_StringMatchingMethod)

            If m_matchingMethod <> value Then
                m_matchingMethod = value

                If m_dropDown.Visible Then
                    showDropDown()
                End If
            End If
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
    Public Overloads Property AutoCompleteSource As AutoCompleteSource
        Get
            Return MyBase.AutoCompleteSource
        End Get
        Set(ByVal value As AutoCompleteSource)
            MyBase.AutoCompleteSource = value
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
    Public Overloads Property AutoCompleteCustomSource As AutoCompleteStringCollection
        Get
            Return MyBase.AutoCompleteCustomSource
        End Get
        Set(ByVal value As AutoCompleteStringCollection)
            MyBase.AutoCompleteCustomSource = value
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
    Public Overloads Property AutoCompleteMode As AutoCompleteMode
        Get
            Return MyBase.AutoCompleteMode
        End Get
        Set(ByVal value As AutoCompleteMode)
            MyBase.AutoCompleteMode = value
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), EditorBrowsable(EditorBrowsableState.Never)>
    Public Overloads Property DropDownStyle As ComboBoxStyle
        Get
            Return MyBase.DropDownStyle
        End Get
        Set(ByVal value As ComboBoxStyle)
            MyBase.DropDownStyle = value
        End Set
    End Property
    Private dataGridViewControl As DataGridView
    Public Property EditingControlDataGridView As DataGridView Implements IDataGridViewEditingControl.EditingControlDataGridView
        Get
            Return dataGridViewControl
        End Get
        Set(ByVal value As DataGridView)
            dataGridViewControl = value
        End Set
    End Property

    Public Property EditingControlFormattedValue As Object Implements IDataGridViewEditingControl.EditingControlFormattedValue
        Get
            Return Me.Text
        End Get

        Set(ByVal value As Object)
            Try
                ' This will throw an exception of the string is 
                ' null, empty, or not in the format of a date.
                Me.Text = CStr(value)
            Catch
                ' In the case of an exception, just use the default
                ' value so we're not left with a null value.
                Me.Text = ""
            End Try
        End Set
    End Property

    Private rowIndexNum As Integer
    Public Property EditingControlRowIndex As Integer Implements IDataGridViewEditingControl.EditingControlRowIndex
        Get
            Return rowIndexNum
        End Get
        Set(ByVal value As Integer)
            rowIndexNum = value
        End Set
    End Property

    Private valueIsChanged As Boolean = False
    Public Property EditingControlValueChanged As Boolean Implements IDataGridViewEditingControl.EditingControlValueChanged
        Get
            Return valueIsChanged
        End Get
        Set(ByVal value As Boolean)
            valueIsChanged = value
        End Set
    End Property

    Public ReadOnly Property EditingPanelCursor As Cursor Implements IDataGridViewEditingControl.EditingPanelCursor
        Get
            Return MyBase.Cursor
        End Get
    End Property

    Public ReadOnly Property RepositionEditingControlOnValueChange As Boolean Implements IDataGridViewEditingControl.RepositionEditingControlOnValueChange
        Get
            Return False
        End Get
    End Property

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.SuspendLayout()
        '
        'ucComboboxAutoComplete
        '
        Me.ResumeLayout(False)

    End Sub

    Public Function synonymMode_getID() As String

        Dim sID As String = ""
        If synonymMode Then
            Dim sActualText As String = Me.Text
            Dim iINDEX As Integer = synonymMode_ActualText.IndexOf(sActualText)
            If iINDEX <> -1 Then
                sID = synonymMode_TextID(iINDEX)
            End If
        End If
        Return sID

    End Function

    Public Function synonymMode_getInternalData() As String

        Dim sID As String = ""
        Dim sActualText As String = Me.Text
        If synonymMode Then
            Dim iINDEX As Integer = synonymMode_ActualText.IndexOf(sActualText)
            If iINDEX <> -1 Then
                sID = synonymMode_TextID(iINDEX)
            End If
        End If

        Dim sInternalData As String = sID & "[|]" & sActualText

        Return sInternalData

    End Function
    Public Sub synonymMode_setID(ByRef sID As String)

        If sID <> "" Then
            Dim iINDEX As Integer = synonymMode_TextID.IndexOf(sID)
            If iINDEX <> -1 Then
                Me.SelectedIndex = iINDEX
            End If
        End If

    End Sub

    Public Sub synonymMode_setInternalData(ByRef sInternalData As String)

        If sInternalData <> "" Then
            Dim sSplit() As String = sInternalData.Split(New String() {"[|]"}, StringSplitOptions.None)
            Dim sID As String = sSplit(0)
            Dim sActualText As String = sSplit(1)
            If sID <> "" Then
                Dim iINDEX As Integer = synonymMode_TextID.IndexOf(sID)
                If iINDEX <> -1 Then
                    Me.SelectedIndex = iINDEX
                Else
                    Me.Text = sActualText
                End If
            End If
        End If

    End Sub
    Public Sub synonymMode_SetMode()
        synonymMode = True
        synonymMode_ActualText.Clear()
        synonymMode_DropDownDisplay.Clear()
        synonymMode_TextID.Clear()
        Me.Items.Clear()
    End Sub
    Public Sub synonymMode_AddItem(ByRef sActualText As String, Optional sID As String = "", Optional sSynonymList As String = "")

        If sActualText <> "" Then

            If synonymMode_ActualText.Contains(sActualText) Then
                MessageBox.Show("There are already same value")
                Exit Sub

            End If

            If sID <> "" And synonymMode_TextID.Contains(sID) Then
                MessageBox.Show("There are already same ID")
                Exit Sub
            End If

            Dim sDropDownDisplayText As String = ""
            If sSynonymList <> "" Then
                Dim lstSynonymList() As String = sSynonymList.Split(New String() {"[|]"}, StringSplitOptions.RemoveEmptyEntries)
                Dim ix As Integer
                Dim iEnd As Integer = lstSynonymList.Count - 1
                For ix = 0 To iEnd
                    sDropDownDisplayText = sDropDownDisplayText & lstSynonymList(ix)
                    If ix <> iEnd Then sDropDownDisplayText = sDropDownDisplayText & ", "
                Next
                If sDropDownDisplayText <> "" Then
                    sDropDownDisplayText = sActualText & " (" & sDropDownDisplayText & ")"
                End If
            Else
                sDropDownDisplayText = sActualText
            End If

            synonymMode_TextID.Add(sID)
            synonymMode_ActualText.Add(sActualText)
            synonymMode_DropDownDisplay.Add(sDropDownDisplayText)
            Me.Items.Add(sActualText)
        End If

    End Sub

    Public Function showActualText(ByRef sInputString As String) As String

        Dim sActualText As String = sInputString
        'Dim nItemCount = Me.Items.Count - DefaultItemCount()

        If synonymMode Then
            'If nItemCount = synonymMode_DropDownDisplay.Count And nItemCount = synonymMode_ActualText.Count Then
            Dim iINDEX As Integer = synonymMode_DropDownDisplay.IndexOf(sInputString)
            If iINDEX <> -1 Then
                sActualText = synonymMode_ActualText(iINDEX)
            End If
            'End If
        End If

        Return sActualText

    End Function

    Public Sub ApplyCellStyleToEditingControl(dataGridViewCellStyle As DataGridViewCellStyle) Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl
        Me.Font = dataGridViewCellStyle.Font
        Me.ForeColor = dataGridViewCellStyle.ForeColor
        Me.BackColor = dataGridViewCellStyle.BackColor
    End Sub

    Public Function EditingControlWantsInputKey(keyData As Keys, dataGridViewWantsInputKey As Boolean) As Boolean Implements IDataGridViewEditingControl.EditingControlWantsInputKey
        Return True
    End Function

    Public Function GetEditingControlFormattedValue(context As DataGridViewDataErrorContexts) As Object Implements IDataGridViewEditingControl.GetEditingControlFormattedValue
        Return Text
    End Function

    Public Sub PrepareEditingControlForEdit(selectAll As Boolean) Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

    End Sub


    Private Sub NotifyDataGridViewOfValueChange()
        CType(Me, IDataGridViewEditingControl).EditingControlValueChanged = True
        Dim dataGridView As DataGridView = CType(Me, IDataGridViewEditingControl).EditingControlDataGridView
        If dataGridView IsNot Nothing Then
            dataGridView.NotifyCurrentCellDirty(True)
        End If
    End Sub

End Class
