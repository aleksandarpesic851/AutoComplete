Imports System.ComponentModel

Public Class ucComboboxAutoCompleteCell
    Inherits DataGridViewComboBoxCell

    Public mActualText As New List(Of String)
    Public mID As New List(Of String)
    Public mSynonymList As New List(Of String)

    Public Overrides ReadOnly Property EditType As Type
        Get
            Return GetType(ucComboboxAutoComplete)
        End Get
    End Property
    Public Overrides ReadOnly Property ValueType() As Type
        Get
            Return GetType(String)
        End Get
    End Property

    Private bAllowForItemsOnly As Boolean = True
    <DefaultValue(True)> <Browsable(False)>
    Public Property AllowForItemsOnly() As Boolean
        Get
            Return bAllowForItemsOnly
        End Get
        Set(value As Boolean)
            bAllowForItemsOnly = value
        End Set
    End Property

    Private bAllowUserToAdd As Boolean = True
    <DefaultValue(True)> <Browsable(False)>
    Public Property AllowUserToAddItem() As Boolean
        Get
            Return bAllowUserToAdd
        End Get
        Set(value As Boolean)
            bAllowUserToAdd = value
        End Set
    End Property
    Public Sub New()
        MyBase.New()
        Items.Add("")
        DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        FlatStyle = FlatStyle.Flat
    End Sub
    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer,
            ByVal initialFormattedValue As Object,
            ByVal dataGridViewCellStyle As DataGridViewCellStyle)
        ' Call base...
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)

        Dim ctl As ucComboboxAutoComplete = CType(DataGridView.EditingControl, ucComboboxAutoComplete)

        ' Make sure you have an instance...
        If ctl IsNot Nothing Then
            ctl.AllowForItemsOnly = AllowForItemsOnly
            ctl.AllowUserToAdd = AllowUserToAddItem

            ctl.DropDownStyle = ComboBoxStyle.DropDown
            ctl.AutoCompleteMode = AutoCompleteMode.None
            ctl.BeginUpdate()
            ctl.DropDownStyle = ComboBoxStyle.DropDown
            ctl.synonymMode_SetMode()
            For i As Integer = 0 To mActualText.Count - 1
                ctl.synonymMode_AddItem(mActualText(i), mID(i), mSynonymList(i))
            Next
            If Not Me.Value Is Nothing Then
                ctl.Text = initialFormattedValue
            End If
            ctl.EndUpdate()
            ctl.showDropDown()
        End If
    End Sub

    Public Sub SetData(ByRef sActualText As String, sID As String, sSynonymList As String)
        If mActualText.Contains(sActualText) Then
            Return
        End If

        Dim ucColumn As DataGridViewComboBoxColumn
        ucColumn = TryCast(OwningColumn, DataGridViewComboBoxColumn)
        mActualText.Add(sActualText)
        mID.Add(sID)
        mSynonymList.Add(sSynonymList)
        Me.Items.Add(sActualText)
        If ucColumn IsNot Nothing Then
            ucColumn.Items.Add(sActualText)
        End If
    End Sub

End Class