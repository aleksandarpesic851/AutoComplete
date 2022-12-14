Imports System.ComponentModel
Imports System.Drawing.Design

Public Class ucComboboxAutoCompleteColumn
    Inherits DataGridViewComboBoxColumn

    <DefaultValue(True)>
    Public Property AllowForItemsOnly As Boolean

    <DefaultValue(True)>
    Public Property AllowUserToAdd As Boolean
    Public Overrides Property CellTemplate As DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As DataGridViewCell)
            ' Ensure that the cell used for the template is a MoneyTextBoxCell.
            If value IsNot Nothing AndAlso Not value.[GetType]().IsAssignableFrom(GetType(ucComboboxAutoCompleteCell)) Then
                Throw New InvalidCastException("Must be a ucComboboxAutoCompleteCell")
            End If
            MyBase.CellTemplate = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        DisplayMember = String.Empty
        ValueMember = String.Empty
        DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        MyBase.CellTemplate = New ucComboboxAutoCompleteCell()
    End Sub

End Class