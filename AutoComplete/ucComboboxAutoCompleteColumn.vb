﻿Imports System.ComponentModel
Imports System.Drawing.Design

Public Class ucComboboxAutoCompleteColumn
    Inherits DataGridViewComboBoxColumn
    Public Overrides Property CellTemplate As DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As DataGridViewCell)
            If value IsNot Nothing AndAlso Not value.[GetType]().IsAssignableFrom(GetType(ucComboboxAutoCompleteCell)) Then
                Throw New InvalidCastException("Must be a ucComboboxAutoCompleteCell")
            End If
            MyBase.CellTemplate = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        MyBase.CellTemplate = New ucComboboxAutoCompleteCell()
    End Sub

End Class