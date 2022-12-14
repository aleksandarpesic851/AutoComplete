' Copyright © Serge Weinstock 2014.
'
' This library is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public License
' along with this library.  If not, see <http://www.gnu.org/licenses/>.

Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms


''' <summary>
''' A dropdown window for combos.
''' </summary>
<ToolboxItem(False)>
Public Class ucComboboxAutoComplete_DropDownControl
    Inherits ToolStripDropDown
    Implements IMessageFilter

#Region "Properties"
    ''' <summary>
    ''' Gets the content of the pop-up.
    ''' </summary>
    Private _Content As Control

    Public Property Content As Control
        Get
            Return _Content
        End Get
        Private Set(ByVal value As Control)
            _Content = value
        End Set
    End Property

#End Region

#Region "Fields"
    ''' <summary>The toolstrip host</summary>
    Private m_host As ToolStripControlHost
    ''' <summary>The control for which we display this dropdown</summary>
    Private m_opener As Control

#End Region

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New(ByVal content As Control)
        If content Is Nothing Then
            Throw New ArgumentNullException("content")
        End If

        Me.Content = content
        Me.Content.Dock = DockStyle.None
        Me.Content.Location = Point.Empty

        m_host = New ToolStripControlHost(Me.Content)
        ' NB: AutoClose must be set to false, because otherwise the ToolStripManager would steal keyboard events
        AutoClose = False
        ' we do ourselves the sizing
        AutoSize = False
        MyBase.DoubleBuffered = True
        ResizeRedraw = False
        ' we adjust the size according to the contents
        MyBase.MinimumSize = Me.Content.MinimumSize
        content.MinimumSize = Me.Content.Size
        MyBase.MaximumSize = Me.Content.MaximumSize
        content.MaximumSize = Me.Content.Size
        Size = Me.Content.Size
        ' set up the content
        MyBase.Items.Add(m_host)
        ' we must listen to mouse events for "emulating" AutoClose
        Application.AddMessageFilter(Me)
        Me.Padding = New Padding(0)
        m_host.Padding = New Padding(0)
        Me.Margin = New Padding(0)
        m_host.Margin = New Padding(0)
    End Sub


    ''' <summary>
    ''' Display the dropdown and adjust its size and location
    ''' </summary>
    Public Overloads Sub Show(ByVal opener As Control, ByVal Optional preferredSize As Size = Nothing)
        If opener Is Nothing Then
            Throw New ArgumentNullException("opener")
        End If

        If preferredSize = Nothing Then
            preferredSize = New Size(10, 10)
        End If

        m_opener = opener
        Dim w As Integer = If(preferredSize.Width = 0, ClientRectangle.Width, preferredSize.Width)
        Dim h As Integer = If(preferredSize.Height = 0, Content.Height, preferredSize.Height)
        h += Padding.Size.Height + Content.Margin.Size.Height
        Dim screen As Rectangle = Windows.Forms.Screen.FromControl(m_opener).WorkingArea

        ' let's try first to place it below the opener control
        Dim loc As Rectangle = m_opener.RectangleToScreen(New Rectangle(m_opener.ClientRectangle.Left, m_opener.ClientRectangle.Bottom, m_opener.ClientRectangle.Left + w, m_opener.ClientRectangle.Bottom + h))
        Dim cloc As Point = New Point(m_opener.ClientRectangle.Left, m_opener.ClientRectangle.Bottom)

        If Not screen.Contains(loc) Then
            ' let's try above the opener control
            loc = m_opener.RectangleToScreen(New Rectangle(m_opener.ClientRectangle.Left, m_opener.ClientRectangle.Top - h, m_opener.ClientRectangle.Left + w, m_opener.ClientRectangle.Top))

            If screen.Contains(loc) Then
                cloc = New Point(m_opener.ClientRectangle.Left, m_opener.ClientRectangle.Top - h)
            End If
        End If

        Me.Content.Location = Point.Empty
        Width = w
        Height = h
        MyBase.Show(m_opener, cloc, ToolStripDropDownDirection.BelowRight)
    End Sub


#Region "internals"
    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If Content IsNot Nothing Then
                Content.Dispose()
                Content = Nothing
            End If

            Application.RemoveMessageFilter(Me)
        End If

        MyBase.Dispose(disposing)
    End Sub


    ''' <summary>
    ''' On resizes, resize the contents
    ''' </summary>
    Protected Overrides Sub OnSizeChanged(ByVal e As EventArgs)
        If Content IsNot Nothing Then
            Content.MinimumSize = Size
            Content.MaximumSize = Size
            Content.Size = Size
            Content.Location = Point.Empty
        End If

        MyBase.OnSizeChanged(e)
    End Sub

#End Region

#Region "IMessageFilter implementation"
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const WM_NCRBUTTONDOWN As Integer = &HA4
    Private Const WM_NCMBUTTONDOWN As Integer = &HA7

    <DllImport("user32.dll", ExactSpelling:=True, CharSet:=CharSet.Auto)>
    Public Shared Function MapWindowPoints(ByVal hWndFrom As IntPtr, ByVal hWndTo As IntPtr,
        <[In], Out> ByRef pt As Point, ByVal cPoints As Integer) As Integer
    End Function


    ''' <summary>
    ''' We monitor all messages in order to detect when the users clicks outside the dropdown and the combo
    ''' If this happens, we close the dropdown (as AutoClose is false)
    ''' </summary>
    Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
        If Visible Then

            Select Case m.Msg
                Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN
                    Dim i As Long = m.LParam
                    Dim x As Short = CShort(i And &HFFFF)
                    Dim y As Short = CShort(i >> 16 And &HFFFF)
                    Dim pt As Point = New Point(x, y)
                    ' client area: x, y are relative to the client area of the windows
                    ' non-client area: x, y are relative to the desktop
                    Dim srcWnd As IntPtr = If(m.Msg = WM_LBUTTONDOWN OrElse m.Msg = WM_RBUTTONDOWN OrElse m.Msg = WM_MBUTTONDOWN, m.HWnd, IntPtr.Zero)
                    MapWindowPoints(srcWnd, Handle, pt, 1)

                    If Not ClientRectangle.Contains(pt) Then
                        ' the user has clicked outside the dropdown
                        pt = New Point(x, y)
                        MapWindowPoints(srcWnd, m_opener.Handle, pt, 1)

                        If Not m_opener.ClientRectangle.Contains(pt) Then
                            ' the user has clicked outside the opener control
                            Close()
                        End If
                    End If
            End Select
        End If

        Return False
    End Function

    Private Class CSharpImpl
        <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
        Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class
#End Region
End Class
