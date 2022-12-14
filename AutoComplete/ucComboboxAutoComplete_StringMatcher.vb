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
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' Possible ways of matching methods
''' </summary>
Public Enum ucComboboxAutoComplete_StringMatchingMethod
    NoWildcards
    UseWildcards
    UseRegexs
End Enum


''' <summary>
''' Allows to decompose a strings against a pattern
''' </summary>
Public Class ucComboboxAutoComplete_StringMatcher

    ''' <summary>
    ''' Match against a string
    ''' </summary>
    Private _Match As MatchDelegate

    Public Delegate Function MatchDelegate(ByVal source As String) As StringMatch
    Private ReadOnly m_pattern As Object

    Public Property Match As MatchDelegate
        Get
            Return _Match
        End Get
        Private Set(ByVal value As MatchDelegate)
            _Match = value
        End Set
    End Property


    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New(ByVal method As ucComboboxAutoComplete_StringMatchingMethod, ByVal pattern As String)
        Select Case method
            Case ucComboboxAutoComplete_StringMatchingMethod.NoWildcards
                Me.m_pattern = ucComboboxAutoComplete_StringMatcher.prepareNoWildcard(pattern)
                Me.Match = AddressOf Me.buildNoWildcard
            Case ucComboboxAutoComplete_StringMatchingMethod.UseWildcards
                Me.m_pattern = ucComboboxAutoComplete_StringMatcher.prepareWithWildcards(pattern)
                Me.Match = AddressOf Me.buildWithWildcards
            Case ucComboboxAutoComplete_StringMatchingMethod.UseRegexs
                Me.m_pattern = ucComboboxAutoComplete_StringMatcher.prepareRegex(pattern)
                Me.Match = AddressOf Me.buildRegex
            Case Else
                Throw New System.ArgumentException("Unknown StringMatchingMethod value")
        End Select
    End Sub


#Region "No wildcards"
    ''' <summary>
    ''' Prepare the pattern for a "non wildcard" match
    ''' </summary>
    Private Shared Function prepareNoWildcard(ByVal pattern As String) As Object
        Return pattern.ToLower()
    End Function


    ''' <summary>
    ''' Compare the source against the pattern and if successfull returns a StringMatch
    ''' 
    ''' There is a match if source contains all the characters of pattern in the right order
    ''' but not consecutively
    ''' </summary>
    Private Function buildNoWildcard(ByVal source As String) As StringMatch
        Dim pattern As String = CStr(Me.m_pattern)

        Dim lsource As String = source.ToLower()
        Dim segments As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)()
        Dim j As Integer, i As Integer = 0
        Dim NP As Integer = pattern.Length, NS As Integer = source.Length
        Dim startsOnMatch As Boolean = False, isMatch As Boolean = False

        If pattern.Length < 1 Then
            segments.Add(source)
            Return New StringMatch With {
                .Text = source,
                .Segments = segments,
                .StartsOnMatch = startsOnMatch
            }
        End If


        For j = 0 To NP - 1
            Dim s As Integer = i

            ' skip characters until we have a match
            While i < NS

                If lsource(i) = pattern(j) Then
                    isMatch = (j = NP - 1)

                    If s <> i Then
                        segments.Add(source.Substring(s, i - s))
                        segments.Add(source.Substring(i, 1))
                    Else

                        If segments.Count = 0 Then
                            segments.Add(source.Substring(i, 1))
                            startsOnMatch = True
                        Else
                            segments(segments.Count - 1) += source(i)
                        End If
                    End If

                    i += 1
                    Exit While ' i loop
                End If

                i += 1
            End While

            If i >= NS Then
                Exit For ' j loop
            End If
        Next

        If Not isMatch Then
            Return Nothing
        End If

        If i < NS Then
            segments.Add(source.Substring(i, NS - i))
        End If

        Return New StringMatch With {
                .Text = source,
                .Segments = segments,
                .StartsOnMatch = startsOnMatch
            }
    End Function

#End Region

#Region "Wildcards"
    ''' <summary>
    ''' Prepare the pattern for a "wildcard" match
    ''' </summary>
    Private Shared Function prepareWithWildcards(ByVal pattern As String) As Object
        ' starts the match anywhere
        Dim buf As System.Text.StringBuilder = New System.Text.StringBuilder(If(pattern.StartsWith("*"), "", "*"))

        ' reduce consecutive '*'s into a single one
        Dim NP As Integer = pattern.Length

        For i As Integer = 0 To NP - 1
            Dim c As Char = pattern(i)

            If c = "*"c Then

                ' "****" == "*"
                While (i < NP) AndAlso (pattern(i) = "*"c)
                    i += 1
                End While

                buf.Append("*"c)

                ' don't need to process a '*' at the end of the pattern
                If i < NP Then
                    buf.Append(pattern(i))
                End If
            Else
                buf.Append(c)
            End If
        Next

        Return buf.ToString().ToLower()
    End Function


    ''' <summary>
    ''' Add a chunk to list of segments
    ''' </summary>
    Private Shared Sub addChunkToStringMatch(ByVal res As StringMatch, ByVal txt As String, ByVal isMatch As Boolean)
        If res.Segments.Count = 0 Then
            res.StartsOnMatch = isMatch
            res.Segments.Add(txt)
            Return
        End If

        Dim currentIsMatch As Boolean = (((res.Segments.Count Mod 2) <> 0) = res.StartsOnMatch)

        If currentIsMatch = isMatch Then
            ' we can add to the previous segment
            res.Segments(res.Segments.Count - 1) = res.Segments(res.Segments.Count - 1) & txt
        Else
            res.Segments.Add(txt)
        End If
    End Sub


    ''' <summary>
    ''' Compare the source against the pattern and if successfull returns a StringMatch
    ''' 
    ''' There match is done against the source which can contain the wildcard '*'
    ''' </summary>
    Private Function buildWithWildcards(ByVal source As String) As StringMatch
        Dim res As StringMatch = New StringMatch With {
                .Text = source,
                .Segments = New System.Collections.Generic.List(Of String)(),
                .StartsOnMatch = False
            }
        Dim pattern As String = CStr(Me.m_pattern)
        Dim lsource As String = source.ToLower()
        Dim iP As Integer = 0, [iS] As Integer = 0, icP As Integer = -1, icS As Integer = -1
        Dim ioS1 As Integer = -1, ioS2 As Integer = -1
        Dim NP As Integer = pattern.Length, NS As Integer = source.Length

        If pattern.Length < 1 Then
            res.Segments.Add(source)
            Return res
        End If


        ' skip all the starting characters untils we either don't match or reach a '*'
        While (iP <> NP) AndAlso ([iS] <> NS) AndAlso (pattern(iP) <> "*"c)

            If lsource([iS]) <> pattern(iP) Then
                ' mismatch
                Return Nothing
            End If

            [iS] += 1
            iP += 1
        End While

        If [iS] <> 0 Then
            ucComboboxAutoComplete_StringMatcher.addChunkToStringMatch(res, source.Substring(0, [iS]), True)
        End If

        If iP <> NP Then

            ' we have not finished parsing

            ' now we start from a '*'
            ' we try to get a sequence of matches to ('*', non '*'s), ('*', non '*'s), ....
            While [iS] <> NS

                If pattern(iP) = "*"c Then
                    ' we have reached a new '*', we will now compare against the next pattern character
                    iP += 1

                    If iP = NP Then
                        'we have matched all the pattern
                        Exit While ' while
                    End If

                    If icP <> iP Then

                        ' this is the 1st time we try to process this '*'
                        If icP >= 0 Then

                            ' not the first run
                            ' so we have successfully processed the previous '*'
                            If ioS1 <> ioS2 Then
                                ucComboboxAutoComplete_StringMatcher.addChunkToStringMatch(res, source.Substring(ioS1, ioS2 - ioS1), False)
                            End If

                            If ioS2 <> [iS] Then
                                ucComboboxAutoComplete_StringMatcher.addChunkToStringMatch(res, source.Substring(ioS2, [iS] - ioS2), True)
                            End If
                        End If

                        ioS1 = [iS]
                        ioS2 = -1
                    End If

                    ' we save the current position in case the '*' needs to match more characters
                    icP = iP
                    icS = [iS] + 1
                ElseIf lsource([iS]) = pattern(iP) Then

                    ' if we were processing a '*', we now have a successfull match
                    ' we are still matching sucessfully the pattern
                    If ioS2 = -1 Then
                        ioS2 = [iS]
                    End If

                    [iS] += 1
                    iP += 1
                Else
                    ' we have a mismatch
                    ' let's try again with '*' matching one more character
                    iP = icP
                    [iS] = System.Math.Min(System.Threading.Interlocked.Increment(icS), icS - 1)
                End If

                If iP = NP Then
                    'we have matched all the pattern
                    Exit While ' while
                End If
            End While

            ' we have processed either the whole source or either the whole pattern
            ' let's process any remaining '*' which here matches the empty string
            While (iP <> NP) AndAlso (pattern(iP) = "*"c)
                iP += 1
            End While
        End If

        If iP = NP Then

            ' we have a match!
            If icP <> iP Then

                If ioS1 <> -1 Then

                    If ioS2 = -1 Then
                        ucComboboxAutoComplete_StringMatcher.addChunkToStringMatch(res, source.Substring(ioS1, [iS] - ioS1), False)
                    Else

                        If ioS2 <> ioS1 Then
                            ucComboboxAutoComplete_StringMatcher.addChunkToStringMatch(res, source.Substring(ioS1, ioS2 - ioS1), False)
                        End If

                        If ioS2 <> [iS] Then
                            ucComboboxAutoComplete_StringMatcher.addChunkToStringMatch(res, source.Substring(ioS2, [iS] - ioS2), True)
                        End If
                    End If
                End If
            End If

            If [iS] <> NS Then
                ucComboboxAutoComplete_StringMatcher.addChunkToStringMatch(res, source.Substring([iS], NS - [iS]), False)
            End If

            Return res
        End If

        Return Nothing
    End Function

#End Region

#Region "Regexs"
    ''' <summary>
    ''' Prepare the pattern for a "regex" match
    ''' </summary>
    Private Shared Function prepareRegex(ByVal pattern As String) As Object
        Try
            Return New System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Catch
            ' the user has entered an invalid regex, just ignore it
        End Try

        Return Nothing
    End Function


    ''' <summary>
    ''' Compare the source against the pattern and if successfull returns a StringMatch
    ''' 
    ''' There is a match if source contains all the characters of pattern in the right order
    ''' but not consecutively
    ''' </summary>
    Private Function buildRegex(ByVal source As String) As StringMatch
        If Me.m_pattern Is Nothing Then
            Return Nothing
        End If

        Dim re As System.Text.RegularExpressions.Regex = CType(Me.m_pattern, System.Text.RegularExpressions.Regex)
        Dim segments As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)()
        Dim startsOnMatch As Boolean = False, isMatch As Boolean = False
        Dim idx As Integer = 0

        Dim pattern As String = CStr(Me.m_pattern)
        If pattern.Length < 1 Then
            segments.Add(source)
            Return New StringMatch With {
                .Text = source,
                .Segments = segments,
                .StartsOnMatch = startsOnMatch
            }
        End If

        ' we need only one match
        Dim m As System.Text.RegularExpressions.Match = re.Match(source, idx)

        If m.Success Then
            isMatch = True

            If m.Index = idx Then

                If m.Index = 0 Then
                    startsOnMatch = True
                End If

                segments.Add(source.Substring(m.Index, m.Length))
            Else
                segments.Add(source.Substring(idx, m.Index - idx))
                segments.Add(source.Substring(m.Index, m.Length))
            End If

            idx = m.Index + m.Length
        End If

        If Not isMatch Then
            Return Nothing
        End If

        If idx < source.Length Then
            segments.Add(source.Substring(idx, source.Length - idx))
        End If

        Return New StringMatch With {
            .Text = source,
            .Segments = segments,
            .StartsOnMatch = startsOnMatch
        }
    End Function
#End Region
End Class


''' <summary>
''' The result of a match
''' There are the items we store in the suggestion listbox
''' </summary>
Public Class StringMatch
    ''' <summary>
    ''' The original source
    ''' </summary>
    Private _Text As String, _Segments As List(Of String), _StartsOnMatch As Boolean
    Private _AppendOption As Boolean = False

    Public Property AppendOption As Boolean
        Get
            Return _AppendOption
        End Get
        Friend Set(ByVal value As Boolean)
            _AppendOption = value
        End Set
    End Property

    Public Property Text As String
        Get
            If AppendOption Then
                Return ""
            Else
                Return _Text
            End If
        End Get
        Friend Set(ByVal value As String)
            _Text = value
        End Set
    End Property

    Public Property Segments As List(Of String)
        Get
            Return _Segments
        End Get
        Friend Set(ByVal value As List(Of String))
            _Segments = value
        End Set
    End Property

    Public Property StartsOnMatch As Boolean
        Get
            Return _StartsOnMatch
        End Get
        Friend Set(ByVal value As Boolean)
            _StartsOnMatch = value
        End Set
    End Property
End Class
