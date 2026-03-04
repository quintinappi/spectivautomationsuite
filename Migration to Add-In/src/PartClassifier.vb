' ============================================================================
' INVENTOR AUTOMATION SUITE - PART CLASSIFIER
' ============================================================================
' Description: Classify parts based on description for heritage naming
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
'
' Ported from: Assembly_Cloner.vbs - ClassifyByDescription()
'
' CLASSIFICATION LOGIC:
' - I and H sections (UB, UC) -> B
' - Platework (PL + S355JR) -> PL
' - Liners (PL + NOT S355JR) -> LPL
' - Angles (L with X) -> A
' - Channels (PFC, TFC) -> CH
' - Circular hollow (CHS) -> P
' - Square hollow (SHS) -> SQ
' - Flatbar (FL, not FLOOR) -> FL
' - IPE beams -> IPE
' - Flanges (contains FLANGE) -> FLG
' - Pipe (contains PIPE) -> P
' - Roundbar (R + digit) -> R
' - Hardware (BOLT, SCREW, WASHER, NUT) -> SKIP
' ============================================================================
'
' MIGRATION STATUS: ✅ COMPLETE
' ============================================================================

Imports System

Namespace SpectivInventorSuite

    ''' <summary>
    ''' Part classification based on description for heritage naming
    ''' </summary>
    Public Class PartClassifier

        ''' <summary>
        ''' Classify part based on description using client's exact requirements
        ''' VBScript: ClassifyByDescription()
        ''' </summary>
        ''' <param name="description">Part description from iProperties</param>
        ''' <returns>Group code (PL, B, CH, A, FL, etc.) or "SKIP" for hardware</returns>
        Public Shared Function ClassifyByDescription(description As String) As String
            Try
                If String.IsNullOrWhiteSpace(description) Then
                    Return "OTHER"
                End If

                Dim desc As String = description.Trim().ToUpper()

                ' Skip hardware and bolts first
                If desc.Contains("BOLT") OrElse desc.Contains("SCREW") OrElse
                   desc.Contains("WASHER") OrElse desc.Contains("NUT") Then
                    Return "SKIP"
                End If

                ' Check for FLANGE in description
                If desc.Contains("FLANGE") Then
                    Return "FLG"  ' Flanges
                End If

                ' Check for PIPE
                If desc.Contains("PIPE") Then
                    Return "P"  ' Pipes
                End If

                ' Check for Roundbar R followed by digits
                If desc.Length >= 2 AndAlso desc(0) = "R"c AndAlso Char.IsDigit(desc(1)) Then
                    Return "R"  ' Roundbar
                End If

                ' Client's grouping logic - exact requirements
                If desc.StartsWith("UB") Then
                    Return "B"  ' I and H sections - UB beams
                ElseIf desc.StartsWith("UC") Then
                    Return "B"  ' I and H sections - UC columns
                ElseIf desc.StartsWith("PL") Then
                    ' Check if it's platework (PL + S355JR) or liners (PL + NOT S355JR)
                    If desc.Contains("S355JR") Then
                        Return "PL"  ' Platework
                    Else
                        Return "LPL" ' Liners
                    End If
                ElseIf desc.StartsWith("L") AndAlso (desc.Contains("X") OrElse desc.Contains(" X ")) Then
                    Return "A"   ' Angles - L50x50x6, L70x70x6 etc.
                ElseIf desc.StartsWith("PFC") Then
                    Return "CH"  ' Parallel flange channels
                ElseIf desc.StartsWith("TFC") Then
                    Return "CH"  ' Taper flange channels
                ElseIf desc.StartsWith("CHS") Then
                    Return "P"   ' Circular hollow sections
                ElseIf desc.StartsWith("SHS") Then
                    Return "SQ"  ' Square/rectangular hollow sections
                ElseIf desc.StartsWith("FL") AndAlso Not desc.Contains("FLOOR") Then
                    Return "FL"  ' Flatbar (but not floor grating)
                ElseIf desc.StartsWith("IPE") Then
                    Return "IPE"  ' European I-beams (separate group)
                Else
                    ' Default - unclassified part
                    Return "OTHER"
                End If

            Catch ex As Exception
                ' Return OTHER on any error
                Return "OTHER"
            End Try
        End Function

        ''' <summary>
        ''' Get all supported group codes
        ''' </summary>
        Public Shared Function GetSupportedGroups() As String()
            Return {"PL", "B", "CH", "A", "FL", "LPL", "SQ", "P", "R", "FLG", "IPE", "OTHER"}
        End Function

        ''' <summary>
        ''' Get group description
        ''' </summary>
        Public Shared Function GetGroupDescription(groupCode As String) As String
            Select Case groupCode
                Case "PL"
                    Return "Platework (PL + S355JR)"
                Case "B"
                    Return "I and H sections (UB, UC)"
                Case "CH"
                    Return "Channels (PFC, TFC)"
                Case "A"
                    Return "Angles (L with X)"
                Case "FL"
                    Return "Flatbar"
                Case "LPL"
                    Return "Liners (PL + NOT S355JR)"
                Case "SQ"
                    Return "Square/Rectangular hollow (SHS)"
                Case "P"
                    Return "Circular hollow (CHS) / Pipes"
                Case "R"
                    Return "Roundbar"
                Case "FLG"
                    Return "Flanges"
                Case "IPE"
                    Return "European I-beams"
                Case "SKIP"
                    Return "Hardware (BOLT, SCREW, WASHER, NUT)"
                Case "OTHER"
                    Return "Unclassified"
                Case Else
                    Return "Unknown group"
            End Select
        End Function

    End Class

End Namespace
