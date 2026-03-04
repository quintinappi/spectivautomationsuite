' PlateDocumentSettings.vb
' Applies Document Settings to plate parts to eliminate decimal precision
' Manual trigger - no background service
' Author: Quintin de Bruin © 2026

Imports Inventor
Imports System
Imports System.Windows.Forms

Public Class PlateDocumentSettings
    Private m_InventorApp As Inventor.Application

    ' Constants for Document Settings
    Private Const DISPLAY_AS_VALUE As Integer = 34821

    Public Sub New(ByVal inventorApp As Inventor.Application)
        m_InventorApp = inventorApp
    End Sub

    Public Sub ApplyDecimalFix(ByVal partDoc As PartDocument)
        Try
            Dim params As Parameters = partDoc.ComponentDefinition.Parameters

            ' Setting 1: Linear Dimension Precision = 0 decimals
            params.LinearDimensionPrecision = 0

            ' Setting 2: Modeling Dimension Display = "Display as value" (34821)
            params.DimensionDisplayType = CType(DISPLAY_AS_VALUE, DimensionDisplayTypeEnum)

            ' Setting 3: Default Parameter Input Display = "Display as expression" (True)
            params.DisplayParameterAsExpression = True

            ' Save the document to apply changes
            partDoc.Update()

            System.Diagnostics.Debug.WriteLine("Decimal fix applied to: " & partDoc.DisplayName)

        Catch ex As Exception
            Throw New Exception("Failed to apply decimal fix: " & ex.Message)
        End Try
    End Sub

End Class
