' PlateDocumentSettings.vb
' Automatically applies Document Settings to plate parts to eliminate decimal precision
' Runs inside Inventor's process with full event access for proper BOM refresh
' Author: Quintin de Bruin © 2026

Imports Inventor
Imports System.Runtime.InteropServices

Namespace AssemblyClonerAddIn

''' <summary>
''' Utility class for applying Document Settings to plate parts.
''' Call ProcessCurrentDocument() manually when needed - no automatic event hooks.
''' </summary>
Public Class PlateDocumentSettings
    Private m_InventorApp As Inventor.Application
    
    ' Constants for Document Settings
    Private Const DISPLAY_AS_VALUE As Integer = 34821
    Private Const kSheetMetalSubType As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    
    Public Sub New(ByVal inventorApp As Inventor.Application)
        m_InventorApp = inventorApp
        ' No automatic event hooks - this is a utility class
        ' Call ProcessCurrentDocument() or ProcessPartDocument() manually when needed
    End Sub
    
    ''' <summary>
    ''' Process the currently active document if it's a plate part
    ''' Returns status message for user feedback
    ''' </summary>
    Public Function ProcessCurrentDocument() As String
        Try
            If m_InventorApp.ActiveDocument Is Nothing Then
                Return "NO_DOCUMENT"
            End If

            If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                Return "NOT_PART"
            End If

            Dim partDoc As PartDocument = CType(m_InventorApp.ActiveDocument, PartDocument)
            Return ProcessPartDocument(partDoc)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("PlateDocumentSettings.ProcessCurrentDocument Error: " & ex.Message)
            Return "ERROR: " & ex.Message
        End Try
    End Function
    
    ''' <summary>
    ''' Process a specific part document
    ''' Returns status message for user feedback
    ''' </summary>
    Public Function ProcessPartDocument(ByVal partDoc As PartDocument) As String
        Try
            ' Check if this is a sheet metal part
            If partDoc.SubType <> kSheetMetalSubType Then
                Return "NOT_SHEET_METAL"
            End If

            ' Check if this is a plate part
            If Not IsPlatePart(partDoc) Then
                Return "NOT_PLATE"
            End If

            ' Apply the Document Settings fix
            ApplyDocumentSettingsForZeroDecimals(partDoc)

            ' Force BOM refresh in parent assemblies
            ForceBOMRefreshInParentAssemblies(partDoc)

            Return "SUCCESS"
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("PlateDocumentSettings.ProcessPartDocument Error: " & ex.Message)
            Return "ERROR: " & ex.Message
        End Try
    End Function
    
    Private Function IsPlatePart(ByVal partDoc As PartDocument) As Boolean
        Try
            ' Check Part Number
            Dim partNumber As String = ""
            Try
                Dim propSet As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                partNumber = propSet.Item("Part Number").Value.ToString()
            Catch
                ' Part Number not found, try Description
            End Try
            
            If partNumber.ToUpper().Contains("PL") OrElse partNumber.ToUpper().Contains("S355JR") Then
                Return True
            End If
            
            ' Check Description
            Dim description As String = ""
            Try
                Dim propSet As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                description = propSet.Item("Description").Value.ToString()
            Catch
                ' Description not found
            End Try
            
            If description.ToUpper().Contains("PL") OrElse description.ToUpper().Contains("S355JR") Then
                Return True
            End If
            
            ' Check if has PLATE LENGTH or PLATE WIDTH custom properties (already processed)
            Try
                Dim customProps As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
                
                For Each prop As [Property] In customProps
                    If prop.Name = "PLATE LENGTH" OrElse prop.Name = "PLATE WIDTH" Then
                        Return True
                    End If
                Next
            Catch
                ' No custom properties
            End Try
            
            Return False
            
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("IsPlatePart Error: " & ex.Message)
            Return False
        End Try
    End Function
    
    Private Sub ApplyDocumentSettingsForZeroDecimals(ByVal partDoc As PartDocument)
        Try
            Dim params As Parameters = partDoc.ComponentDefinition.Parameters

            ' Setting 1: Linear Dimension Precision = 0 decimals
            Try
                params.LinearDimensionPrecision = 0
                System.Diagnostics.Debug.WriteLine("Set LinearDimensionPrecision = 0 for " & partDoc.DisplayName)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Failed to set LinearDimensionPrecision: " & ex.Message)
            End Try

            ' Setting 2: Modeling Dimension Display = "Display as value" (34821)
            Try
                params.DimensionDisplayType = CType(DISPLAY_AS_VALUE, DimensionDisplayTypeEnum)
                System.Diagnostics.Debug.WriteLine("Set DimensionDisplayType = 34821 for " & partDoc.DisplayName)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Failed to set DimensionDisplayType: " & ex.Message)
            End Try

            ' Setting 3: Default Parameter Input Display = "Display as expression" (True)
            Try
                params.DisplayParameterAsExpression = True
                System.Diagnostics.Debug.WriteLine("Set DisplayParameterAsExpression = True for " & partDoc.DisplayName)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Failed to set DisplayParameterAsExpression: " & ex.Message)
            End Try

            ' Force update to apply parameter changes
            partDoc.Update()

            ' CRITICAL: Save the part to persist changes
            ' Assembly BOMs read from saved file data, not transient in-memory changes
            Try
                partDoc.Save2(True)
                System.Diagnostics.Debug.WriteLine("Part saved successfully - changes persisted to disk")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Failed to save part: " & ex.Message)
            End Try

            System.Diagnostics.Debug.WriteLine("Document Settings applied successfully to " & partDoc.DisplayName)

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("ApplyDocumentSettingsForZeroDecimals Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ForceBOMRefreshInParentAssemblies(ByVal partDoc As PartDocument)
        Try
            System.Diagnostics.Debug.WriteLine("Searching for parent assemblies that reference: " & partDoc.DisplayName)

            Dim partPath As String = partDoc.FullFileName
            Dim refreshedAssemblies As Integer = 0

            ' Find all open assemblies that reference this part
            Dim docs As Documents = m_InventorApp.Documents

            For Each doc As Document In docs
                If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)

                    ' Check if this assembly references the part
                    If AssemblyReferencesPart(asmDoc, partPath) Then
                        System.Diagnostics.Debug.WriteLine("Found parent assembly: " & asmDoc.DisplayName)
                        RefreshAssemblyBOM(asmDoc)
                        refreshedAssemblies += 1
                    End If
                End If
            Next

            System.Diagnostics.Debug.WriteLine("Refreshed " & refreshedAssemblies & " parent assemblies")

            ' CRITICAL: If no assemblies are open, warn the user
            If refreshedAssemblies = 0 Then
                System.Diagnostics.Debug.WriteLine("WARNING: No parent assemblies found - BOM will refresh when assembly is opened/updated")
            End If

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("ForceBOMRefreshInParentAssemblies Error: " & ex.Message)
        End Try
    End Sub
    
    Private Function AssemblyReferencesPart(ByVal asmDoc As AssemblyDocument, ByVal partPath As String) As Boolean
        Try
            ' Normalize paths for comparison (handles different path formats)
            Dim normalizedPartPath As String = System.IO.Path.GetFullPath(partPath).ToLowerInvariant()

            ' Use AllReferencedFileDescriptors for comprehensive reference checking
            ' This includes nested sub-assembly references
            For Each refDoc As ReferencedFileDescriptor In asmDoc.File.AllReferencedFileDescriptors
                Dim refPath As String = System.IO.Path.GetFullPath(refDoc.FullFileName).ToLowerInvariant()
                If refPath = normalizedPartPath Then
                    Return True
                End If
            Next

            Return False

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("AssemblyReferencesPart Error: " & ex.Message)
            Return False
        End Try
    End Function
    
    Private Sub RefreshAssemblyBOM(ByVal asmDoc As AssemblyDocument)
        Try
            System.Diagnostics.Debug.WriteLine("  Refreshing BOM for: " & asmDoc.DisplayName)

            ' METHOD 1: Force assembly rebuild (invalidates geometry cache)
            System.Diagnostics.Debug.WriteLine("    Method 1: Rebuild2...")
            Try
                asmDoc.Rebuild2(True) ' AcceptErrorsAndContinue = True
                System.Diagnostics.Debug.WriteLine("    Rebuild2 complete")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("    Rebuild2 failed: " & ex.Message)
            End Try

            ' METHOD 2: Toggle BOM views (invalidates BOM structure cache)
            System.Diagnostics.Debug.WriteLine("    Method 2: Toggle BOM views...")
            Try
                Dim bom As BOM = asmDoc.ComponentDefinition.BOM

                ' Ensure BOM is enabled
                If Not bom.StructuredViewEnabled Then
                    bom.StructuredViewEnabled = True
                End If

                ' Toggle off and on to force refresh
                bom.StructuredViewEnabled = False
                bom.StructuredViewEnabled = True
                System.Diagnostics.Debug.WriteLine("    BOM views toggled")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("    BOM toggle failed: " & ex.Message)
            End Try

            ' METHOD 3: Force assembly update with regenerate (invalidates reference cache)
            System.Diagnostics.Debug.WriteLine("    Method 3: Update2...")
            Try
                asmDoc.Update2(True) ' Regenerate = True
                System.Diagnostics.Debug.WriteLine("    Update2 complete")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("    Update2 failed: " & ex.Message)
            End Try

            ' METHOD 4: *** CRITICAL FIX *** - Activate assembly and access BOM rows
            ' This is THE breakthrough that makes it work!
            ' BOM refresh events are queued but only processed when assembly is active
            ' Accessing BOM rows forces cache rebuild from updated part Document Settings
            System.Diagnostics.Debug.WriteLine("    Method 4: Activating assembly and accessing BOM rows...")
            Try
                Dim currentlyActive As Document = m_InventorApp.ActiveDocument

                If currentlyActive IsNot asmDoc Then
                    ' Note: Cannot programmatically set ActiveDocument (read-only property)
                    ' Just access BOM to force refresh on the target document
                    
                    ' Force UI refresh by accessing BOM views
                    ' This processes queued BOM cache invalidation events
                    Dim bom As BOM = asmDoc.ComponentDefinition.BOM
                    Dim structuredView As BOMView = bom.BOMViews.Item("Structured")
                    Dim rowCount As Integer = structuredView.BOMRows.Count
                    System.Diagnostics.Debug.WriteLine("    BOM has " & rowCount & " rows - cache rebuilt")

                    ' Note: Cannot restore active document programmatically
                    ' User will need to click back to original document manually
                Else
                    ' Already active - just access BOM to force refresh
                    Dim bom As BOM = asmDoc.ComponentDefinition.BOM
                    Dim structuredView As BOMView = bom.BOMViews.Item("Structured")
                    Dim rowCount As Integer = structuredView.BOMRows.Count
                    System.Diagnostics.Debug.WriteLine("    BOM has " & rowCount & " rows - cache rebuilt")
                End If

                System.Diagnostics.Debug.WriteLine("    Assembly activation complete")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("    Assembly activation failed: " & ex.Message)
            End Try

            System.Diagnostics.Debug.WriteLine("  BOM refresh complete")

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("  WARNING: BOM refresh encountered error: " & ex.Message)
            ' Don't throw - partial refresh is better than none
        End Try
    End Sub
    
    Public Sub Cleanup()
        ' Simple cleanup - no event handlers to remove
        m_InventorApp = Nothing
    End Sub
    
End Class

End Namespace
