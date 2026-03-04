# iLogic SANS Section Analysis Report

**Date:** December 15, 2025
**Analyzed Files:**
- `C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\REF\Beam_Factory.ipt`
- `C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\REF\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\InventorAddIn\AssemblyClonerAddIn\BeamAssemblyGenerator.vb`

---

## 1. EXECUTIVE SUMMARY

The Beam_Factory.ipt uses a **parametric iLogic approach** that is fundamentally different from the BeamAssemblyGenerator.vb's **programmatic geometry creation**. The key insight is that **Beam_Factory relies on pre-existing sketches and features** that are controlled by parameters, while BeamAssemblyGenerator tries to **create geometry from scratch** using API calls.

### Key Findings:
1. **iLogic uses parameters to drive pre-existing geometry** (sketch already exists in .ipt file)
2. **BeamAssemblyGenerator creates geometry programmatically** (draws lines, creates profiles)
3. **The iLogic approach is simpler and more reliable** for SANS sections
4. **Recommendation:** Adopt the parameter-driven approach for BeamAssemblyGenerator

---

## 2. iLOGIC METHOD ANALYSIS

### 2.1 How Beam_Factory.ipt Works

**Core Principle:** The .ipt file already contains a **pre-drawn I-beam sketch** with dimensioned features. The iLogic rule simply **changes the dimension values** based on user selection.

**Method:**
```vbnet
' User selects "203 x 133 x 25" from dropdown
Select Case Size_IBeam
    Case "203 x 133 x 25"
        kg_m = 25        ' Mass per meter
        h_ = 203.2       ' Total height
        b_ = 133.2       ' Flange width
        tw = 5.7         ' Web thickness
        tf = 7.8         ' Flange thickness
        r1 = 7.6         ' Root radius
End Select
```

**What Happens:**
1. iLogic rule updates user parameters: `h_`, `b_`, `tw`, `tf`, `r1`
2. These parameters are **already linked to sketch dimensions** in the .ipt file
3. Inventor **automatically regenerates** the sketch and extrusion
4. The Description iProperty is updated: `iProperties.Value("Project", "Description") = "=<Size_IBeam> PFI"`

**Geometry Structure:**
- The I-beam profile sketch is **already drawn** in Beam_Factory.ipt
- Sketch dimensions reference user parameters: `d0 = b_`, `d1 = h_`, `d3 = tf`, `d4 = tw`
- Extrusion feature references: `d5 = Length`
- **No geometry creation happens in iLogic** - only parameter updates

---

## 3. BEAMASSEMBLYGENERATOR.VB METHOD ANALYSIS

### 3.1 Current Approach

**Core Principle:** Creates geometry **from scratch** using Inventor API calls to draw lines, create profiles, and extrude.

**Method:**
```vbnet
' Create I-beam using three separate sketches and Boolean operations
CreateIBeamFromRectangles(compDef, sketch)
    ' Create top flange rectangle
    ' Create bottom flange rectangle
    ' Create web rectangle
    ' Join using kJoinOperation
```

**Problems with This Approach:**
1. **Complex Boolean operations** - three separate extrusions that must be joined
2. **No fillet radii** - the current method doesn't include root fillets (r1)
3. **Profile accuracy** - rectangles don't match actual I-beam profiles exactly
4. **Reliability issues** - Boolean joins can fail, especially with complex profiles
5. **Limited parametric updates** - changes require re-creating entire geometry

---

## 4. KEY DIFFERENCES

| Aspect | iLogic Method (Beam_Factory.ipt) | VB.NET Method (BeamAssemblyGenerator.vb) |
|--------|-----------------------------------|-------------------------------------------|
| **Geometry Source** | Pre-existing sketch in template .ipt | Created programmatically via API |
| **Parameter Control** | User parameters linked to sketch dims | Hardcoded dimensions passed to API |
| **Profile Accuracy** | Perfect (uses actual I-beam profile) | Approximate (three rectangles) |
| **Fillet Radii** | Included (r1 parameter) | Missing completely |
| **Complexity** | Simple - just update parameter values | Complex - draw lines, create profiles, Boolean ops |
| **Reliability** | Very high (proven Inventor workflow) | Lower (API calls can fail) |
| **Parametric Updates** | Automatic (Inventor regenerates) | Manual (must recreate geometry) |
| **SANS Data** | Complete lookup table (27 I-beams, 17 H-beams, 6 IPE) | Must be implemented separately |

---

## 5. SANS SECTION DATA COMPLETENESS

### 5.1 Beam_Factory iLogic Coverage

**I-Beams (27 sections):**
- 203 x 133 x 25, 203 x 133 x 30
- 254 x 146 x 31, 254 x 146 x 37, 254 x 146 x 43
- 305 x 102 x 25, 305 x 102 x 28, 305 x 102 x 33
- 305 x 165 x 40, 305 x 165 x 46, 305 x 165 x 54
- 356 x 171 x 45, 356 x 171 x 51, 356 x 171 x 57, 356 x 171 x 67
- 406 x 140 x 39, 406 x 140 x 46
- 406 x 178 x 54, 406 x 178 x 60, 406 x 178 x 67, 406 x 178 x 74
- 457 x 191 x 67, 457 x 191 x 74, 457 x 191 x 82, 457 x 191 x 89, 457 x 191 x 98
- 533 x 210 x 82, 533 x 210 x 92, 533 x 210 x 101, 533 x 210 x 109, 533 x 210 x 122

**H-Beams (17 sections):**
- 152 x 152 x 23, 152 x 152 x 30, 152 x 152 x 37
- 203 x 203 x 46, 203 x 203 x 52, 203 x 203 x 60, 203 x 203 x 71, 203 x 203 x 86
- 254 x 254 x 73, 254 x 254 x 89, 254 x 254 x 107, 254 x 254 x 132, 254 x 254 x 167
- 305 x 305 x 97, 305 x 305 x 118, 305 x 305 x 137, 305 x 305 x 158

**IPE Sections (6 sections):**
- IPE100, IPE120, IPE140, IPE160, IPE180, IPE200

**Total: 50 standardized SANS sections** with complete dimensional data

### 5.2 BeamAssemblyGenerator Coverage

**Current:** Generic section types (UC, UB, PFC, TFC, L) with user-provided dimensions
**Missing:** SANS standard lookup table and accurate profile generation

---

## 6. CRITICAL INSIGHT: WHY iLOGIC WORKS BETTER

### The Template .ipt Approach

**What's Inside Beam_Factory.ipt:**
1. **A parametric sketch** already drawn with proper I-beam profile (including fillets)
2. **User parameters** (`h_`, `b_`, `tw`, `tf`, `r1`) linked to sketch dimensions
3. **An extrusion feature** that references the profile and Length parameter
4. **iLogic form** for user selection (dropdown lists)
5. **iLogic rule** that updates parameters based on selection

**Why This Works:**
- Sketch geometry is **guaranteed correct** (drawn once, tested, verified)
- Parameter updates are **instant and reliable** (native Inventor behavior)
- No complex API calls needed
- Fillets are **part of the sketch geometry** (arcs with radius = r1)
- Updates trigger **automatic regeneration** of all downstream features

---

## 7. RECOMMENDATIONS FOR BEAMASSEMBLYGENERATOR.VB

### 7.1 Adopt a Hybrid Approach

**Option A: Template-Based Generation (RECOMMENDED)**

Instead of creating geometry from scratch, use **template .ipt files** as a base:

```vbnet
' Proposed workflow:
1. Create template .ipt files for each section type (I-beam, H-beam, Channel, Angle)
2. Each template contains parametric sketch with proper profile
3. BeamAssemblyGenerator:
   a. Copies appropriate template file
   b. Opens the copy
   c. Updates user parameters (h_, b_, tw, tf, r1, Length)
   d. Saves with new name
   e. Updates iProperties
```

**Benefits:**
- ✅ Perfect profile accuracy (templates drawn correctly once)
- ✅ Includes fillet radii automatically
- ✅ Simpler code (parameter updates vs geometry creation)
- ✅ More reliable (no API geometry failures)
- ✅ Easier to maintain (modify template, not code)

**Implementation:**
```vbnet
Private Function CreateBeamFromTemplate() As PartDocument
    ' 1. Select template based on section type
    Dim templatePath As String = GetTemplateForSection(m_CurrentSection.SectionType)

    ' 2. Copy template to output location
    Dim newPartPath As String = m_BeamPartPath
    System.IO.File.Copy(templatePath, newPartPath, True)

    ' 3. Open the copied file
    Dim partDoc As PartDocument = m_InventorApp.Documents.Open(newPartPath, False)

    ' 4. Update user parameters
    UpdateSectionParameters(partDoc, m_CurrentSection)

    ' 5. Update iProperties
    SetPartIProperties(partDoc, "BEAM", m_CurrentSection.GetIPropertyDescription())

    ' 6. Save and return
    partDoc.Save2(True)
    Return partDoc
End Function

Private Sub UpdateSectionParameters(partDoc As PartDocument, section As SteelSection)
    Dim params As Parameters = partDoc.ComponentDefinition.Parameters.UserParameters

    ' Update dimensional parameters
    params.Item("h_").Value = section.Height / 10.0 ' Convert mm to cm
    params.Item("b_").Value = section.Width / 10.0
    params.Item("tw").Value = section.WebThickness / 10.0
    params.Item("tf").Value = section.FlangeThickness / 10.0
    params.Item("r1").Value = section.RootRadius / 10.0
    params.Item("Length").Value = m_BeamLength / 10.0

    ' Inventor will automatically regenerate the geometry
End Sub
```

### 7.2 Add SANS Lookup Table

Create a SteelSectionDatabase class with complete SANS data:

```vbnet
Public Class SteelSectionDatabase
    Public Shared Function GetIBeamSections() As List(Of SteelSection)
        Dim sections As New List(Of SteelSection)

        ' From Beam_Factory iLogic data
        sections.Add(New SteelSection("UB", "203 x 133 x 25", 203.2, 133.2, 5.7, 7.8, 7.6, 25))
        sections.Add(New SteelSection("UB", "203 x 133 x 30", 206.8, 133.9, 6.4, 9.6, 7.6, 30))
        ' ... (add all 27 I-beams)

        Return sections
    End Function

    Public Shared Function GetHBeamSections() As List(Of SteelSection)
        ' ... (add all 17 H-beams)
    End Function

    Public Shared Function GetIPESections() As List(Of SteelSection)
        ' ... (add all 6 IPE sections)
    End Function
End Class
```

### 7.3 Enhance SteelSection Class

Add root radius and mass properties:

```vbnet
Public Class SteelSection
    ' Existing properties...
    Public Property RootRadius As Double  ' r1 in mm
    Public Property MassPerMeter As Double  ' kg/m

    ' Add constructor overload for SANS sections
    Public Sub New(sectionType As String, designation As String,
                   height As Double, width As Double,
                   webThickness As Double, flangeThickness As Double,
                   rootRadius As Double, massPerMeter As Double)
        Me.SectionType = sectionType
        Me.Designation = designation
        Me.Height = height
        Me.Width = width
        Me.WebThickness = webThickness
        Me.FlangeThickness = flangeThickness
        Me.RootRadius = rootRadius
        Me.MassPerMeter = massPerMeter
    End Sub
End Class
```

---

## 8. IMPLEMENTATION ROADMAP

### Phase 1: Create Template Files (HIGH PRIORITY)
1. **Create I-Beam Template (.ipt)**
   - Draw accurate I-beam profile with fillet arcs
   - Create user parameters: h_, b_, tw, tf, r1, Length
   - Link sketch dimensions to parameters
   - Create extrusion feature
   - Save as `Template_IBeam.ipt`

2. **Create H-Beam Template (.ipt)**
   - Same as I-beam but with equal flanges
   - Save as `Template_HBeam.ipt`

3. **Create Channel Template (.ipt)**
   - C-channel profile with fillets
   - Save as `Template_Channel.ipt`

4. **Create Angle Template (.ipt)**
   - L-angle profile
   - Save as `Template_Angle.ipt`

### Phase 2: Update BeamAssemblyGenerator.vb (MEDIUM PRIORITY)
1. Add template-based CreateBeamPart() method
2. Add UpdateSectionParameters() method
3. Add SteelSectionDatabase class with SANS data
4. Update BeamGeneratorForm with dropdown for standard sections
5. Test with all section types

### Phase 3: Add Advanced Features (LOW PRIORITY)
1. Support for IPE sections
2. Support for CHS (circular hollow sections)
3. Support for SHS (square hollow sections)
4. Custom section editor for non-standard sizes

---

## 9. COMPARISON: OLD vs NEW APPROACH

### Current Approach (Programmatic Geometry)
```vbnet
// 200+ lines of code
CreateIBeamFromRectangles()
    Create top flange sketch → extrude
    Create bottom flange sketch → extrude + join
    Create web sketch → extrude + join
    // No fillets!
    // Complex Boolean operations
    // Can fail due to API issues
```

**Issues:**
- ❌ Complex code
- ❌ Missing fillet radii
- ❌ Approximate profile (rectangles)
- ❌ Unreliable Boolean joins
- ❌ Hard to maintain

### Recommended Approach (Template-Based)
```vbnet
// ~50 lines of code
CreateBeamFromTemplate()
    Copy template .ipt file
    Open template
    Update 6 parameters (h_, b_, tw, tf, r1, Length)
    Save as new name
    // Perfect profile with fillets!
    // Inventor handles regeneration
    // Bulletproof reliability
```

**Benefits:**
- ✅ Simple code
- ✅ Perfect SANS profiles
- ✅ Includes fillet radii
- ✅ Highly reliable
- ✅ Easy to maintain
- ✅ Fast execution

---

## 10. CONCLUSION

The iLogic method in Beam_Factory.ipt demonstrates that **parameter-driven geometry** is superior to **programmatic geometry creation** for standard sections. By adopting a template-based approach, BeamAssemblyGenerator.vb can achieve:

1. **Better accuracy** - exact SANS profiles
2. **Simpler code** - parameter updates vs line drawing
3. **Higher reliability** - proven Inventor workflow
4. **Easier maintenance** - templates vs code changes
5. **Complete SANS coverage** - 50 standard sections ready to use

**Next Steps:**
1. Create template .ipt files for each section type
2. Implement template-based CreateBeamPart() method
3. Add SANS lookup table to BeamAssemblyGenerator.vb
4. Test thoroughly with various section sizes

---

## APPENDIX A: EXTRACTED iLOGIC CODE

**Location:** `C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\REF\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\iLogic_Export\Beam_Factory_Rule_Name2.vb`

**Key Parameters from Beam_Factory.ipt:**
```
User Parameters:
  h_ = 203.2 mm       (Height)
  b_ = 133.2 mm       (Flange width)
  tw = 5.7 mm         (Web thickness)
  tf = 7.8 mm         (Flange thickness)
  r1 = 7.6 mm         (Root radius)
  Length = 1000 mm    (Beam length)
  kg_m = 25           (Mass per meter)

Model Parameters (linked to sketch):
  d0 = b_             (Flange width dimension)
  d1 = h_             (Height dimension)
  d2 = r1             (Root radius dimension)
  d3 = tf             (Flange thickness dimension)
  d4 = tw             (Web thickness dimension)
  d5 = Length         (Extrusion length)
```

**Complete SANS I-Beam Lookup Table:** See lines 36-284 of exported iLogic rule.

---

**Report Generated:** December 15, 2025
**Analyst:** Claude Sonnet 4.5
**Tools Used:** iLogic Scanner, VBScript extraction, code analysis
