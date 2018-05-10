Attribute VB_Name = "AutoNew"
' Naming Conventions
' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/naming-conventions

' Naming conventions for controls: https://msdn.microsoft.com/en-us/library/aa263493(v=vs.60).aspx
' E.g. use prefix txt for control of type Text Box

Dim part As CustomXMLPart
Dim node As CustomXMLNode

Dim pp As New PrivateProfile

Private Const namespaceCore     As String = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
Private Const namespaceExt      As String = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
Private Const namespaceCover    As String = "http://schemas.microsoft.com/office/2006/coverPageProps"
Private Const namespaceId       As String = "http://schemas.htwchur.ch/identification"

Sub Main()
    ' XMLParts.DebugPrintNamespace (namespaceCore)

    Dim idTitle As identification
    Set idTitle = Factory.CreateIdentification("title", "dc:title", namespaceCore)
    
    Dim idSubject As identification
    Set idSubject = Factory.CreateIdentification("subject", "dc:subject", namespaceCore)
    
    Dim idAuthor As identification
    Set idAuthor = Factory.CreateIdentification("author", "dc:creator", namespaceCore)
    
    Dim idDate As identification
    Set idDate = Factory.CreateIdentification("issuingDate", "PublishDate", namespaceCover)

    Dim idClassification As identification
    Set idClassification = Factory.CreateIdentification("classification", "cp:category", namespaceCore)
    
    Dim idIssuingOffice As identification
    Set idIssuingOffice = Factory.CreateIdentification("issuingOffice", "Manager", namespaceExt)
    
    Dim idScope As identification
    Set idScope = Factory.CreateIdentification("scope", "Company", namespaceExt)
    
    Dim idVersion As identification
    Set idVersion = Factory.CreateIdentification("version", "cp:contentStatus", namespaceCore)
    
    Dim idDistribution As identification
    Set idDistribution = Factory.CreateIdentification("distribution", "distribution", namespaceId)
   
    Dim idLogo As ContentControlIdentification
    Set idLogo = Factory.CreateContentControlIdentification("logo", "logo")
    
    ' title
    Dim titleTBC As TextBoxConnector
    Set titleTBC = Factory.CreateTextBoxConnector(idTitle, pp, idForm.txtTitle)
    Call titleTBC.LoadAndSet
    
    ' subject
    ' Dim subjectTBC As TextBoxConnector
    ' Set subjectTBC = Factory.CreateTextBoxConnector(idSubject, pp, idForm.txtSubject)
    ' Call subjectTBC.LoadAndSet
    
    ' author
    Dim authorTBC As TextBoxConnector
    Set authorTBC = Factory.CreateTextBoxConnector(idAuthor, pp, idForm.txtAuthor)
    Call authorTBC.LoadAndSet
        
    ' date
    Dim dateDTPC As DTPickerConnector
    Set dateDTPC = Factory.CreateDTPickerConnector(idDate, pp, idForm.dtpDate)
    Call dateDTPC.LoadAndSet

    ' classification
    Dim classificationCCC As ComboBoxConnector
    Set classificationCCC = Factory.CreateComboBoxConnector(idClassification, pp, idForm.cbClassification)
    Call classificationCCC.LoadAndSetComboBox
    
    ' issuing office
    Dim issuingOfficeCCC As ComboBoxConnector
    Set issuingOfficeCCC = Factory.CreateComboBoxConnector(idIssuingOffice, pp, idForm.cbIssuingOffice)
    Call issuingOfficeCCC.LoadAndSetComboBox
    
    ' scope
    Dim scopeCCC As ComboBoxConnector
    Set scopeCCC = Factory.CreateComboBoxConnector(idScope, pp, idForm.cbScope)
    Call scopeCCC.LoadAndSetComboBox
    
    ' version
    Dim versionTBC As TextBoxConnector
    Set versionTBC = Factory.CreateTextBoxConnector(idVersion, pp, idForm.txtVersion)
    Call versionTBC.LoadAndSet
    
    ' distribution
    Dim distributionTBC As TextBoxConnector
    Set distributionTBC = Factory.CreateTextBoxConnector(idDistribution, pp, idForm.txtDistribution)
    Call distributionTBC.LoadAndSet
    
    ' logo
    Dim logoBBC As BuildingBlockConnector
    Set logoBBC = Factory.CreateBuildingBlockConnector(idLogo, pp, idForm.cbLogo)
    logoBBC.LoadAndSetComboBox
    
    idForm.Show
    
    If idForm.cancelled Then
        Exit Sub
    End If
    
    
    Call titleTBC.Store
    ' Call subjectTBC.Store
    Call authorTBC.Store
    Call dateDTPC.Store
    Call classificationCCC.Store
    Call issuingOfficeCCC.Store
    Call scopeCCC.Store
    Call versionTBC.Store
    Call distributionTBC.Store
    Call logoBBC.Store
    
    Dim revisionControl As Table
    Dim t As Table
    For Each t In ActiveDocument.Tables
        If t.Title = "Änderungskontrolle" Then
            Set revisionControl = t
        End If
    Next
    If revisionControl Is Nothing Then
    Else
        Dim cc As contentControl
        For Each cc In revisionControl.Range.ContentControls
            cc.Delete
        Next
    End If
    
End Sub


