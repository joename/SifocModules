Attribute VB_Name = "G_Form"
Option Explicit
Option Compare Database

'-------------------------------------------------------------
'   isFormOpened
'-------------------------------------------------------------
Public Function isFormOpened(formName As String) As Boolean
    If CurrentProject.AllForms(formName).IsLoaded Then
        isFormOpened = True
    Else
        isFormOpened = False
    End If
End Function

Public Static Function openForm(formName As String, _
                         Optional formView As AcFormView = AcFormView.acNormal, _
                         Optional filterName As String = "", _
                         Optional whereCondition As String = "", _
                         Optional dataMode As AcFormOpenDataMode = AcFormOpenDataMode.acFormPropertySettings, _
                         Optional windowMode As AcWindowMode = AcWindowMode.acWindowNormal, _
                         Optional openargs As String = "")

    Dim frm As New Form
    
    If isFormOpened(formName) Then
        DoCmd.Close acForm, formName
    End If
    
    DoCmd.SetWarnings False
    
    DoCmd.openForm formName, formView, filterName, whereCondition, dataMode, windowMode, openargs
    
    DoCmd.SetWarnings True
    
    'Debug.Print CurrentProject.AllForms.Item("BuscarPersona")
    
    'Debug.Print Forms(FormName).form.name
    
    'Dim frm As New forms("BuscarPersona")
    
    'frm.NavigationButtons = False
    'frm.visible = True
    
    'LigaFormulario.LigaFormulario frm
    
End Function

Public Function AllReports()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllReports
        If obj.IsLoaded = False Then
            ' Print name of obj.
            Debug.Print "Report_" & obj.name
        End If
    Next obj
End Function

Public Function AllForms()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllForms
        If obj.IsLoaded = False Then
            ' Print name of obj.
            Debug.Print "Form_" & obj.name
        End If
    Next obj
End Function

Public Function AllModules()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllModules collection.
    For Each obj In dbs.AllModules
        If obj.IsLoaded = True Then
            ' Print name of obj.
            Debug.Print obj.name
        End If
    Next obj
End Function

Public Function AllMacros()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllMacros collection.
    For Each obj In dbs.AllMacros
        If obj.IsLoaded = True Then
            ' Print name of obj.
            Debug.Print obj.name
        End If
    Next obj
End Function

