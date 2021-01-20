VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gitExportForm 
   Caption         =   "Git Export"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "gitExportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "gitExportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    Set wkb = ThisWorkbook
    Dim component As VBIDE.VBComponent
    
    For Each component In wkb.VBProject.VBComponents
        
        Select Case component.Type
        
            Case vbext_ct_ClassModule
                moduleName = component.Name & ".cls"
                
            Case vbext_ct_MSForm
                moduleName = component.Name & ".frm"
                
            Case vbext_ct_StdModule
                moduleName = component.Name & ".bas"
                
            Case vbext_ct_Document
                moduleName = ""
                                
        End Select
        
        If moduleName <> "" Then moduleList.AddItem moduleName

    Next

End Sub

Private Sub selectAllButton_Click()

    For i = 0 To moduleList.ListCount - 1
    
        moduleList.Selected(i) = True
    
    Next i

End Sub

Private Sub gitFolderImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    gitFolder = OpenFolderDialog()
    
    If gitFolder <> "" Then gitFolderLabel.Caption = gitFolder
   
End Sub

Private Sub exportFiles_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim modulesArray()
    Dim ignoreFrx As Boolean
    arrayBound = 0
    noneSelected = True
    
    If gitFolderLabel.Caption = "Select the Git path using the folder icon." Then
    
        MsgBox "Select a valid folder and try again."
        
        Exit Sub
    
    End If
    
    For i = 0 To moduleList.ListCount - 1
    
        If moduleList.Selected(i) = True Then
            
            noneSelected = False
        
            ReDim Preserve modulesArray(0 To arrayBound)
            
            modulesArray(arrayBound) = moduleList.List(i)
            
            arrayBound = arrayBound + 1
        
        End If
    
    Next i
    
    If noneSelected = True Then
    
        MsgBox "Please select at least one file to export."
        
        Exit Sub
    
    End If
    
    If ignoreFrxCheck.Value = True Then
    
        ignoreFrx = True
    
    Else
    
        ignoreFrx = False
    
    End If
    
    GitSave gitFolderLabel.Caption, ignoreFrx, modulesArray

End Sub
