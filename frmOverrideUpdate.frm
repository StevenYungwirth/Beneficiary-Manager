VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOverrideUpdate 
   Caption         =   "Update Override"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2325
   OleObjectBlob   =   "frmOverrideUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOverrideUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_overrideType As String
Private m_overrideValue As String

Public Property Get OverrideType() As String
    OverrideType = m_overrideType
End Property

Public Property Get OverrideValue() As Variant
    OverrideValue = m_overrideValue
End Property

Private Sub UserForm_Initialize()
    'Start the form in the middle of the screen with Excel
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
'    'Hide the title bar
'    FormProcedures.HideBar Me
    
    'Hide the controls for override options
    fraActive.Visible = False
    txtTag.Visible = False
    
    'Disable the OK button
    btnUpdate.Enabled = False
End Sub

Private Sub rdoActive_Click()
    'Show the true/false radios
    fraActive.Visible = True
    
    'Hide the tag textbox
    txtTag.Visible = False
    
    'Reset the controls
    rdoTrue.value = False
    rdoFalse.value = False
    txtTag.value = vbNullString
    
    'Disable the OK button
    btnUpdate.Enabled = False
End Sub

Private Sub rdoTag_Click()
    'Hide the true/false radios
    fraActive.Visible = False
    
    'Show the tag textbox
    txtTag.Visible = True
    
    'Reset the controls
    rdoTrue.value = False
    rdoFalse.value = False
    txtTag.value = vbNullString
    
    'Disable the OK button
    btnUpdate.Enabled = False
End Sub

Private Sub rdoTrue_Click()
    'Enable the OK button
    btnUpdate.Enabled = True
End Sub

Private Sub rdoFalse_Click()
    'Enable the OK button
    btnUpdate.Enabled = True
End Sub

Private Sub txtTag_Change()
    'Enable the OK button if the textbox isn't empty
    btnUpdate.Enabled = (txtTag.value <> vbNullString)
End Sub

Private Sub btnUpdate_Click()
    'Set the override value
    If m_overrideType = "Active" Then
        If rdoTrue.value Then
            m_overrideValue = "True"
        ElseIf rdoFalse.value Then
            m_overrideValue = "False"
        End If
    ElseIf m_overrideType = "Tag" Then
        m_overrideValue = txtTag.value
    End If

    'Hide the form
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    'Hide the form
    Me.Hide
End Sub

Public Sub LoadOverride(overrideNode As IXMLDOMElement)
    'Get the selected node's override
    m_overrideType = overrideNode.Attributes(0).BaseName
    m_overrideValue = overrideNode.Attributes(0).Text
    
    'Set the proper controls to be visible
    Select Case m_overrideType
    Case "Active"
        fraActive.Visible = True
        rdoTrue.value = CBool(m_overrideValue)
        rdoFalse.value = Not CBool(m_overrideValue)
        txtTag.Visible = False
    Case "Tag"
        fraActive.Visible = False
        txtTag.Visible = True
        txtTag.value = m_overrideValue
    End Select
End Sub
