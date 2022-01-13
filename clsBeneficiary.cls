VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBeneficiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_id As Integer
Private m_name As String
Private m_relation As String
Private m_level As String
Private m_percent As Double
Private m_addDate As String
Private m_addedBy As String
Private m_updatedDate As String
Private m_updatedBy As String
Private m_account As clsAccount

Private Sub Class_Initialize()
    Set m_account = New clsAccount
End Sub

Public Property Get ID() As Integer
    ID = m_id
End Property

Public Property Let ID(value As Integer)
    m_id = value
End Property

Public Property Get NameOfBeneficiary() As String
    NameOfBeneficiary = m_name
End Property

Public Property Let NameOfBeneficiary(value As String)
    m_name = value
End Property

Public Property Get Relation() As String
    Relation = m_relation
End Property

Public Property Let Relation(value As String)
    m_relation = value
End Property

Public Property Get Level() As String
    Level = m_level
End Property

Public Property Let Level(value As String)
    m_level = value
End Property

Public Property Get Percent() As Double
    Percent = m_percent
End Property

Public Property Let Percent(value As Double)
    m_percent = value
End Property

Public Property Get AddDate() As String
    AddDate = m_addDate
End Property

Public Property Let AddDate(value As String)
    m_addDate = value
End Property

Public Property Get AddedBy() As String
    AddedBy = m_addedBy
End Property

Public Property Let AddedBy(value As String)
    m_addedBy = value
End Property

Public Property Get UpdatedDate() As Date
    UpdatedDate = m_updatedDate
End Property

Public Property Let UpdatedDate(value As Date)
    m_updatedDate = value
End Property

Public Property Get UpdatedBy() As String
    UpdatedBy = m_updatedBy
End Property

Public Property Let UpdatedBy(value As String)
    m_updatedBy = value
End Property

Public Property Get account() As clsAccount
    Set account = m_account
End Property

Public Property Let account(value As clsAccount)
    Set m_account = value
End Property
