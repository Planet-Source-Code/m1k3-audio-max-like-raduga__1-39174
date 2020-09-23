Attribute VB_Name = "modCompName"
Private m_str_CompName As String

Property Get str_CompName() As String
    str_CompName = m_str_CompName
End Property

Property Let str_CompName(newValue As String)
    m_str_CompName = newValue
End Property

