Set objUserClass = GetObject("LDAP://schema/user")
Set objSchemaClass = GetObject(objUserClass.Parent)

i = 0
WScript.Echo "Mandatory attributes:"
For Each strAttribute in objUserClass.MandatoryProperties
    i= i + 1
    strEcho = i & vbTab & strAttribute & vbTab
    Set objAttribute = objSchemaClass.GetObject("Property",  strAttribute)
    strEcho = strEcho & "(" & objAttribute.Syntax & ")" & vbTab
    If objAttribute.MultiValued Then
        strEcho = strEcho & "[Multivalued]"
    Else
        strEcho = strEcho & "[Single-valued]"
    End If
    WScript.Echo strEcho
Next

WScript.Echo VbCrLf & "Optional attributes:"
For Each strAttribute in objUserClass.OptionalProperties
    i=i + 1
    strEcho = i & vbTab & strAttribute & vbTab
    Set objAttribute = objSchemaClass.GetObject("Property",  strAttribute)
    strEcho = strEcho & "(" & objAttribute.Syntax & ")" & vbTab
    If objAttribute.MultiValued Then
        strEcho = strEcho & "[Multivalued]"
    Else
        strEcho = strEcho & "[Single-valued]"
    End If
    WScript.Echo strEcho
Next
