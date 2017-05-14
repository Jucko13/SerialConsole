Attribute VB_Name = "ObjectSearcher"
Sub GetMembers(pList As ListBox, pObject As Object)
Dim TLI         As TLIApplication
Dim lInterface  As InterfaceInfo
Dim lMember     As MemberInfo
        
    Set TLI = New TLIApplication
    Set lInterface = TLI.InterfaceInfoFromObject(pObject)
    pList.Clear
    
    For Each lMember In lInterface.Members
        pList.AddItem lMember.Name & " - " & WhatIsIt(lMember)
    Next
    
    'Set pObject = Nothing
    Set lInterface = Nothing
    Set TLI = Nothing
End Sub

Private Function WhatIsIt(lMember As Object) As String
    Select Case lMember.InvokeKind
        Case INVOKE_FUNC
            If lMember.ReturnType.VarType <> VT_VOID Then
                WhatIsIt = "Function"
            Else
                WhatIsIt = "Method"
            End If
        Case INVOKE_PROPERTYGET
            WhatIsIt = "Property Get"
        Case INVOKE_PROPERTYPUT
            WhatIsIt = "Property Let"
        Case INVOKE_PROPERTYPUTREF
            WhatIsIt = "Property Set"
        Case INVOKE_CONST
            WhatIsIt = "Const"
        Case INVOKE_EVENTFUNC
            WhatIsIt = "Event"
        Case Else
            WhatIsIt = lMember.InvokeKind & " (Unknown)"
    End Select
End Function
