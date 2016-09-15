Set objClass = New Person
'  For CEO  -> Scenario 1
msgbox objClass.GetCEODesignation("Schalk Nolte","GSMMobile")
' For Non Executive Members -> Scenario 2
msgbox objClass.NonExexutiveMembers("RAMZI","Nigeria")
msgbox objClass.NonExexutiveMembers("NATHAN","CapeTown")
msgbox objClass.NonExexutiveMembers("WILLEM","Nimbula")
msgbox objClass.NonExexutiveMembers("ANDREAS","Germany")



Class Person
  '  Scenario 1 :  Schalk Nolte is the CEO of Entersekt.
  '                         Assertion :  a CEO need to have GSM Mobile experience
Public function GetCEODesignation(strValue,strProfile)
  If ((strValue = "Schalk Nolte")  AND (strProfile="GSMMobile")) Then
        GetCEODesignation = "CEO"
        Exit Function
    End If
    GetCEODesignation = "Not CEO"
 End function

 '  Scenario 2 :    Assertion :  a Non exe member name should be one of Ramzi,NATHAN ,WILLEM ,ANDREAS  and their
 '                                                work location is not Capetown

Public function NonExexutiveMembers(strName,strRepresentation)
   if((strName = "RAMZI") AND (strRepresentation <>"CapeTown")) Then
   NonExexutiveMembers ="Non Executives"
   Exit Function
 End If
  IF ((strName = "NATHAN") AND (strRepresentation <>"CapeTown")) Then
   NonExexutiveMembers ="Non Executives"
   Exit Function

   End If

IF ((strName = "WILLEM") AND (strRepresentation <>"CapeTown")) Then
	  NonExexutiveMembers ="Non Executives"
	Exit Function
 End If

IF ((strName = "ANDREAS") AND (strRepresentation <>"CapeTown")) Then
	NonExexutiveMembers ="Non Executives"
Exit Function
End If


NonExexutiveMembers = "Not Non Executive "
End Function

End Class