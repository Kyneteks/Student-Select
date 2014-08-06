<%@  language="VBScript" %>
<%
    Option Explicit
    Dim bCache
    bCache = False
%>
<!--#Include File="inc/Includes.asp"-->
<%


Dim rsStudent
Dim strError
Dim Obj
Dim rsXML
Dim StudentUID
Dim Result
Dim IsAdvisee
'ADD - WILLIAM CAREY UNIVERSITY 
Dim rsCheckStudentUID
'END ADD
Dim arStudents
Dim StudentToken

TRSSession("IsAdvisee") = "False"
StudentUID = ""
IsAdvisee = "False"
'ADD - WILLIAM CAREY UNIVERSITY
StudentUID = Request.Form("txtStudentUID")
'END ADD
StudentToken = Request.Form("txtStudentUID")
arStudents = Deserialize(TRSSession("Advisees"))
If UBound(arStudents) > -1 Then
    If ValidateARElements("Advisees", StudentToken) Then
        StudentUID = arStudents(StudentToken)
    End If
'ADD - WILLIAM CAREY UNIVERSITY
End If
    
set obj = CreateObject("CAMSData.DataLayer")
call obj.Constructor(strSvrName, strDBName)
    
Set rsCheckStudentUID = obj.RunSQLReturnRS_RW("Select StudentUID From Student Where StudentUID=" & StudentUID)
If rsCheckStudentUID.eof or rsCheckStudentUID.bof Then
    Response.Redirect("ceSelectStudent.asp?Failed=" & StudentUID)
'END ADD
End If

If TestMyID(StudentUID) And TestMyID(TermID) And TestMyID(FacultyID) Then
    Set Obj = CAMSCreateObject("CAMSPortal.busPortal")
    Call RSToXML(rsStudent, rsXML)
    Result = Obj.GetStudentsInfo(rsXML, strSvrName, strDBName, strError, StudentUID)
    If Result = 0 Then
        Call XMLToRS(rsXML, rsStudent)
        TRSSession("StudentName") = rsStudent.Fields("LastName").Value + ", " + rsStudent.Fields("FirstName").Value + " " + rsStudent.Fields("MiddleName").Value

        TRSSession("StudentUID") = StudentUID
        Result = Obj.IsStudentAdviseeOfFaculty(IsAdvisee, FacultyID, StudentUID, TermID, strSvrName, strDBName, strError)
        TRSSession("IsAdvisee") = CStr(IsAdvisee)
    End If
Else
    Result = -1
End If
If Result = 0 Then
    Response.Redirect("ceStudentOptions.asp?ak=" & accessKey)
Else
    Response.Redirect("ceSelectStudent.asp?ak=" & accessKey)
End If
%>
