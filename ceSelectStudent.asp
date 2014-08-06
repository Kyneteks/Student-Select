<%@  language="VBScript" %>
<%
    Option Explicit
    Dim bCache
    bCache = False
%>
<!--#Include File="inc/Includes.asp"-->
<%
    'Page specific server side code begins here.

    Response.Buffer = True

    TRSSession("sReferer") = "ceStudentOptions.asp"

    Dim rsFaculty
    Dim rsXML
    Dim strError
    Dim busPortal
    Dim Result
    Dim IDName
    Dim PageCount
    Dim strUserPass
    Dim rowClass
    Dim msg
    Dim busRegReports
    Dim rsCriteria
    Dim rsCourses
    Dim rsCoursesXML
    Dim rsStudents
    Dim rsStudentsXML
    Dim rsAttendance
    Dim rsAttendanceXML
    Dim rsSchedule
    Dim rsScheduleXML
    Dim rsStudentAddresses
    Dim rsStudentAddressesXML
    Dim rsReportParams
    Dim rsReportParamsXML
    Dim strDisplayCourseID

    Dim stuCount
    Dim ShowSROfferID
    
    Dim StudentName
    Dim ShowPreferredName
    Dim showStudent
    
    Dim ShowPhotoCkd
    Dim ShowWithdrawnCkd
    Dim strFilter
    
    Dim arStudents()
    Dim cntStudent
    Dim verifyAccess

    verifyAccess = VerifyAccessFn(True, Request.Form("accessKey"), Request.QueryString("ak"))
    
    cntStudent = -1
        
    ShowPreferredName = True
    strUserPass = Request.Form("password")

    ShowPhotoCkd = Request.Form("hShowPhoto")
    ShowWithdrawnCkd = Request.Form("hShowWithdrawn")

    If Len(ShowWithdrawnCkd) = 0 Then
        ShowWithdrawnCkd = bIncludeWithdrawn
    Else
        bIncludeWithdrawn = ShowWithdrawnCkd
        TRSSession("bIncludeWithdrawn") = bIncludeWithdrawn
    End If

    If TestMyID(FacultyID) And (verifyAccess) Then
        Set busPortal = CAMSCreateObject("CAMSPortal.busPortal")
        Call RSToXML(rsFaculty, rsXML)
        
        Result = busPortal.ReadFacultySchedule(FacultyID, TextTerm, rsXML, strUserName, strSvrName, strDBName, strError )
        
        If Result = 0 Then
            Call XMLToRS(rsXML, rsFaculty)
        End If
    Else
        Result = -1
    End If    
    ShowSROfferID = Request.Form("showOfferID")
    If Len(ShowSROfferID) = 0 Then
        ShowSROfferID = "None"
    ElseIf LCase(ShowSROfferID) = LCase("All") Then
    ElseIf Not TestMyID(ShowSROfferID) Then
        ShowSROfferID = "None"
    End If

    PageTitle = "My Students"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en">
<head>
    <meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
    <title>
        <%=PageTitle %>
    </title>
    <!--#Include File="styles/style.inc"-->
    <!--#Include File="scripts/jscript.inc"-->

    <script language="javascript" type="text/javascript">
        //<![CDATA[
        var a1 = '';
        getLocalKey();
        function OnChangeCourse()
        {
            showWaitPanel();
            document.getElementById('idShowOffer').value = document.getElementById('idCourseName').value;
            document.getElementById('ak').value = a1;
            document.forms['Navigation'].action = 'ceSelectStudent.asp?ak=' + a1;
            document.forms['Navigation'].submit();
        }
        function ShowPhoto(cb)
        {
            document.getElementById('hShowPhoto').value = cb.checked;
        }
        function ShowWithdrawn(cb)
        {
            showWaitPanel();
            document.getElementById('hShowWithdrawn').value = cb.checked;
            document.getElementById('idShowOffer').value = document.getElementById('idCourseName').value;
            document.getElementById('ak').value = a1;
            document.forms['Navigation'].action = 'ceSelectStudent.asp?ak=' + a1;
            document.forms['Navigation'].submit();
        }
        function SelectMyStudent(StudentUID)
        {
            showWaitPanel();
            document.getElementById('txtStudentUID').value = StudentUID;
            document.getElementById('ak').value = a1;
            document.forms['Navigation'].action = 'setSessionObjects.asp?ak=' + a1;
            document.forms['Navigation'].submit();
		}
		// ADD - WILLIAM CAREY UNIVERSITY
        function SelectStudent() 
        {
            var valSelectStudentUID = frmSelectStudent.txtSelectStudentUID;
            if (valSelectStudentUID.value != "")
            {
                document.getElementById('txtStudentUID').value = valSelectStudentUID.value;
                document.getElementById('ak').value = a1;
				document.forms['Navigation'].action = 'setSessionObjects.asp';
                document.forms['Navigation'].submit();
            }
		// END ADD	
        }
        
        //]]>
    </script>

</head>
<body>
    <div id="doc3" class="yui-t3">
        <div id="hd">
            <!--#Include File="inc/Header.inc"-->
            <!--#Include File="inc/Topmenu.asp"-->
        </div>
        <div id="bd">
            <!--Body-->
            <div id="yui-main">
                <!--Main-->
                <div class="yui-b">
                    <div class="yui-g" id="mainBody">
                        <!--Page Specific Content-->
                        <div class="Page_Logo">
                            <!--#Include File="inc/PageTitle.asp"-->
                        </div>
                        <!-- ADD - WILLIAM CAREY UNIVERSITY --> 
						<div class="pageOptions" style="line-height:1.7em;">
						<% If Request.QueryString("Failed") <> "" Then %>
                            <div class="Portal_Message_Blue">
                                <table>
                                    <tr>
                                        <td>
                                            <img src="../Student/images/Portal_Information_Error.gif" />
                                        </td>
                                        <td style="font-size: large">
                                            No student could be located with the id number you entered: <span style="font-weight: bold"> <%Response.Write Request.QueryString("Failed") %></span>.<br />
                                        </td>
                                    </tr>                                
                                </table>
                            </div>
                            <br />
                        <% End If %>
                        <form action="#" method="post" id="frmSelectStudent">
                            <table style="width: 100%">
                                <tr>
                                    <td>
                                        <span style="font-size: medium">
                                            If you would like to view information for a student that is not your advisee or currently in one of your courses, please enter their student id number below and click Select Student.<br /><br />
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <span style="font-size: medium; font-weight: bold">Student UID:</span> <input type="text" id="txtSelectStudentUID" name="txtSelectStudentUID" /> &nbsp&nbsp <a onclick="javascript:SelectStudent();" href="#"><span style="font-size: medium; font-weight: bold">Select Student</span></a>
                                    </td>
                                </tr>
                            </table>                                                                                                    
                        </form>
                        <br />
						<!-- END ADD -->
                            <label for="idCourseName">Select List:&nbsp;</label>
                                    <select title="Select course or all courses" name="CourseName" id="idCourseName" onchange="OnChangeCourse();">
                                        <option value="None" <%If ShowSROfferID = "None" Then Response.Write("selected=""selected""") End If  %>>Advisee List</option>
                                        <%
                                            PageCount = 0
                                            Dim isSelected
                                            
                                            Do While Not rsFaculty.EOF
                                                isSelected = ""
                                                PageCount = PageCount + 1
                                                strDisplayCourseID = rsFaculty.Fields("Department").Value & rsFaculty.Fields("Course").Value & rsFaculty.Fields("CourseType").Value & rsFaculty.Fields("Section").Value
                                                
                                                If CStr(rsFaculty.Fields("SROfferID").Value) = CStr(ShowSROfferID) Then
                                                    Response.Write("<option value=""" & CStr(rsFaculty.Fields("SROfferID").Value) & """ selected=""selected"">" & Server.HTMLEncode(strDisplayCourseID) & " - " & Server.HTMLEncode(rsFaculty.Fields("CourseName").Value) & "</option>")
                                                Else
                                                    Response.Write("<option value=""" & CStr(rsFaculty.Fields("SROfferID").Value) & """>" & Server.HTMLEncode(strDisplayCourseID) & " - " & Server.HTMLEncode(rsFaculty.Fields("CourseName").Value) & "</option>")
                                                End If

                                                rsFaculty.MoveNext
                                            Loop
                                        %>
                                        <option value="All" <%If ShowSROfferID = "All" Then Response.Write("selected=""selected""") End If  %>>All Courses for Term</option>
                                    </select>
                                    
                                
                                    <% 
                                        If ShowSROfferID <> "None" Then
                                    %>
                                    Show Withdrawn Students&nbsp;<input id="showWithdrawn" type="checkbox" title="Check to show withdrawn students." onclick="javascript:ShowWithdrawn(this);" <%If LCase(ShowWithdrawnCkd) = "true" Then Response.Write("checked=""checked""") End If%> />
                                    <% 
                                        Else
                                            Response.Write("&nbsp;")
                                        End If
                                    %>
                                
                                
                        <form id="Navigation" action="#" method="post">
                            <input type="hidden" id="txtStudentUID" name="txtStudentUID" />
                            <input type="hidden" id="idShowOffer" name="showOfferID" />
                            <input type="hidden" id="hShowPhoto" name="hShowPhoto" value="<%=ShowPhotoCkd%>" />
                            <input type="hidden" id="hShowWithdrawn" name="hShowWithdrawn" value="<%=ShowWithdrawnCkd%>" />
                            <input type="hidden" id="ak" name="accessKey" />
                        </form>
                        </div><!-- end Form_Container headerOnly -->
                        
                        <div class="Page_Content">
                        <form id="forPhotos" action="#">
                            <div id="SROfferDIV0">
                                <table class="Portal_Group_Table center striped" summary="My Students">
                                    <thead>
                                        <tr style="background:none;">
                                            <th>
                                                Select
                                            </th>
                                            <th>
                                                Student ID
                                            </th>
                                            <th>
                                                Name
                                            </th>
                                            <th>
                                                Photo
                                            </th>
                                        </tr>
                                    </thead>
                                    <% 
                                        If ShowSROfferID = "None" Then
                                            rsFaculty.Filter = "SROfferID = -1"
                                            Set busPortal = CAMSCreateObject("CAMSPortal.dbFacultyPortal")
                                            Call RSToXML(rsStudents, rsXML)
                                            Result = busPortal.GetFacultyAdvisorList(rsXML, FacultyID, TermID, strSvrName, strDBName, strError)
                                            If Result <> 0 Then
                                    %>
                                    <tr class="<%=rowClass%>">
                                        <td colspan="4">
                                            <%Response.Write(FormatMessage("", "RetrieveError")) %>
                                        </td>
                                    </tr>
                                    <%
                                            Else
                                                Call XMLToRS(rsXML, rsStudents)
                                                Set busPortal = Nothing
                                                stuCount = 0
                                                'cntStudent = -1
                                                rowClass = "row"
                                                Do While Not rsStudents.EOF
                                                    stuCount = stuCount + 1
                                                    IDName = Server.URLEncode("ID: " & rsStudents.Fields("StudentID").Value & " " & rsStudents.Fields("StudentName").Value)
                                                    StudentName = rsStudents.Fields("StudentName").Value
                                                    If rowClass = "row" Then
                                                        rowClass = "altRow"
                                                    Else
                                                        rowClass = "row"
                                                    End If
                                                    
                                                    cntStudent = cntStudent + 1
                                                    ReDim Preserve arStudents(cntStudent)
                                                    arStudents(cntStudent) = rsStudents.Fields("StudentUID").Value
                                    %>
                                    <tr>
                                        <td style="width: 10%;">
                                            <a class="button" onclick="javascript:SelectMyStudent(<%=cntStudent%>);" href="#">Select</a>
                                        </td>
                                        <td style="width: 25%;">
                                            <%
                                                If Not IsNull(rsStudents.Fields("StudentID").Value) Then
                                                    Response.Write(rsStudents.Fields("StudentID").Value)
                                                Else
                                                    Response.Write("&nbsp;")
                                                End If
                                                                        
                                            %>
                                        </td>
                                        <td style="width: 65%;text-align:left;">
                                            &nbsp;<%=stuCount%>.&nbsp;
                                            <% 
                                                If Not IsNull(StudentName) Then
                                                    Response.Write(Server.HTMLEncode(StudentName))
                                                Else
                                                    Response.Write("&nbsp;")
                                                End If
                                            %>
                                        </td>
                                        <td style="width: 5%; text-align: center">
                                            <a href="#" onclick="openpopup('ceFacultyDispPic.asp??ak=<%=accessKey %>&amp;pt=studentAdvisee&amp;sn=<%=Server.URLEncode(rsStudents.Fields("StudentName").Value)%>&amp;s2=<%=CStr(cntStudent)%>');">
                                                <img style="width:80%;" class="Portal_Img_Action" alt="Student Picture" src="images/icon_photo.png" />
                                            </a>
                                        </td>
                                    </tr>
                                    <%
                                                    rsStudents.MoveNext    
                                                Loop
                                                If cntStudent >= 0 Then
                                                    TRSSession("Advisees") = Serialize(arStudents, cntStudent)
                                                End If
                                            End If
                                        Else
                                            rsFaculty.Sort = "Department, Course, CourseType, Section, SROfferID"
                                            SROfferID = -1
                                            If Not IsNull(Request.QueryString("srofferid") ) Then
                                                SROfferID = Request.QueryString("srofferid")
                                            End If
                                            rsXML = ""
                        
                                            Set busRegReports = CAMSCreateObject("CAMSPortal.busGeneral")
                                            Call RSToXML(rsCriteria, rsXML)
                                            
                                            Result = busRegReports.GetCriteriaRecord(rsXML, strSvrName, strDBName, strError)
                                            If Result = 0 Then
                                                Call XMLToRS(rsXML, rsCriteria)
                                            Else
                                                Response.Write(FormatMessage(Server.HTMLEncode(strError), "Warning"))
                                            End If
                                            
                                            rsCriteria.Fields("ReportType") = 3
                                            rsCriteria.Fields("count") = 2
                                            rsCriteria.Fields("Name1") = 1
                                            rsCriteria.Fields("value1") = TRSSession("TermID")
                                            rsCriteria.Fields("Name2") = 30
                                            rsCriteria.Fields("value2") = TRSSession("FacultyID")
                                            rsCriteria.Fields("Username") = TRSSession("FacultyID")
                                            rsCriteria.Update
                                            rsCriteria.MoveFirst

                                            strUserName = TRSSession("FacultyID")
                                            
                                            Call RSToXML(rsCriteria, rsXML)
                                            Call RSToXML(rsCourses, rsCoursesXML)
                                            Call RSToXML(rsStudents, rsStudentsXML)
                                            Call RSToXML(rsAttendance, rsAttendanceXML)
                                            Call RSToXML(rsSchedule, rsScheduleXML)
                                            Call RSToXML(rsStudentAddresses, rsStudentAddressesXML)
                                            Call RSToXML(rsReportParams, rsReportParamsXML)

                                            Result = busRegReports.PrintOfferingRoster(rsXML, rsCoursesXML, rsStudentsXML, rsAttendanceXML, _
                                                        rsScheduleXML, rsStudentAddressesXML, rsReportParamsXML, strUserName, strSvrName, strDBName, strError)    
                                                
                                            If Result = 0 Then
                                                Call XMLToRS(rsXML, rsCriteria)
                                                Call XMLToRS(rsCoursesXML, rsCourses)
                                                Call XMLToRS(rsStudentsXML, rsStudents)
                                                Call XMLToRS(rsAttendanceXML, rsAttendance)
                                                Call XMLToRS(rsScheduleXML, rsSchedule)
                                                Call XMLToRS(rsStudentAddressesXML, rsStudentAddresses)
                                                Call XMLToRS(rsReportParamsXML, rsReportParams)
                                                
                                                rsStudents.Sort = "Term, Department, CourseID, CourseType, Section, SROfferID, StudentName"
                                                PageCount = 0
                                            Else
                                                Response.Write(FormatMessage("", "RetrieveError"))
                                            End If                                        

                                            If ShowSROfferID = "All" Then
                                                rsFaculty.Filter = 0
                                            Else
                                                rsFaculty.Filter = "SROfferID = " & ShowSROfferID
                                            End If
                                            Do While Not rsFaculty.EOF
                                                PageCount = PageCount + 1
                                                
                                                rowClass = "Row"
                                                rsStudents.Filter = 0
                                                If rsStudents.RecordCount > 0 Then
                                                    rsStudents.MoveFirst
                                                    strFilter = "SROfferID = " & rsFaculty.Fields("SROfferID").Value  
                                                    If Not ShowWithdrawnCkd Then
                                                        strFilter = strFilter & CreateWithdrawnFilterList(WithdrawnGradesList, "Grade", False)
                                                    End If
                                                    rsStudents.Filter = strFilter
                                                    rsStudents.sort = "StudentName ASC"
                                                End If
                                    %>
                                    <tr class="Portal_Table_Caption">
                                        <td colspan="4">
                                            <%=Server.HTMLEncode(rsFaculty.Fields("Department").Value & rsFaculty.Fields("Course").Value & rsFaculty.Fields("CourseType").Value & rsFaculty.Fields("Section").Value & " - " & rsFaculty.Fields("CourseName").Value) %>
                                        </td>
                                    </tr>
                                    <%
                                                If rsStudents.BOF And rsStudents.EOF Then
                                                    rowClass = "altRow"
                                    %>
                                    <tr>
                                        <td colspan="4">
                                            <%
                                                            msg = FormatMessage("There are no students enrolled for this course for this term.", "Information")
                                                            Response.Write msg
                                            %>
                                        </td>
                                    </tr>
                                    <%
                                                Else
                                                    stuCount = 0
                                                    'cntStudent = -1
                                                    rowClass = "row"
                                                    Do While Not rsStudents.EOF
                                                        showStudent = True
                                                        If (rsStudents.Fields("VarCredits").Value = "Yes") And (rsStudents.Fields("IndStdyFacultyID").Value <> 0) And (rsStudents.Fields("IndStdyFacultyID").Value <> rsFaculty.Fields("FacultyID").Value) Then
                                                            showStudent = False
                                                        End If
                                                        
                                                        If showStudent Then
                                                            stuCount = stuCount + 1
                                                            cntStudent = cntStudent + 1
                                                            ReDim Preserve arStudents(cntStudent)
                                                            arStudents(cntStudent) = rsStudents.Fields("StudentUID").Value
                                                            IDName = Server.URLEncode("ID: " & rsStudents.Fields("StudentID").Value & " " & rsStudents.Fields("StudentName").Value)
                                                            StudentName = rsStudents.Fields("StudentName").Value
                                                            If ShowPreferredName Then
                                                                If Len(Trim(rsStudents.Fields("PreferredName").Value)) <> 0 Then
                                                                    StudentName = StudentName & " (" & rsStudents.Fields("PreferredName").Value & ")"
                                                                End If
                                                            End If
                                                            If rowClass = "row" Then
                                                                rowClass = "altRow"
                                                            Else
                                                                rowClass = "row"
                                                            End If
                                                            If StudentWithdrawn(WithdrawnGradesList, rsStudents.Fields("Grade").Value) Then
                                                                 StudentName = StudentName & " - Withdrawn"
                                                            End If
                                    %>
                                    <tr>
                                        <td style="width: 10%;">
                                            <a class="button" onclick="javascript:SelectMyStudent(<%=cntStudent%>);" href="#">Select</a>
                                        </td>
                                        <td style="width: 15%;">
                                            <%
                                                                        If Not IsNull(rsStudents.Fields("StudentID").Value) Then
                                                                            Response.Write(rsStudents.Fields("StudentID").Value)
                                                                        Else
                                                                            Response.Write("&nbsp;")
                                                                        End If
                                            %>
                                        </td>
                                        <td style="width: 75%">
                                            &nbsp;<%=stuCount%>.&nbsp;
                                            <% 
                                                                        If Not IsNull(StudentName) Then
                                                                            Response.Write(Server.HTMLEncode(StudentName))
                                                                        Else
                                            %>
                                            &nbsp;
                                            <%
                                                                        End If
                                            %>
                                        </td>
                                        <td style="width: 5%; text-align: center">
                                            <a href="#" title="View Photo for <%=StudentName %>" onclick="openpopup('ceFacultyDispPic.asp?ak=<%=accessKey %>&amp;pt=studentAdvisee&amp;s2=<%=CStr(cntStudent)%>');">
                                                <img style="width:80%;" class="Portal_Img_Action" alt="Photo" src="images/icon_photo.png" title="View Photo for <%=StudentName %>" />
                                            </a>
                                        </td>
                                    </tr>
                                    <%
                                                            End If
                                                            rsStudents.MoveNext    
                                                        Loop
                                                        If cntStudent >= 0 Then
                                                            TRSSession("Advisees") = Serialize(arStudents, cntStudent)
                                                        End If

                                                        Response.Flush
                                                        
                                                    End If
                                                    rsFaculty.MoveNext    
                                                Loop
                                        End If
                                    %>
                                </table>
                                </div>
                            </div>
                        </form>
                        <!--/Page Specific Content-->
                        <div id="status" style="display: none;">
                            Status:
                        </div>
                        <div id="error" style="display: none;">
                            Error:
                        </div>
                    </div>
                </div>
                <!--Main-->
            </div>
            <div class="yui-b" id="leftSideBar">
                <!--LeftSide-->
                <!--#Include File = "inc/Profile.asp"-->
                <!--#Include File = "inc/Menu.inc"-->
                <!--#Include File = "inc/LeftSide.inc"-->
                <!--/LeftSide-->
            </div>
            <!--Body-->
        </div>
        <div id="ft">
            <!--#Include File = "inc/Footer.inc"-->
        </div>
    </div>
</body>
</html>
