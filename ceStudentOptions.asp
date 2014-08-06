<%@  language="VBScript" %>
<%
    Option Explicit
    Dim bCache
    
    Dim rsXML
    Dim Obj
    Dim strError
    Dim Result
    bCache = False

%>
<!--#Include File="inc/Includes.asp"-->
<%
    'Page specific server side code begins here. 
    'TRSSession("sReferer") = "cmCourseManagementOptions.asp"
    Dim picpath
    Dim strStudentName
    Dim StudentAddressType
    Dim rsAddress
    Dim msg
    Dim Address1
    Dim Address2
    Dim City
    Dim State
    Dim ZipCode
    Dim Phone1
    Dim Email1
    Dim StudentUID
    Dim rsFerpa
    Dim rowClass
    Dim i
    Dim bShowTranscript
    Dim verifyAccess

    verifyAccess = VerifyAccessFn(False, Request.Form("accessKey"), Request.QueryString("ak"))

    PageTitle = "Student Options"
    
    bShowTranscript = False
    Result = 0
    StudentUID = TRSSession("StudentUID")
    StudentAddressType = TRSSession("StudentAddressType")
    
    strStudentName = "Student Information in not available."
    Address1 = "No Active Address for Address Type of " & StudentAddressType
    Address2 = ""
    City = ""
    State = ""
    ZipCode = ""
    Phone1 = ""
    Email1 = ""
    picpath = ""    
    
    If TestMyID(StudentUID) And (verifyAccess) Then
        Set Obj = CAMSCreateObject("CAMSPortal.busPortal")
        Call RSToXML(rsAddress, rsXML)
        
        Result = Obj.GetStudentActiveAddressOfType(rsXML, StudentUID, StudentAddressType, strSvrName, strDBName, strError)
        
        If Result = 0 Then
            Call XMLToRS(rsXML, rsAddress)

            If IsObject(rsAddress) Then
                rsAddress.Filter = "AddressType='" & StudentAddressType & "' and ActiveFlag = 'Yes'"
                If rsAddress.RecordCount > 0 Then
                    strStudentName = Server.HTMLEncode(TRSSession("StudentName"))
                    Address1 = Trim(rsAddress.Fields("Address1").Value)
                    If Not IsNull(rsAddress.Fields("Address2").Value) And Len(rsAddress.Fields("Address2").Value) > 0 Then
                        Address2 = Trim(rsAddress.Fields("Address2").Value)
                    End If
                    City = rsAddress.Fields("city").Value
                    State = rsAddress.Fields("State").Value
                    ZipCode = rsAddress.Fields("ZipCode").Value
                    Phone1 = Trim(rsAddress.Fields("Phone1"))
                    
                    
                   
                    For i = 1 to Len(TRSSession("StudentEmailAddress"))
                        If Len(rsAddress.Fields("Email" & Mid(TRSSession("StudentEmailAddress"), i, 1)).Value) > 0 Then
                            Email1 = Email1 & rsAddress.Fields("Email" & Mid(TRSSession("StudentEmailAddress"), i, 1)).Value & ";"
                        End If
                    Next
                    
                    If Right(Email1, 1) = ";" Then
                        Email1 = Left(Email1, Len(Email1) - 1)
                    End If
                    
                    
                    picpath = Application("BasePicURL") & Application("PicPath") & CStr(rsAddress.Fields("StudentUID").Value) & ".jpg"
                Else
                    msg = "<p>The selected student does not have the following address type setup.<br />Address Type: <b>"+ StudentAddressType + "</b><br />Please login to CAMS and add the missing address type.</p>"
                End If        
            End If
            
            rsXML = ""
            
            Call RSToXML(rsFerpa, rsXML)
            Result = Obj.ReadStudentFERPARestrictions(rsXML, StudentUID, strSvrName, strDBName) 
            If Result = 0 Then
                Call XMLToRS(rsXML, rsFerpa)
            End If
            bShowTranscript = AllowTranscriptDisplay()
        End If
    Else
        Result = -1
    End If
    ' Determine if user has access to this area

    Function AllowTranscriptDisplay()
        Dim retVal
        Dim rs
        Dim Obj
        
        retVal = False
        
        '1. If Advisor, always show the transcript.
        If LCase(TRSSession("IsAdvisee")) = LCase("true") Then
            retVal = True
        Else
        '2. If Faculty - Limit Transcript to Advisors Only?
            Set Obj = CAMSCreateObject("CAMSPortal.busPortal")
            Call RSToXML(rs, rsXML)
            Result = Obj.ReadCamsPortalConfig(rsXML, strSvrName, strDBName, strError)
            If Result = 0 Then
                Call XMLToRS(rsXML, rs)                
                If rs.Fields("RestrictTranscriptsToAdvisees").Value <> True Then
                    retVal = True
                End If
            End If        
            ' Faculty Portal Defaults to always show, but uncomment the line below to make it adhere to the configuration
            'retVal = (rsPortalConfig.Fields("ShowTranscript").Value = True)
            Set rs = Nothing
            Set Obj = Nothing
        End If
        
        AllowTranscriptDisplay = retVal
    End Function
    
        
 
   
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en">
<head>
    <meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
    <title>
        <%=PageTitle%>
    </title>
    <!--#Include File="styles/style.inc"-->
    <!--#Include File="scripts/jscript.inc"-->

    <script language="javascript" type="text/javascript">
        //<![CDATA[
        var a1 = '';
        getLocalKey();
        function showOption(token)
        {
            switch (token)
            {
            <% 
                If (bShowTranscript) Then
            %>
                case 1:
                    window.location = 'cePortalTranscript.asp?ak=' + a1;
                    break;
            <% 
                End If
            %>
                case 2:
                    window.location = 'cePortalMatrixSchedule.asp?ak=' + a1;
                    break
                case 3:
                    window.location = 'saudit.asp?ak=' + a1;
                    break;
            <%
                If LCase(TRSSession("IsAdvisee")) = "true" Then
            %>
                case 4:
                    window.location = 'cePortalGradeReport.asp?ak=' + a1;
                    break;
            <% 
                End If
            %>
                case 5:
                    window.location = 'ceStudentRisk.asp?ak=' + a1;
                    break;
                    
                default:
                    break;
            }
        }

        // ADD  - WILLIAM CAREY UNIVERSITY
        function registerStudent(UID)
        {
            document.getElementById('ak').value = a1;
            document.getElementById('registerUID').value = UID;
            document.forms['registerStudentUID'].action = 'ceFacultyRegistrationParams.asp?ak=' + a1;
            document.forms['registerStudentUID'].submit();
        }
        // END ADD

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
                            <!--#Include File = "inc/ceStudentName.asp"-->
                        </div>
                        <!--#Include File = "inc/PortalCrumbs.asp" -->
                        <%
                            If Result = 0 Then 
                        %>
                        <div class="Page_Content">
                           
                                
                                <div class="moduleContainer">
                                    <div class="course_management_options">
                                        <h1>Student Actions</h1>
                                        <div class="course_management_links">
                                            <ul>
                                                <% 
                                                    If (bShowTranscript) Then
                                                        Response.Write("<li><a  href=""#"" onclick=""javascript:showOption(1);"" title=""View Student Transcript"">Transcript</a></li>")
                                                    End If    
                                                %>
                                                <li><a onclick="javascript:showOption(2);" href="#" title="View Student Schedule">Schedule</a></li>
                                                <li><a onclick="javascript:showOption(3);" href="#" title="View Student Degree Audit">Degree Audit</a></li>
                                                <%
                                                    If LCase(TRSSession("IsAdvisee")) = "true" Then
                                                        Response.Write("<li><a href=""#"" onclick=""javascript:showOption(4);"" title=""View Advisee Grade Report"">Grade Report</a></li>")
                                                    End If
                                                %>
                                                <li><a onclick="javascript:showOption(5);" href="#" title="View Student Risk">Student Risk</a></li>
                                                <!-- ADD - WILLIAM CAREY UNIVERSITY -->
                                                <li><a onclick="javascript:registerStudent(<%=StudentUID %>);" href="#" title="Register Student">Register Student</a></li>
                                                <form id="registerStudentUID" action="#" method="post">
                                                    <input id="registerUID" name="UID" type="hidden" />
                                                    <input type="hidden" id="ak" name="accessKey" />
                                                </form>
                                                <!-- END ADD -->
                                            </ul>
                                        </div>
                                    </div>
                                
                                
                               
                                    <div class="course_management_options">
                                        <h1>Student Information</h1>
                                        <div class="course_management_links">
                                            <address>
                                                <% 
                                                    Response.Write(Server.HTMLEncode(Address1))
                                                    If Len(Address2) > 0 Then
                                                        Response.Write("<br />" & Server.HTMLEncode(Address2))
                                                    End If
                                                    Response.Write ("<br />" & Server.HTMLEncode(City) & ", " & Server.HTMLEncode(State) & " " & Server.HTMLEncode(ZipCode))
                                                    If Len(Phone1) > 0 Then
                                                        Response.Write("<br />" & Server.HTMLEncode(Phone1))
                                                    End If
                                                %>
                                                <br />
                                                <a id="A1" href="mailto:<%=Email1 %>" title="Email for <%=Server.HTMLEncode(strStudentName) %>">
                                                    <%=Email1%>
                                                </a>
                                            </address>
                                        </div>
                                    </div>
                                
                                
                                
                                    <div class="course_management_options">
                                        <h1>
                                            <%=strStudentName%>
                                        </h1>
                                        <div class="course_management_links">
                                            <img class="Portal_Img_Action" alt="Student Photo" src="GetPictureImage.asp?ak=<%=accessKey %>&amp;i=<%=StudentUID %>" width="100%" />
                                        </div>
                                    </div>
                               </div><!-- end moduleContainer -->
                            
                            <table width="100%" summary="FERPA Restrictions">
                                <caption class="Portal_Table_Caption">
                                    FERPA Restrictions</caption>
                                <thead>
                                    <tr style="background:none;">
                                        <th abbr="Category">
                                            FERPA Item
                                        </th>
                                        <th abbr="Allow">
                                            Provide Info
                                        </th>
                                        <th abbr="Provide">
                                            To Whom
                                        </th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <% 
                                        If rsFerpa.RecordCount > 0 Then
                                            rsFerpa.MoveFirst
                                            rowClass = "altrow"
                                            Do
                                                If rowClass = "row" Then
                                                    rowClass = "altrow"
                                                Else
                                                    rowClass = "row"
                                                End If
                                    %>
                                    <tr class="<%=rowClass %>">
                                        <td style="text-align: left">
                                            <%=Server.HTMLEncode(rsFerpa.Fields("FerpaItem").Value) %>
                                        </td>
                                        <td style="text-align: left">
                                            <%=Server.HTMLEncode(rsFerpa.Fields("AllowDisplay").Value) %>
                                        </td>
                                        <td style="text-align: left">
                                            <%=Server.HTMLEncode(rsFerpa.Fields("RelationCanRecv").Value) %>
                                        </td>
                                    </tr>
                                    <%
                                                rsFerpa.MoveNext
                                            Loop Until rsFerpa.EOF
                                        Else
                                            Response.Write("<tr class=""row""><td style=""text-align: left"">N/A</td><td>&nbsp;</td><td>&nbsp;</td></tr>")
                                        End If
                                    %>
                                </tbody>
                            </table>
                        </div>
                        <!--/Page Specific Content-->
                        <div id="status">
                            <% 
                                If Len(msg) > 0 Then
                                    Response.Write(FormatMessage(msg, "Information"))
                                End If
                            %>
                        </div>
                        <div id="error" style="display: none;">
                            Error:
                        </div>
                        <%
                            Else
                                Response.Write(FormatMessage("", "RetriveError"))
                            End If
                        %>
                    </div>
                </div>
                <!--/Main-->
            </div>
            <!--/Main-->
            <div class="yui-b" id="leftSideBar">
                <!--LeftSide-->
                <!--#Include File = "inc/Profile.asp"-->
                <!--#Include File = "inc/Menu.inc"-->
                <!--#Include File = "inc/LeftSide.inc"-->
                <!--/LeftSide-->
            </div>
            <!--/Body-->
        </div>
        <div id="ft">
            <!--#Include File = "inc/Footer.inc"-->
        </div>
    </div>
</body>
</html>
