<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="lib.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>K-12: The Changing Mathematics Curriculum: An Annotated Bibliography</title>
<link href="../css/refs.css" rel="stylesheet" type="text/css" media="screen">
<link href="../css/second.css" rel="stylesheet" type="text/css" media="screen">

</head>

<body>
<div id="container">

<div id="header">
<h1>Welcome to The K-12 Mathematics Curriculum Center</h1>
</div>

<div id="navigation">
<ul>
  <li><a href="../default.asp">Home</a></li>
  <li><a href="../about/default.asp">About Us</a></li>
  <li><a href="default.asp">Publications</a></li>
  <li><a href="../research/default.asp">Research</a></li>
  <li><a href="../resources/default.asp">Other Resources </a></li>
  <li><a href="../contact.asp">Contact Us</a></li>
</ul>

<div id="search">
<form name="look" method="GET" action="http://google2.edc.org/search">
<input name="q" type="text" value="Search K-12" size="20">
<br/><br/>
<input type="submit" value="Search Site" name="btng">
<input type="hidden" name="site" value="mcc">
<input type="hidden" name="output" value="xml_no_dtd">
<input type="hidden" name="client" value="edc_general">
<input type="hidden" name="proxystylesheet" value="edc_general">
</form>
</div>

</div>

<div id="maincontent">
 <img src="../Images/pubslogo.gif" alt="Publications" width="70" height="55" class="sectionlogo">
   <h1>Publications</h1>
   <img src="../Images/3squares.gif" width="100" height="35" class ="bigsquares" alt="Squares" title="Squares">
  <h2>Selecting Mathematics Instructional Materials: An Annotated Bibliography</h2>
<%
'Dimension variables
Dim dbConn			'Holds the Database Connection Object
Dim rsRef			'Holds the recordset for the records in the database
Dim rsAuthors		'Holds the recordset for all the authors
Dim refSQL			'Holds the SQL query to query the database
Dim author
Dim authorSQL		'Holds the SQL query for the author list
Dim refURL
Dim authorFound

author = Trim(Request.Form("author"))
%>
  <form action="mbiblionewlist.asp" method="post" id="searchForm" name="searchForm">
  <input name="author" type="text" size="20" maxlength="20" value="<%=author%>">
  <input name="Search" type="submit" value="Search Author Name">
  </form>
<%
'Create an ADO connection object
dbConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath("../data/abdatabase.mdb")

'strSQL = "SELECT RecordNo, reftypeID, pubyear, title, editor, journaltitle, city, publisher, volume, issue, pages FROM refs;"

refSQL = "SELECT * FROM refs"
authorFound = 1

If not author="" Then
	Response.Write "<h3>Search Results</h3>" & vbCRLF
	sAuthorSQL = "SELECT * FROM authors WHERE (authorLast LIKE '" & author & "%');"
	Set rsSearch = getRORS(sAuthorSQL, dbConn)
	numResults = rsSearch.recordcount
	Select Case numResults
		Case 0
			authorFound = 0
			Response.Write "<p style=""color:#CC0000"">No results found. Please try again or <a href=""mbiblionewbrowse.asp"">browse the categories</a>.</p>"
		Case 1
			refSQL = refSQL & " WHERE RecordNo = " & rsSearch("refID") & ";"
		Case Else
			refSQL = refSQL & " WHERE (RecordNo = " & rsSearch("refID") & ")"
			rsSearch.MoveNext
			Do While Not rsSearch.EOF
				refSQL = refSQL & " OR (RecordNo = " & rsSearch("refID") & ")"
				rsSearch.MoveNext
			Loop
			refSQL = refSQL & " ORDER BY pubyear DESC;"
	End Select

If authorFound = 1 Then
'gets a read-only recordset
		Set rsRef = getRORS(refSQL, dbConn)	
'Response.Write "<h2>References List Secured</h2>" & vbCRLF

'This will list all Column headings in the table


		Dim item
	
		Do While Not rsRef.EOF
			Set refType = rsRef("reftypeID")
			Response.Write "<p class=""cite"">" & vbCRLF
'			Response.Write "<a href=""mbiblionewref.asp?refID=" & rsRef("RecordNo") & """>" & rsRef("RecordNo") & "." & "</a>" & vbCRLF
			authorSQL = "SELECT * FROM authors WHERE refID = " & rsRef("RecordNo") & " ORDER BY posID;"
			Set rsAuthors = getRORS(authorSQL, dbConn)
			numAuthors = rsAuthors.recordcount
			If numAuthors = 1 Then
				If isNull(rsAuthors("authorFirst")) Then
					Response.Write rsAuthors("authorLast") & "." & vbCRLF
				Else
					Response.Write rsAuthors("authorLast") & ", " & rsAuthors("authorFirst") & vbCRLF
				End If
			Else
				If numAuthors = 2 Then
					Response.Write rsAuthors("authorLast") & ", " & rsAuthors("authorFirst") & " &" & vbCRLF
					rsAuthors.MoveNext
					Response.Write rsAuthors("authorLast") & ", " & rsAuthors("authorFirst") & vbCRLF
				Else
					For acount = 1 To (numAuthors - 2)
						Response.Write rsAuthors("authorLast") & ", " & rsAuthors("authorFirst") & "," & vbCRLF
						rsAuthors.MoveNext
					Next
					Response.Write rsAuthors("authorLast") & ", " & rsAuthors("authorFirst") & " &" & vbCRLF
					rsAuthors.MoveNext
					Response.Write rsAuthors("authorLast") & ", " & rsAuthors("authorFirst") & vbCRLF
				End If
			End If
			Set rsAuthors = Nothing
			Response.Write "(" & rsRef("pubyear") & ")." & vbCRLF
			Select Case refType
				Case "1"
					Response.Write rsRef("title") & "." & vbCRLF
					Response.Write "<em>" & rsRef("journaltitle") & "," & vbCRLF
					Response.Write rsRef("volume") & "</em>(" & rsRef("issue") & ")," & vbCRLF
					Response.Write rsRef("pages") & "." & vbCRLF
				Case "2"
					Response.Write rsRef("title") & "." & vbCRLF
					Response.Write "In " & rsRef("editor") & " (Eds.)," & vbCRLF
					Response.Write "<em>" & rsRef("journaltitle") & "</em>" & vbCRLF
					Response.Write "(pp. " & rsRef("pages") & ")." & vbCRLF
					Response.Write rsRef("city") & ":" & vbCRLF
					Response.Write rsRef("publisher") & "." & vbCRLF
				Case "3"
					Response.Write "<em>" & rsRef("title") & "</em>." & vbCRLF
					Response.Write rsRef("city") & ":" & vbCRLF
					Response.Write rsRef("publisher") & "." & vbCRLF
				Case "4"
					Response.Write rsRef("title") & "." & vbCRLF
					Response.Write rsRef("city") & ":" & vbCRLF
					Response.Write rsRef("publisher") & "." & vbCRLF
				Case "5"
					Response.Write "<em>" & rsRef("title") & "</em>" & vbCRLF
					Response.Write rsRef("editor") & "(Ed.)." & vbCRLF
					Response.Write rsRef("city") & ":" & vbCRLF
					Response.Write rsRef("publisher") & "." & vbCRLF
				Case Else
					Response.Write "Error: incorrect reference type."
			End Select
			Response.Write "</p>" & vbCRLF
			Response.Write "<p class=""indent"">" & rsRef("abstract") & "</p>" & vbCRLF
			Set refURL = rsRef("url")
			If Not isNull(refURL) Then
				Response.Write "<p class=""indent""><a href=""" & refURL & """ target=""_blank"">" & refURL & "</a></p>" & vbCRLF
			End If
			Set refLegalNote = rsRef("legalnote")
			If Not isNull(refLegalNote) Then
				Response.Write "<p class=""indent"">" & refLegalNote & "</p>" & vbCRLF
			End If
			rsRef.MoveNext
		Loop

'Reset server objects
		rsRef.Close
		Set rsRef = Nothing
	End If
	Set dbConn = Nothing
End If
%>
  </div> 

<div id="footer">
<hr noshade size="1">
<div id="textlinks">
<p><a href="../default.asp">Home</a>  |  <a href="../about/default.asp">About Us</a>  |  <a href="default.asp">Publications</a>  |  <a href="../research/default.asp">Research</a>  |  <a href="../resources/default.asp">Other Resources </a>  |  <a href="../contact.asp">Contact Us</a> <!-- |  <a href="#">Sitemap</a>--></p>
</div>

<div id="nsf">
<p>The K&ndash;12 Mathematics Curriculum Center is funded by the <a href="http://www.nsf.gov">National
    Science Foundation</a> to inform and assist schools and districts as they
    select and implement Standards-based mathematics curricula.</p>
</div>

<div id="edc">
<p>Site hosted by<br>
  <a href="http://main.edc.org">Education Development Center, Inc.</a> <br>
  <a href="http://main.edc.org/info/legal.asp">&copy; 2009  All rights reserved.</a></p>
</div>

</div>

</div>
</body>
</html>
