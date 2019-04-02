<%
dim username, user

Response.Buffer=True
Response.Expires = -100
 	
username=request.form("username")
userpwd=request.form("userpwd")

Veritabani_Yol=SERVER.MAPPATH("Veritabanim.mdb")
Set Baglanti=Server.CreateObject("Adodb.Connection")
Baglanti.Open "DBQ=" & Veritabani_Yol &   ";Driver={Microsoft Access Driver (*.mdb)}"
Set Rs=Server.CreateObject("Adodb.recordset")

Sorgu="select * from kayitt where Mail = '" & request.form("Mail") & "' and Pass = '" & Request.form("Userpwd") & "'"
    Set grup = Baglanti.Execute(Sorgu) 'ppp
    

		Rs.Open Sorgu, Baglanti, 1, 3
		If RS.BOF And RS.EOF Then
		    Response.Write "Bilgiler onaylanmadi. Yanlis Kullanici Adi veya Sifre."
		Else
			session("UserLoggedIn")=Rs("Ad")
	     	response.redirect("index.asp")
 		End If
 		%>