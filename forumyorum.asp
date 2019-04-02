<%	Set Baglanti = Server.CreateObject("ADODB.Connection")
		Baglanti.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
		Set Tablom = server. CreateObject("ADODB.Recordset")
		Tablom.Open "Forum", Baglanti, 1, 3

		Tablom.AddNew
		Tablom("Tarih") = now()
		Tablom("Kim") =	Session("UserLoggedIn")
		Tablom("Yorum")=request("metin")
		Tablom.Update

	Tablom.close
	set Tablom= Nothing
	Baglanti.close
	set Baglanti= Nothing
	Response.Redirect("forum.asp")
%>