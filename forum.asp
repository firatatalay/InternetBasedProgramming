<% 
	if Session("UserLoggedIn") = "" then
		Response.redirect("giris.htm")
	else
%>
<html>

<head>
		<meta charset="utf-8">
		<title>Fırat ATALAY</title>
	<link rel="stylesheet" type="text/css" href="üstmenü.css" />
	<style>
		a{ 
			color:white;
			font-face:tahoma;
		}
	
	</style>
	
</head>

<body bgcolor="#EAEAEA" >
	<div class=header> 
	<br>
	<br>
	<br>
	<div class=logo style="float:right;"> <img src="resim/logo.png" width="600" height="130" alt="dededassa" </div>
	</div>
	
	
	
		</div>
		<div class=üstmenü>
		<ul>
			<li> <a href="index.asp"><img src="resim/anasayfa.png"></a> </li>
			<li> <a href="kisiler.asp">Kisiler</a> </li>
			<li> <a href="fotograflar.asp">Fotoğraflar</a> </li>
			<li> <a href="videolar.asp">Videolar </a> </li>
			<li> <a href="kimnerede.asp">Kim,Nerede, Ne Yapıyor? </a> </li>
			<li> <a href="harita.asp?x=3375&y=4445&zoom=1100"> Harita </a> </li>
			<li> <a href="forum.asp">Forum</a> </li>
			<li> <a href="kayit.htm"> Kayıt </a> </li>
			<li> <a href="giris.htm"> Panele Giriş</a> </li>
			
		</ul>
		<div style="float:right;">
		<ul>
		<li> <a href="cikis.asp"> <img src="resim/logout.png" /></a> </li>
		<li> <a href="https://www.instagram.com/firatatalay34/"> <img src="resim/instagram.png" /></a> </li>
		<li> <a href="https://www.facebook.com/firatatalay34"> <img src="resim/facebook.png" /></a> </li>
		
		</ul>
		
		</div>
	
		</div>
	<div class=main style="height:1000;">
		<center>
		
		<br> 
		<h4> Yorum yapmanız için <a color="black" href="kayit.htm"> üye ol</a>manız gerekmektedir. Eğer üyeliğiniz varsa <a href="giris.htm"> giriş yap</a>ınız.<h5>
		<br>
		
			<table>
			<th>Bize Bir Yorum Bırakın ! </th>
				
				<tr>
					<form action="forumyorum.asp" method="POST" ><textarea name="metin" rows="15" cols="100">Yorum Yap</textarea>
				</tr>			
			</table>
			
			<br>
			<input style="width:150px;"  type="submit" value="  Yorumu Gönder  " /></form>
			
			
			<br>
			<br>
			<br>
 
		<table  height="100" width="720" border="" cellspacing="1">
		
					<tr> 
					<td width="150"> <b> Tarih/Saat </b>
					</td>
					<td><b> Kullanıcı Adı </b>
					</td>
					<td width="400"> <b> Yorum </b>
					</td>
					</tr>
			<%
				Set Baglanti = Server.CreateObject("ADODB.Connection")
				Baglanti.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("Veritabanim.mdb"))
				sql="select * from Forum;"
				Set Tablom = Baglanti.Execute(sql)

set Rs=server.createobject("adodb.recordset")
Rs.open sql,baglanti,1,3

If Not Rs.EOF Then
Rs.PageSize= 5
If Request.QueryString("s") <>"" Then
'Bulunduğumuz sayfayı bu değişkenin değeri olarak atayalım

Sayfa = CInt(Request.QueryString("s"))
Else
'Değilse başlangıç sayfa numaramızı 1 olarak atayalım
Sayfa = 1
End If 've Kayıtsetimize hangi sayfada bulunduğumuzu söyleyelim.

Rs.AbsolutePage = Sayfa
 
i=0
'Kayıtsetimizi bir sayfada gösterilecek kayıt sayısı adedince döndürelim.
Do While Not Rs.EOF And i<Rs.PageSize
%>


<tr><td><%=rs("Tarih")%></td>	<td><%=rs("Kim")%></td> 	<td><%=rs("Yorum")%></td>  </tr>

<%
i=i+1
Rs.MoveNext
Loop
%>
<div>Sayfa: 
<a href="?s=<% if sayfa >1 then response.write sayfa-1 else response.write "1" end if%>" title="Onceki sayfa"><<<</a>
<%
If Rs.PageCount > 0 Then
For s=1 To Rs.PageCount
	Response.Write "<a href=?s=" & s & ">" & s & "</a>"
Next
End If
%>
<a href="?s=<% if sayfa < Rs.pagecount then response.write sayfa+1 else response.write Rs.pagecount end if%>" title="Sonraki sayfa">>>></a>
</div>
<%else
 response.write "There is no message"
End If
%>
		</table>
		<%
			Baglanti.Close
			Set Tablom= Nothing
			Set Baglanti= Nothing
		%>

	
	
		</center>
	
	</div>
	
	<br>
	<br>
	<br>
	<br>
	
	
	
	<div class="referanslarım"> 
		  <center>
		  <h3></h3>
		  <br>
		
		 <table width="1000" height  cellspacing="5" cellpadding="5">
		 <tr>
		 <td><a href="http://gezginkisi.com/"><img src="https://i1.wp.com/gezginkisi.com/wp-content/uploads/2017/08/logo-5.png?fit=200%2C90"/></a></td>
		 <td> <img src="resim/ref.png" /> </td>
		 <td> <img src="resim/ref.png" /> </td>
		 <td> <img src="resim/ref.png" /> </td>
		 <td> <img src="resim/ref.png" /> </td>
		 </tr>
		 </table>
		 </center>

	</div>
	
	
	

	<div class=footer>
		<ul>	
		<li> <a href="#"> Gizlilik Koşulları </a> </li>
		<li> <a href="#"> Yardım Merkezi </a> </li>
		<li> <a href="iletisim.htm"> İletişim </a> </li>
		<li> <a href="benkimim.htm"> Ben Kimim? </a> </li>
		<li> <a href="#">  </a> </li>
		</ul>
	
	<div style="clear:both;">
	<center> 
	 <hr> Copyright © 2018 TÜM HAKLARI SAKLIDIR. <hr>	</center>
	</div>
	

</body>



<html>
<%end if%>