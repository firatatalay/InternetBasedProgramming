
<head><meta charset="utf-8"></head>
<% 
'kutuyu boş bırakmayı engelleme

'--------------
'VT baglantisinin yapimasi:
Set Baglantim = CreateObject("ADODB.Connection") 
'VT'nin acilmasi:
Baglantim.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ="& Server.MapPath("Veritabanim.mdb"))
'Tablo nesnesinin olusturulmasi:
Set Tablom = server. CreateObject("ADODB.Recordset")
'Tablonun acilmasi:
Tablom.Open "kayitt", Baglantim, 1, 3

'Tabloya veri eklemeye baslangic:
Tablom.AddNew 
'Tablodaki alanlara veri aktarma
Tablom("Ad") =  request("ad")
Tablom("Soyad") =  request("soyad")
Tablom("Cins") =  request("cins")
Tablom("Username") =  request("username")
Tablom("Pass") =  request("pass")
Tablom("Mail") =  request("mail")
Tablom("Mailtekrar") =  request("mailtekrar")
Tablom("Hobi") =  request("hobi")
Tablom("Muhendislik") =  request("Muhendislik")
Tablom("Fakulte") =  request("fakulte")
Tablom("Sehir") =  request("sehir")
Tablom("Adres") =  request("adres")
Tablom("Gizliyanit") =  request("gizliyanit")
Tablom("Yanit") =  request("yanit")
'aktarma islemi birince tablonun guncellenmesi:
Tablom.Update

'tablonun kapatilmasi:
  Tablom.close
  set Tablom= Nothing
'baglantinin kesilmesi:
  Baglantim.close
  set Baglantim= Nothing

response.redirect("giris.htm")
%>
<p><a href="kayit.htm"></a></p>