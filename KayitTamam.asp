<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Kayıt Tamamlandı</title>
        <link rel="stylesheet" type="text/css" href="stil.css">
    </head>
    <body>
        
         <div class="AnaDiv">
            <div class="UstMenu">
                <div id="foto">
                   <a href="index.html"><img src="arkaplan.jpg" alt="tema" height="150px" width="1000px" ></a>
                </div>
                <div id="giris"> 
                    <%    
                            dim username,sifre
                            Set oConn = Server.CreateObject("ADODB.Connection")
                            oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("database.mdb"))
                            ssql="select Ad+' '+Soyad,Sifre from Kullanıcı_Bilgileri where Ad+' '+Soyad='" & request.form("KullanıcıAd")& "' and Sifre='"& request.form("KullanıcıSif")&"';"
                            Set oRS = oConn.Execute(sSQL)     
                           
                            if oRS.EOF then
                    %>
                        <form action="index.asp" method="post">
                        <table style="width: 100%;height: 90%;margin-top: 10px" border="0" >
                            <tr>
                                <td style="width: 110px">
                                    <span>Ad Soyad</span>
                                    <br>
                                    <input type="text" style="border-radius: 10px" name="KullanıcıAd">   
                                </td>
                                <td style="width: 170px">
                                    Şifre
                                    <br>
                                     <input type="password" style="border-radius: 10px" name="KullanıcıSif">
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <a href="sifremiunuttum.asp"><span style="margin-left: 160px">Şifremi Unuttum</span></a>
                                    <a href=""><input type="submit" value="Giriş Yap" style="color:  #d25313;margin-left: 20px;"></a>  
                                </td>
                            </tr>
                    
                        </table>
                        </form>
                    <%
                        Session("UserLoggedIn") =request.form("KullanıcıAd")
                        end if
                    %>

                </div>
                <span style=" font-family: 'Times New Roman'; padding-left:100px; font-size: 45px;  font:bold 30px tahoma;  color: #000000; text-align: center"><i>Necip Fazıl Anadolu Lisesi</i></span>
            </div>
            <div class="SolMenu">
                <center>
             <br>
             <a href="index.asp">Ana Sayfa</a>
             <hr>
             <a href="kisiler.asp">Kişiler</a>
             <hr>
             <a href="Fotolar.asp">Fotoğraflar</a>
             <hr>
             <a href="video.asp">Videolar</a>
             <hr>
             <a href="kimnerde.asp">Kim,Nerde,Ne Yapıyor?</a>
             <hr>
             <a href="Forum.asp">Forum</a>
             <hr>
             <a href="Harita.asp">Harita</a>
             <hr>
             <a href="KisiAra.asp">Ara</a>
             <%
                         if session("UserLoggedIn") <> "" then
                            response.write("<hr>")
                            response.write ("<a href='Cikis.asp'>Çıkış Yap!</a>")
                        else
                            response.write("<hr>")
                            response.write("<a href='kayitsayfasi.html'>Kayıt Ol!</a>")
                        end if

                     %>                
             </center>
            </div>
            <div class="icerik">
                  <%
                    if trim(request.form("isim"))="" or trim(request.form("soyisim"))="" or trim(request.form("yas"))="" or trim(request.form("cinsiyet"))="" or trim(request.form("lise"))="" or trim(request.form("mail"))="" or trim(request.form("ps1"))="" or trim(request.form("ps"))="" then
                        response.write("<center><h4>Eksik Bilgi . Lütfen Boş Bırakmayanız. [<a href=""javascript:history.back()""]>Geri Gelmek İçin Tıklayınız</a></center>")
                    elseif trim(request.form("ps1"))<>trim(request.form("ps"))then
                        response.write("<center><h4>Girilen Şifreler Birbirini Tutmuyor[<a href=""javascript:history.back()""]>Geri Gelmek İçin Tıklayınız</a></center>")
                    else
                        response.write(" <center><h1>Kayıdınız Tamamlanmıştır . İyi Eğlenceler :)</h1> </center>")
                        'VT baglantisinin yapimasi:
                        Set Baglantim = CreateObject("ADODB.Connection") 
                        'VT'nin acilmasi:
                        Baglantim.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ="& Server.MapPath("database.mdb"))
                        'Tablo nesnesinin olusturulmasi:
                        Set Tablom = server. CreateObject("ADODB.Recordset")
                        'Tablonun acilmasi:
                        Tablom.Open "Kullanıcı_Bilgileri", Baglantim, 1, 3      
            
                        'Tabloya veri eklemeye baslangic:
                        Tablom.Addnew
                        'Tablodaki alanlara veri aktarma
                        Tablom("Ad") = trim(request("isim"))
                        Tablom("Soyad") = trim(request("soyisim"))
                        Tablom("Yas") = request("yas")
                        Tablom("Cinsiyet") = request("cinsiyet")
                        Tablom("Lise") = request("lise")
                        Tablom("Mezuniyet_yili") = request("yil1") +" - "+request("yil2")
                        Tablom("uni") = request("uni")
                        Tablom("Bolum") = request("bolum")
                        Tablom("Meslek") = request("meslek")
                        Tablom("Mail") = request("mail")
                        Tablom("Adres") = request("adres")
                        Tablom("Hobiler")="-"
                        Tablom("Sifre") = request("ps")
                        'aktarma islemi birince tablonun guncellenmesi:
                        Tablom.Update     

                        'tablonun kapatilmasi:
                          Tablom.close
                          set Tablom= Nothing
                        'baglantinin kesilmesi:
                          Baglantim.close
                          set Baglantim= Nothing
                          end if
                    %>

            </div>
        </div>
    </body>
</html>
