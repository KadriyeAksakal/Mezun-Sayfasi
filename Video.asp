<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Videolar</title>
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
                         if session("UserLoggedIn") <> "" then
                           ' response.write("<h4>"&"Hoşgeldiniz "&session("UserLoggedIn"))
                    %>
                          <h4 style="color: #ffd800;float: right;margin-right: 30px; text-shadow: 5px 2px 1px #000">Hoşgeldin <%=response.write(session("UserLoggedIn"))%></h4>
                    <%    
                         else
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
                        else             
                    %>
                    <h4 style="float: right;margin-right: 20px;color: #ffd800;    text-shadow: 5px 2px 1px #000">Hoşgeldin <%= response.write(request.form("KullanıcıAd"))%></h4>
                    <%
                        Session("UserLoggedIn") =request.form("KullanıcıAd")
                        end if
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
                         if session("UserLoggedIn") <> " " then
                            response.write("<hr>")
                            response.write ("<a href='Cikis.asp'>Çıkış Yap!</a>")
                        else
                            response.write("<hr>")
                            response.write("<a href='kayitsayfasi.asp'>Kayıt Ol!</a>")
                        end if

                     %>                  
            </div>
            
              <div class="icerik">
                <table border="0" style="margin-left: 20px;margin-top: 70px;width: 95%">
                  
                    <tr>
                        <td rowspan="4"><a href="VideoSayfasi.asp"><img src="kep.png" alt="konser teoman" style="width: 350px;height: 160px"></a></td>
                        <td ><b>Ekleyen:</b></td>
                        <td style="width: 40%"><a href="huseyin.asp">Münevver Akdoğan</a></td>
                    </tr>
                    <tr>
                        <td><b>Yanındakiler:</b></td>
                        <td><a href="Furkan.asp">12-C sınıfı</a></td>
                    </tr>
                    <tr>
                        <td><b>Açıklama:</b></td>
                        <td>Mezuniyet Hatırası</td>
                    </tr>
                    <tr>
                        <td><b>Ekleme Tarihi:</b></td>
                        <td>19.04.2018 02:30</td>
                    </tr>
                    <tr>
                        <td colspan="6" style="height: 20px"></td>
                    </tr>
                     <tr>
                        <td rowspan="4"><a href="YoutubeVideo.asp"><img src="bayrak.jpg" alt="athena " style="width: 350px;height: 160px"></a></td>
                        <td ><b>Ekleyen:</b></td>
                        <td style="width: 40%"><a href="Vefa.asp">Zafer Sarı</a></td>
                    </tr>
                    <tr>
                        <td><b>Yanındakiler:</b></td>
                        <td>-</td>
                    </tr>
                    <tr>
                        <td><b>Açıklama:</b></td>
                        <td>:) </td>
                    </tr>
                    <tr>
                        <td><b>Ekleme Tarihi:</b></td>
                        <td>19.04.2018 02:40</td>
                    </tr>
                    <tr>
                        <td colspan="6" style="height: 20px"></td>
                    </tr>
                    <tr>
                        <td rowspan="4"><a href="VideoSayfasi2.asp"><img src="kep2.png" alt="konser teoman2" style="width: 350px;height: 160px"></a></td>
                        <td ><b>Ekleyen:</b></td>
                        <td style="width: 40%"><a href="huseyin.asp">Münevver Akdoğan</a></td>
                    </tr>
                    <tr>
                        <td><b>Yanındakiler:</b></td>
                        <td><a href="Furkan.asp">12-C</a></td>
                    </tr>
                    <tr>
                        <td><b>Açıklama:</b></td>
                        <td>Mezuniyet hatırası2</td>
                    </tr>
                    <tr>
                        <td><b>Ekleme Tarihi:</b></td>
                        <td>19.04.2018 02:50</td>
                    </tr>
                    <tr>
                        <td colspan="6" style="height: 20px"></td>
                    </tr>
                </table>
            </div>
        </div>
        <%
            if session("UserLoggedIn") = "" then
                response.redirect("Login.asp")
             end if
        %>
    </body>
</html>
