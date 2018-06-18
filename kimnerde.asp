<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Kim? Nerde ? Ne Yapıyor?</title>
        <link rel="stylesheet" type="text/css" href="stil.css">
    </head>
    <body>
        <%
        
            Set oConn = Server.CreateObject("ADODB.Connection")
            oConn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("database.mdb"))
            ssql="select * from Kullanıcı_Bilgileri where Ad IS NOT NULL and Soyad IS NOT NULL and YAS IS NOT NULL and Lise IS NOT NULL and Mezuniyet_yili IS NOT NULL and Meslek IS NOT NULL and Mail IS NOT NULL ;"
            set say = oConn.Execute(ssql)
                    
              do while not say.eof
                a=a+1
                say.movenext
              loop
              set say = nothing
              if a=0 then
                  sayfa=0
              end if
             sonsayfa=(int((a-1)/5))+1

            Set oRS = oConn.Execute(ssql)
        
        %>
         <div class="AnaDiv">
            <div class="UstMenu">
                <div id="foto">
                    <a href="index.html"><img src="arkaplan.jpg" alt="tema" height="150px" width="1000px" ></a>
                </div>
                <div id="giris">
                    <%  
                        If Request.QueryString("Sayfa") = "" Then
	                         sayfa = 1
                        Else
	                          sayfa = cInt(Request.QueryString("Sayfa"))
                        end if

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
                            Set RS = oConn.Execute(sSQL)     
                           
                            if RS.EOF then
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
               </center>     
            </div>
            <div class="icerik">
               <table style="margin-left: 20px;margin-top: 90px;margin-right: 10px;" border="1" style="width: 100%" class="tablo">
                <tr>
                    <th>Ad</th>
                    <th>Soyad</th>
                    <th>Yaş</th>
                    <th>Lise</th>
                    <th>Okuduğu Yıllar</th>
                    <th>Üniversite</th>
                    <th>Bölüm</th>
                    <th>Meslek</th>
                    <th>Mail</th>
                    
                </tr>
            <%
                do while not oRS.EOF
                i=i+1
                if i>=((sayfa-1)*5)+1 and i<=sayfa*5 then
            %>
                <tr>
                    <td><%=oRS("Ad")%></td>
                    <td><%=oRS("Soyad")%></td>
                    <td><%=oRS("Yas")%></td>
                    <td><%=oRS("Lise")%></td>
                    <td><%=oRS("Mezuniyet_yili")%></td>
                    <td><%=oRS("uni")%></td>
                    <td><%=oRS("Bolum")%></td>
                    <td><%=oRS("Meslek")%></td>
                    <td><%=oRS("Mail")%></td>
                </tr>
            <%
             end if
                oRS.MoveNext
             Loop
            %>
                </table>
            <center>
                <%
                    if sayfa>1 then
                %>
                    <a href="kimnerde.asp?Sayfa=<%=1%>"><B>|< İlk Sayfa </B></a>
                    <a href="kimnerde.asp?Sayfa=<%=sayfa-1%>"><span style="margin-left: 5px"><b>< Önceki Sayfa</b></span></a>
                <% end if%>
                    <span style="margin-left: 5px"><b> <%=sayfa%> / <%=sonsayfa%> </b></span>
                <%
                    if sonsayfa<>sayfa then
                %>
                    <a href="kimnerde.asp?sayfa=<%=sayfa+1%>"><span style="margin-left: 5px"><b>Sonraki Sayfa></b></span></a>
                    <a href="kimnerde.asp?sayfa=<%=sonsayfa%>"><span style="margin-left: 5px"><b>Son Sayfa >|</b></span></a>
               <%end if%>
             </center>
              </center>
            <%
                oConn.Close
                Set oRS = Nothing
                Set oConn = Nothing
            %>
                veritabanındaki kayıt sayısı:<%=i%>
            </div>
        </div>
                <%
                if session("UserLoggedIn") = "" then
                    response.redirect("Login.asp")
                 end if
        %>
    </body>
</html>
