<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    </head>
    <body>
        Amerikada şu an saat: <%=time()%>
        <%      dim saat_tr
                'Şu anki USA saatini yerel saate cevirme:
                sat = split(time(),":",-1,1)
                tarih = split(date(),"/",-1,1)
                if right(time(),2)="PM" and sat(0)<>12 then
                sat(0) = sat(0) + 12
                end if
                sat(0) = sat(0) + 7
                if sat(0)>=24 then
                sat(0)=sat(0)-24
                end if
                saat_tr = Date()&" "& sat(0) & ":" & sat(1) & ":" & left(sat(2),2)
        %>
        <br>
        Türkiyede Şu anki saat: <%=saat_tr%>
    </body>
</html>
