<%
  Set rsDemo = Conn.Execute("SELECT ID FROM tbl_name")
  If rsDemo.Eof Then
    Conn.BeginTrans
    Set objExec = Conn.Execute("INSERT INTO tbl_cronjob(CRON_FUNCTION,TANIM) VALUES('"& FonksiyonAdi &"','"& KisaAdi &"')")
    Set objExec = Nothing
    If Err.Number = 0 Then 
      Conn.CommitTrans
    Else 
      Conn.RollbackTrans
      Response.write("Bir Hata Oluştu")  
    End If
  Else 

  End If
  rsDemo.Close : Set rsDemo = Nothing
%>
