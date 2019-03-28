<%@ Language="VBScript" %>
<%
Quote = Chr(34)

KB = 1024
MB = 1048576
GB = 1073741824
TB = 1099511627776
PB = 1125899906842624
EnormousSize = 1152921504606846976

Response.Write "<html>"
Response.Write "<head>"
Response.Write "<meta http-equiv=" & Quote & "Content-Type" & Quote & " content=" & Quote & "text/html; charset=windows-1251" & Quote & ">"
Response.Write "<title>System Information</title>"
Response.Write "</head>"
Response.Write "<body>"

Response.Write "<p>System time: " & Now & "</p>"
Response.Write "<p>Local IP: " & Request.ServerVariables("LOCAL_ADDR") & "<br>Remote IP: " & Request.ServerVariables("REMOTE_ADDR") & "</p>"

Response.Write "<table border=" & Quote & "1" & Quote & ">"
Response.Write "<caption align style=" & Quote & "text-align: left" & Quote & ">Drive usage:</caption>"
Response.Write "<tr><th width=" & Quote & "120" & Quote & ">Drive Label</th><th width=" & Quote & "100" & Quote & ">Total Space</th><th width=" & Quote & "80" & Quote & ">Used</th></tr>"

set FileSystem = Server.CreateObject("Scripting.FileSystemObject")

for each Drive in FileSystem.Drives
  '1 - removable drives, 2 - logical drives
  if Drive.DriveType = 2 then
    TotalDriveSize = GetTotalDriveSize(Drive)
    FreeDriveSpace = GetFreeDriveSpace(Drive)

    Used = Round((TotalDriveSize - FreeDriveSpace) / TotalDriveSize * 100, 2)
    TotalDriveSizeStr = ""

    if TotalDriveSize < KB then
      TotalDriveSizeStr = TotalDriveSize
    elseif TotalDriveSize < MB then
      TotalDriveSizeStr = Round(TotalDriveSize / KB, 2) & " KB"
    elseif TotalDriveSize < GB then
      TotalDriveSizeStr = Round(TotalDriveSize / MB, 2) & " MB"
    elseif TotalDriveSize < TB then
      TotalDriveSizeStr = Round(TotalDriveSize / GB, 2) & " GB"
    elseif TotalDriveSize < PB then
      TotalDriveSizeStr = Round(TotalDriveSize / TB, 2) & " TB"
    elseif TotalDriveSize < EnormousSize then
      TotalDriveSizeStr = Round(TotalDriveSize / PB, 2) & " PB"
    else
      TotalDriveSizeStr = "Infinity"
    end if

    Response.Write "<tr><td>" & Drive.VolumeName & "</td><td>" & TotalDriveSizeStr & "</td><td>" & Used & "%</td></tr>"
  end if
next

set FileSystem = Nothing

Response.Write "</table>"

Response.Write "<p>System uptime: " & GetUpTime & "</p>"

Response.Write "</body>"
Response.Write "</html>"

function GetTotalDriveSize(Drive)
  GetTotalDriveSize = FileSystem.GetDrive(Drive.DriveLetter).TotalSize
end function

function GetFreeDriveSpace(Drive)
  GetFreeDriveSpace = FileSystem.GetDrive(Drive.DriveLetter).FreeSpace
end function

function GetUpTime
  dim WMIService
  dim OperatingSystems

  dim LastBootUpTime

  dim UpTimeSeconds
  dim UpTimeMinutes
  dim UpTimeHours
  dim UpTimeDays

  set WMIService = GetObject("winmgmts:\\localhost\root\cimv2")
  set OperatingSystems = WMIService.ExecQuery("select * from Win32_OperatingSystem")

  for each OperatingSystem in OperatingSystems
    LastBootUpTime = OperatingSystem.LastBootUpTime

    UpTimeSeconds = DateDiff("s", CDate(Mid(LastBootUpTime, 7, 2) & "." & Mid(LastBootUpTime, 5, 2) & "." & Mid(LastBootUpTime, 1, 4) & " " & Mid (LastBootUpTime, 9, 2) & ":" & Mid(LastBootUpTime, 11, 2) & ":" & Mid(LastBootUpTime, 13, 2)), Now)

    if UpTimeSeconds >= 60 then
      UpTimeMinutes = Int(UpTimeSeconds / 60)
      UpTimeSeconds = UpTimeSeconds mod 60
    else
      UpTimeMinutes = 0
      UpTimeHours = 0
      UpTimeDays = 0
    end if

    if UpTimeMinutes >= 60 then
      UpTimeHours = Int(UpTimeMinutes / 60)
      UpTimeMinutes = UpTimeMinutes mod 60
    else
      UpTimeHours = 0
      UpTimeDays = 0
    end if

    if UpTimeHours >= 24 then
      UpTimeDays = Int(UpTimeHours / 24)
      UpTimeHours = (UpTimeHours mod 24)
    else
      UpTimeDays = 0
    end if
  next

  set WMIService = Nothing
  set OperatingSystems = Nothing

  if UpTimeDays > 0 then
    GetUpTime = UpTimeDays & " days, " & UpTimeHours & " hours, " & UpTimeMinutes & " minutes, " & UpTimeSeconds & " seconds"
  elseif UpTimeHours > 0 then
    GetUpTime = UpTimeHours & " hours, " & UpTimeMinutes & " minutes, " & UpTimeSeconds & " seconds"
  elseif UpTimeMinutes > 0 then
    GetUpTime = UpTimeMinutes & " minutes, " & UpTimeSeconds & " seconds"
  else
    GetUpTime = UpTimeSeconds & " seconds"
  end if
end function
%>
