
#---------------------------------------------------------------------------------
###   fill out the variables as per requirement  ###
$SMTPServerName      = 'Enterprise-SMTPServer-DNS-Name'
$ReveipientsList     = "alam.sohel1990@gmail.com;anuj.ray@gmail.com"
$SenderMailID        = "anuj.ray@gmail.com"
$RemoteServerPath    = "\\server\REmoteFolderPath"
$RemoteFolderNames   = "FOLDER1;FOLDER2;FOLDER3"
$MailBodySignature   = "Sohel Alam"
#---------------------------------------------------------------------------------

# Uncomment below line if the mails are not being sent..
# $ReveipientsList     = $ReveipientsList   -split ';' 
$RemoteFolderNames   = $RemoteFolderNames -split ';'

Function Get-MailBody{
    param([String]$FolderName)

    $MailBodySignature = $Global:MailBodySignature

    $MailBody =  @"

    
    <body lang=EN-US link="#0563C1" vlink="#954F72" style='tab-interval:.5in;word-wrap:break-word'>

    <div class=WordSection1>

    <p class=MsoNormal>Hi Team,<o:p></o:p></p>

    <p class=MsoNormal><o:p>&nbsp;</o:p></p>

    <p class=MsoNormal><o:p>&nbsp;</o:p></p>

    <p class=MsoNormal>Kindly find below Report for FOLDER_NAME at DATE_TIME .<o:p></o:p></p>

    <p class=MsoNormal><o:p>&nbsp;</o:p></p>

    <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=473
     style='width:354.9pt;border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;
     mso-yfti-tbllook:1184;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
     <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:14.3pt'>
      <td valign=top style='border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .5pt;
      padding:0in 5.4pt 0in 5.4pt;height:14.3pt'>
      <p class=MsoNormal><b><span style='color:#203864;mso-themecolor:accent1;
      mso-themeshade:128;mso-style-textfill-fill-color:#203864;mso-style-textfill-fill-themecolor:
      accent1;mso-style-textfill-fill-alpha:100.0%;mso-style-textfill-fill-colortransforms:
      lumm=50000'>File Name<o:p></o:p></span></b></p>
      </td>
      <td valign=top style='border:solid windowtext 1.0pt;border-left:none;
      mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
      padding:0in 5.4pt 0in 5.4pt;height:14.3pt'>
      <p class=MsoNormal><b><span style='color:#203864;mso-themecolor:accent1;
      mso-themeshade:128;mso-style-textfill-fill-color:#203864;mso-style-textfill-fill-themecolor:
      accent1;mso-style-textfill-fill-alpha:100.0%;mso-style-textfill-fill-colortransforms:
      lumm=50000'>Last Write Time<o:p></o:p></span></b></p>
      </td>
      <td valign=top style='border:solid windowtext 1.0pt;border-left:none;
      mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
      padding:0in 5.4pt 0in 5.4pt;height:14.3pt'>
      <p class=MsoNormal><b><span style='color:#203864;mso-themecolor:accent1;
      mso-themeshade:128;mso-style-textfill-fill-color:#203864;mso-style-textfill-fill-themecolor:
      accent1;mso-style-textfill-fill-alpha:100.0%;mso-style-textfill-fill-colortransforms:
      lumm=50000'>Size(KB)<o:p></o:p></span></b></p>
      </td>
     </tr>

     Append_Mail_Body

    </table>

    <p class=MsoNormal><o:p>&nbsp;</o:p></p>

    <p class=MsoNormal>Regards,<o:p></o:p></p>

    <p class=MsoNormal>Anuj Ray<o:p></o:p></p>

    </div>

    </body>
        

    <!-- ========================== Signature ==========================. -->
    <p class=MsoNormal><o:p>&nbsp;</o:p></p>

    <p class=MsoNormal><a name="_MailAutoSig"><span lang=EN-US style='font-family:
    "Century Gothic",sans-serif;mso-fareast-font-family:"Times New Roman";
    mso-fareast-theme-font:minor-fareast;mso-ansi-language:EN-US;mso-fareast-language:
    EN-IN;mso-no-proof:yes'>Thanks, <o:p></o:p></span></a></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><b><span lang=EN-US
    style='font-size:10.5pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
    minor-fareast;color:#1C23BA;mso-ansi-language:EN-US;mso-fareast-language:EN-IN;
    mso-no-proof:yes'>Anuj Ray</span></b></span><span style='mso-bookmark:_MailAutoSig'><span
    lang=EN-US style='font-size:10.5pt;font-family:"Times New Roman",serif;
    mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:minor-fareast;
    mso-ansi-language:EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'><o:p></o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span lang=EN-US
    style='font-size:9.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:gray;mso-ansi-language:
    EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'>Sr. Windows Engineer- IT
    Infra Operations</span></span><span style='mso-bookmark:_MailAutoSig'><span
    lang=EN-US style='font-size:10.5pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:black;mso-ansi-language:
    EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'><o:p></o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span lang=EN-US
    style='font-size:9.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:gray;mso-ansi-language:
    EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'>&nbsp;</span></span><span
    style='mso-bookmark:_MailAutoSig'><span lang=EN-US style='font-size:10.5pt;
    mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
    mso-fareast-theme-font:minor-fareast;mso-hansi-font-family:Calibri;mso-bidi-font-family:
    Calibri;color:black;mso-ansi-language:EN-US;mso-fareast-language:EN-IN;
    mso-no-proof:yes'><o:p></o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span lang=EN-US
    style='font-size:9.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:gray;mso-ansi-language:
    EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'>Mobile&nbsp;&nbsp;&nbsp;&nbsp;
    :&nbsp;&nbsp;+91 9920075684</span></span><span style='mso-bookmark:_MailAutoSig'><span
    lang=EN-US style='font-size:10.5pt;mso-fareast-font-family:"Times New Roman";
    mso-fareast-theme-font:minor-fareast;color:black;mso-ansi-language:EN-US;
    mso-fareast-language:EN-IN;mso-no-proof:yes'><o:p></o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span lang=EN-US
    style='font-size:9.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:gray;mso-ansi-language:
    EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'>Zoom Ext:&nbsp;&nbsp;&nbsp;
    6128259<o:p></o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'></span><a
    href="mailto:anuj.ray@intusurg.com"><span style='mso-bookmark:_MailAutoSig'><span
    lang=EN-US style='font-size:9.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:#0563C1;
    mso-ansi-language:EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'>anuj.ray@intusurg.com</span></span><span
    style='mso-bookmark:_MailAutoSig'></span></a><span style='mso-bookmark:_MailAutoSig'><span
    lang=EN-US style='font-size:10.5pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;mso-hansi-font-family:
    Calibri;mso-bidi-font-family:Calibri;color:black;mso-ansi-language:EN-US;
    mso-fareast-language:EN-IN;mso-no-proof:yes'><o:p></o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span lang=EN-US
    style='font-size:10.5pt;mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:
    minor-fareast;color:black;mso-ansi-language:EN-US;mso-fareast-language:EN-IN;
    mso-no-proof:yes'><o:p>&nbsp;</o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span lang=EN-US
    style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:#1630F5;
    mso-ansi-language:EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'>INTUITIVE</span></span><span
    style='mso-bookmark:_MailAutoSig'><span lang=EN-US style='font-size:10.5pt;
    mso-fareast-font-family:"Times New Roman";mso-fareast-theme-font:minor-fareast;
    color:black;mso-ansi-language:EN-US;mso-fareast-language:EN-IN;mso-no-proof:
    yes'><o:p></o:p></span></span></p>

    <p class=MsoNormal><span style='mso-bookmark:_MailAutoSig'><span lang=EN-US
    style='font-size:9.0pt;font-family:"Arial",sans-serif;mso-fareast-font-family:
    "Times New Roman";mso-fareast-theme-font:minor-fareast;color:#757B80;
    mso-ansi-language:EN-US;mso-fareast-language:EN-IN;mso-no-proof:yes'>India.</span></span><span
    style='mso-bookmark:_MailAutoSig'><span style='mso-fareast-font-family:"Times New Roman";
    mso-fareast-theme-font:minor-fareast;mso-fareast-language:EN-IN;mso-no-proof:
    yes'><o:p></o:p></span></span></p>

    <span style='mso-bookmark:_MailAutoSig'></span>

    <p class=MsoNormal><o:p>&nbsp;</o:p></p>
"@ 

    $MailBody = $MailBody -replace 'FOLDER_NAME' ,$FolderName
    $MailBody = $MailBody -replace 'DATE_TIME' , (Get-Date -F 'dd-MMM-yyyy hh:mm' )

    return $MailBody
}

Function Append-MailBody{
    param(
    [String]$mailBody,
    [String]$AppendFileName,
    [String]$AppendFileLastWriteTime,
    [String]$AppendFileSize
    )
    $AppendMailBody = @"
     <tr style='mso-yfti-irow:1;height:14.3pt'>
      <td valign=top style='border:solid windowtext 1.0pt;border-top:none;
      mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
      padding:0in 5.4pt 0in 5.4pt;height:14.3pt'>
      <p class=MsoNormal><o:p>Append_File_Name</o:p></p>
      </td>
      <td valign=top style='border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;
      border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
      mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
      padding:0in 5.4pt 0in 5.4pt;height:14.3pt'>
      <p class=MsoNormal><o:p>Append_File_LastWriteTime</o:p></p>
      </td>
      <td valign=top style='border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;
      border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
      mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;
      padding:0in 5.4pt 0in 5.4pt;height:14.3pt'>
      <p class=MsoNormal><o:p>Append_File_Size</o:p></p>
      </td>
     </tr>
     Append_Mail_Body
"@
    
    $AppendMailBody = $AppendMailBody -replace 'Append_File_Name'          ,  $AppendFileName
    $AppendMailBody = $AppendMailBody -replace 'Append_File_LastWriteTime' ,  $AppendFileLastWriteTime
    $AppendMailBody = $AppendMailBody -replace 'Append_File_Size'          ,  $AppendFileSize

    $NewMailBody = $mailBody -replace 'Append_Mail_Body',$AppendMailBody 

}

Function Sanitize-MailBody{
    param([String]$mailBody)
    $mailBody = $mailBody -replace 'Append_Mail_Body' , ''
    return $mailBody
}

ForEach($RemoteFolderName in $RemoteFolderNames){
    $RemoteFolderFullPath = "$($RemoteServerPath)\$($RemoteFolderName)"

    If( Test-path -literalPath $RemoteFolderFullPath ){
        $MailBody = Get-MailBody -FolderName $RemoteFolderName
        $Files = Get-ChildItem -LiteralPath $RemoteFolderFullPath -force # filter output according to need... Do we need only files and not the subfolders ? etc.
        ForEach($File in $Files){
            $MailBody = Append-MailBody -AppendFileName "$($File.Name)" -AppendFileLastWriteTime "$($File.LastWriteTime | Out-String )" -AppendFileSize "$("{0:n2} KB" -f ($File.Length / 1KB))" -mailBody $MailBody
        }
        $MailBody = Sanitize-MailBody -mailBody $MailBody
        Send-MailMessage -Body $MailBody -BodyAsHtml -From $SenderMailID -SmtpServer $SMTPServerName -To $ReveipientsList `
        -subject "File Status Report for folder $RemoteFolderName" # change the subject as per need...
    }
    Else{
        Write-Host "Folder Not Found" -F Red
    }
}