#Signature tool current user
$current_user = $env:UserName


<#
silent/quiet powershell startup script
  checks if its connected to the network - ping DC? if no response -> try again in 10 minutes?
  if connection to the network is successful -> check if the signature exists -> if not copy it and run the regedit script
  -> if the signature exists -> check the date -> if
                                                    different -> copy the signature and copy the generation date
                                                    same -> skip
#>

$sg_ignore_list = "TEMP1", "TEMP2" #funny enough, it doesnt work if it has only 1 item, dunno why, dont care to find out anyway..

#if($current_user -in $sg_ignore_list){
if($current_user -contains $sg_ignore_list){
  write-host "-$current_user is on ignore list, stopping..."
  Start-Sleep -s 1
  write-host "`r3" -NoNewLine
  Start-Sleep -s 1
  write-host "`r2" -NoNewLine
  Start-Sleep -s 1
  write-host "`r1"
  Start-Sleep -s 1
}
else{
  #check if computer can ping to FILESERVER_NAME
  function CheckNetwork{
    if ((Test-Connection FILESERVER_NAME -quiet)){
      write-host "+FILESERVER_NAME reachable.."
    }else{
      #if not, wait 10 minutes - loop
      write-host "-Cannot reach FILESERVER_NAME, waiting 10 minutes and trying again."
      write-host "-Possible that the user is not connected to the WiFi/Network."
      Start-Sleep -s 600
      CheckNetwork
      write-host "*Trying again."
    }
  }
  CheckNetwork




  #So its easier to change it in the future
  $remote_signature_directory = "\\FILESERVER_NAME\LOCATION_OF_THE_SCRIPT\signature_tool\user_signatures\"

  #finding the user's signature
  $user_specific_signature = $remote_signature_directory + $current_user
  #debugging
  #write-host $user_specific_signature

  #vars needed for the copy of signature files to the local computer
  $OutlookSignatureFolder = "C:\Users\" + $current_user + "\AppData\Roaming\Microsoft\Signatures\"
  $user_specific_signature_formatted = $user_specific_signature + "\*"

  #function that does the signature copy & registry hack
  function CopySignatureToLocalUser{
    #copy the signature
    Copy-Item -Path $user_specific_signature_formatted -Destination $OutlookSignatureFolder -recurse -Force

    #registry hack to set the signature as default without user interaction
    write-host "+Registry magic.."
    $new_signature = "new_msg"
    $reply_signature = "reply_msg"

    if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") {
          get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
          get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $new_signature -propertytype string -force
          get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $reply_signature -propertytype string -force
    }
    if (test-path "HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\General") {
      get-item -path HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
      get-item -path HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $new_signature -propertytype string -force
      get-item -path HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $reply_signature -propertytype string -force
    }
    if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") {
      get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
      get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $new_signature -propertytype string -force
      get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $reply_signature -propertytype string -force
    }
  }#end function

  #check if date of the generation is different to decide whether the update is needed
  #check if file exists
  $LocalLastUpdate_file = $OutlookSignatureFolder + "last_update.dat"
  If(!(test-path $LocalLastUpdate_file))
  {
  	Write-host "+Last update file does not exist on the target"
    write-host "+Signature auto start script might be running for the first time"
    CopySignatureToLocalUser
  }
  else{
  	write-host "+Last update file does exist, comparing it.."
    $ServerLastUpdate_file = "\\FILESERVER_NAME\LOCATION_OF_THE_SCRIPT\signature_tool\templates\last_update.dat"

    #comparing files does not do anything, you have to compare the content of the file - hence the vars
    $data_from_ServerLastUpdate_file = Get-Content $ServerLastUpdate_file
    $data_from_LocalLastUpdate_file = Get-Content $LocalLastUpdate_file

  	if(Compare-Object $data_from_LocalLastUpdate_file $data_from_ServerLastUpdate_file){
      #vars are different, so copy the file
      write-host "+Dates are different, running the update of the signature.."
      CopySignatureToLocalUser
  	}
  	else{
      write-host "+Dates are the same, skipping signature update"
      write-host "---------------------------------------"
      #vars are the same, skip
  	}
  }
  Write-host "Done! Exiting automatically..."
  Start-Sleep -s 1
  write-host "`r3" -NoNewLine
  Start-Sleep -s 1
  write-host "`r2" -NoNewLine
  Start-Sleep -s 1
  write-host "`r1"
  Start-Sleep -s 1

}
