#Signature tool
<#
	email signature deployment script
		C:\Users\%username%\AppData\Roaming\Microsoft\Signatures - everything in .htm is treated like a signature

		powershell script
			ran daily or whenever signature needs updating - for every user in AD
				folder structure:
				__________________________
				|signature_tool
				|	signature_tool.ps1
				|	templates
				|		new_msg.htm
				|		reply_msg.htm
				|	user_signatures
				|		%username%
				|_________________________
				powershell script:
					for each user do function:
						for each user create folder with %username%
						for each user create 2 files - new_msg.htm & reply.htm (taken from "template" folder)
							change $first$, $lastname$, $job_title$, $mobile_phone$, $landline$ variables for data from AD
								#DisplayName,TelephoneNumber,Mobile,sAMAccountName,Title,EmailAddress
							if no title = remove title line
							if no mobile = remove mobile line

				batch powershell script (located on C:\folder) that pulls the complete signatures from a file share
					on logon, take %username%, copy the folder from fileshare and move it over to "\AppData\Roaming\Microsoft\Signatures"
						if folder cannot be reached - ignore/die
					execute a vbs that sets
						new_msg.htm as new message signature
						reply_msg.htm as reply message signature

            timeline:
            generate the signatures automatically for all users
              folder for each employee
              in the root of the script - create a date with when it has been last generated

            silent/quiet powershell startup script
              checks if its connected to the network - ping DC? if no response -> create a scheduled task to try again later?
              if connection to the network is successful -> check if the signature exists -> if not copy it and run the regedit script
              -> if the signature exists -> check the date -> if
                                                              different -> copy the signature and copy the generation date
                                                              same -> skip

#Random thoughts
PS: documentation above might be outdated, it was created before (or during the development) the tool was finished, comments are good so that should be enough
PS: In retrospect, this is overengineered as fuck, in 2012 different IT guy in $company had a guide in .doc on how to make email signature and it seemed to work?
		(Overengineering (or over-engineering) is the act of designing a product to be more robust or have more features than necessary for its intended use, or for a process to be unnecessarily complex or inefficient.)
PS: ^but at least it works well
#>

#logging of the errors - https://www.vexasoft.com/blogs/powershell/7255220-powershell-tutorial-try-catch-finally-and-error-handling-in-powershell
Try{
	write-host " --------------------------------------------------"
	write-host "| Email signature tool initiating..."
	$startTime = get-date
	write-host "| Started on $startTime"


	write-host " --------------------------------------------------"
	write-host "| *Importing required modules..."

	#Import ActiveDirectory module & users and their properties
	Import-Module ActiveDirectory
	write-host "| +Imported AD..."
	#catches lots of random service/non-user accounts, but #yolo
	#ignores everyone that has a .local email - filters out most service/non-user accounts
	#$users = Get-ADUser -LDAPFilter "(mail=*@$company.com)" -properties TelephoneNumber,Mobile,DisplayName,sAMAccountName,Title,EmailAddress|where {$_.Enabled -eq "True"} | select DisplayName,TelephoneNumber,Mobile,sAMAccountName,Title,EmailAddress
	#unfortunately due to migration to office365, the email^ in AD attributes does not work as well as it did before, hence get all users, use at your own risk #yolo
	$users = Get-ADUser -LDAPFilter "(mail=*)" -properties TelephoneNumber,Mobile,DisplayName,sAMAccountName,Title,EmailAddress|where {$_.Enabled -eq "True"} | select DisplayName,TelephoneNumber,Mobile,sAMAccountName,Title,EmailAddress
	write-host "| +Imported users..."
	write-host " --------------------------------------------------"


	#because im lazy and id rather hard code accounts that i want to filter out
	$ignore_list = "TEMP1", "TEMP2"

	#Run timer to measure the speed of the script
	$startTime = get-date
	write-host "| Started on $startTime"

	#output the time of the generation to the template so the autostart script can verify whether the signature has received an update
	$startTime | Out-File templates\last_update.dat

	#count amount of users
	$counter = 0
	$ignored_users_counter = 0

	Write-host "| Starting to generate signatures..."
	write-host " --------------------------------------------------"


	# Process Each User
	foreach ($user in $users)
	{
		#if user is in the ignore list - ignore
		#if($user.sAMAccountName -in $ignore_list){ #would work only on a newer version of powershell
		if($user.sAMAccountName -contains $ignore_list){
			$accountname= $user.sAMAccountName
			write-host "-$accountname is on ignore list, skipping..."
			$ignored_users_counter++
		}
		else{
			#counter
			$counter++
			#rewrite & set variables
			$name = $user.DisplayName
			$telephonenumber = $user.TelephoneNumber
			$mobilenumber = $user.Mobile
			$accountname= $user.sAMAccountName
			$title = $user.Title
			$emailaddress = $user.EmailAddress
			$emailaddress = $emailaddress | % { $_ -replace "local", "com"} #once inbox has been migrated into office365, Active Directory email might show up with .local instead of com

			write-host "---"
			write-host "*Generating email signature for $accountname"

			#create folder for the user
			$SignatureFolderPath = "user_signatures\$accountname"

			#check if folder exists already
			if(Test-Path -Path $SignatureFolderPath){
				#folder exists
			}
			else{
				# folder doesnt exist, create it
				new-item $SignatureFolderPath -type directory
			}


			#copy the templates to the folder
			Copy-Item -path templates\*.htm -Destination $SignatureFolderPath -force -Exclude reply_msg_no_mobile.htm
			Copy-Item -path templates\*.rtf -Destination $SignatureFolderPath -force
			Copy-Item -path templates\last_update.dat -Destination $SignatureFolderPath -force
			Copy-Item -path templates\new_msg_files -Recurse -Destination $SignatureFolderPath -Container -force

			#replace the template variables with attributes from active directory
			$new_msg_template = "$SignatureFolderPath\new_msg.htm"
			(get-content $new_msg_template) | % { $_ -replace "FullNameVar", $name } | set-content $new_msg_template
			(get-content $new_msg_template) | % { $_ -replace "TitleVar", $title } | set-content $new_msg_template
			(get-content $new_msg_template) | % { $_ -replace "EmailVar", $emailaddress } | set-content $new_msg_template
			(get-content $new_msg_template) | % { $_ -replace "PhoneVar", $telephonenumber } | set-content $new_msg_template
			(get-content $new_msg_template) | % { $_ -replace "MobileVar", $mobilenumber } | set-content $new_msg_template

			$new_msg_rtf_template = "$SignatureFolderPath\new_msg.rtf"
			(get-content $new_msg_rtf_template) | % { $_ -replace "FullNameVar", $name } | set-content $new_msg_rtf_template
			(get-content $new_msg_rtf_template) | % { $_ -replace "TitleVar", $title } | set-content $new_msg_rtf_template
			(get-content $new_msg_rtf_template) | % { $_ -replace "EmailVar", $emailaddress } | set-content $new_msg_rtf_template
			(get-content $new_msg_rtf_template) | % { $_ -replace "PhoneVar", $telephonenumber } | set-content $new_msg_rtf_template
			(get-content $new_msg_rtf_template) | % { $_ -replace "MobileVar", $mobilenumber } | set-content $new_msg_rtf_template

			$reply_msg_template = "$SignatureFolderPath\reply_msg.htm"
			#detect whether user has no mobile number and if thats the case, use a different template
			if ([string]::IsNullOrEmpty($mobilenumber)){
				write-host "   *User: $accountname has no mobile number.. using different template"
				Copy-Item -path templates\reply_msg_no_mobile.htm -Destination $SignatureFolderPath\reply_msg.htm -force
			}
			else{
				#debugging write-host "*User has mobile number"
				(get-content $reply_msg_template) | % { $_ -replace "MobileVar", $mobilenumber } | set-content $reply_msg_template
			}

			(get-content $reply_msg_template) | % { $_ -replace "FullNameVar", $name } | set-content $reply_msg_template
			(get-content $reply_msg_template) | % { $_ -replace "TitleVar", $title } | set-content $reply_msg_template
			(get-content $reply_msg_template) | % { $_ -replace "PhoneVar", $telephonenumber } | set-content $reply_msg_template

			#if mobile number is empty in AD properties, remove the line
			<#if(([string]::IsNullOrEmpty($mobilenumber))){
				#true empty
				(get-content $reply_msg_template) | % { $_ -replace "M:</FONT><FONT color=#666666>&nbsp;MobileVar&nbsp;</FONT>", "" } | set-content $reply_msg_template
			}elseif(!([string]::IsNullOrEmpty($mobilenumber))){
				#false, not empty, replace the var with number
				(get-content $reply_msg_template) | % { $_ -replace "MobileVar", $mobilenumber } | set-content $reply_msg_template
			}#>

			write-host "+Email signature for $accountname has been created."
		}#end of ignore list if statement
	} # End User Processing


	$endTime = get-date
	$totalTime = $endTime - $startTime
	write-host " --------------------------------------------------"
	write-host "| Finished in $totalTime seconds"
	Write-host "| Update finished at: $(Get-Date)"
	Write-host "| Processed $counter user signatures."
	Write-host "| Ignored $ignored_users_counter accounts users."
	write-host " --------------------------------------------------"
}
<#Catch [System.OutOfMemoryException]
{
    Restart-Computer localhost
}#>
Catch{
		$ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    #Send-MailMessage -From ExpensesBot@MyCompany.Com -To WinAdmin@MyCompany.Com -Subject "HR File Read Failed!" -SmtpServer EXCH01.AD.MyCompany.Com -Body "We failed to read file $FailedItem. The error message was $ErrorMessage"
		$Time=Get-Date
		write-host "-Error message:"
		"---------" | out-file error.log -append
		"$Time Error message:" | out-file error.log -append
		write-host $ErrorMessage
		$ErrorMessage | out-file error.log -append
		write-host $FailedItem
		$FailedItem | out-file error.log -append
		"---------" | out-file error.log -append

		write-host "*Errors have been saved to error.log"
		write-host "+Press any key to continue.."
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}
finally{
	write-host "+Script finished with no errors"
	Write-host "+Exiting..."
	Start-Sleep -s 1
	write-host "`r3" -NoNewLine
	Start-Sleep -s 1
	write-host "`r2" -NoNewLine
	Start-Sleep -s 1
	write-host "`r1"
	Start-Sleep -s 1
}
