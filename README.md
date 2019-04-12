# signature_tool

Powershell script that automatically creates Outlook email signatures for Active Directory users.

Script is provided with a sample email signature, located in "templates" folder, copy the folder contents to "%AppData%\Microsoft\Signatures" & edit through Outlook (File -> Options - Mail -> Signatures).

sg_generate.ps1 generates 2x signatures for every AD user (unless added to $ignore_list): new_msg & reply_msg and places them all in user_signatures folder.

sg_autostart.ps1 should be setup through GPO to execute on start & placed somewhere on a file server that is accessible to the users. Enabling Powershell script code execution is dangerous, use at your own risk (this is just proof of concept).

sg_generate.ps1 basically takes the templates and replaces variables in it; FullNameVar, TitleVar, EmailVar, PhoneVar, MobileVar etc.
Keep in mind that the script pulls data from Active Directory user attributes - if your users don't have their details on their AD profiles, it probably won't work for you.

Considering this would be ran as an autostart on Windows login - the script generates a file with date on generation of the signature, so that once the signature is copied to the machine, it does not copy it again (or overwrite):
generate signatures -> date x -> user logins -> does the signature exist: no? copy it; yes? -> compare dates: date same as generation date; yes -> skip; no: update the signature
