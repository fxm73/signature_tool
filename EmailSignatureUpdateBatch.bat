:: Rewritten the script in batch.
:: Sometimes executing Powershell might be blocked, so batch goes around that.

:: Since sometimes on the first login, the GPO powershell script copies weird files, lets just clear the directory first
:: as otherwise the command below will not overwrite the files, as they'll be different
del "C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Signatures\" /s /f /q

%systemroot%\System32\xcopy /s/z/y "\\FILE_SERVER\Scripts\signature_tool\user_signatures\%USERNAME%" C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Signatures\
