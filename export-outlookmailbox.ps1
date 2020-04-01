#########################
# This script runs under the current user context.
# - Connects to the user mailbox
# - Creates a PST file
# - Ataches it to the mailbox
# - Copies inbox, calendar, contacts into an "exported" folder within the PST file
# - Detaches the PST file
# reference:  https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._store.storeid?view=outlook-pia#Microsoft_Office_Interop_Outlook__Store_StoreID
########################

$strUserName = get-content env:UserName
$strDate = get-date -format dd-MM-yyyy

#target PST
$exportlocation = get-content env:LOCALAPPDATA
$exportFile = $exportlocation+"\"+$strUserName+"_"+$strDate+".PST"

#kill outlook if already running
get-process | where { $_.Name -like "Outlook" }| kill

#init outlook COM object
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI")

#grab mailbox folders to be exported
$inbox = $namespace.getDefaultFolder($olFolders::olFolderInBox)
$calendar = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
$contacts = $namespace.getDefaultFolder($olFolders::olFolderContacts)


#add PST to user mailbox profile
$namespace.AddStore($exportFile)

#Grab PST folder and create an "exported" subfolder.
$exportstore = ($namespace.stores | ? {$_.FilePath -eq $exportFile})
$exportRootFolder = $exportstore.getrootfolder()
$exportFolder = $exportRootFolder.Folders.Add("Exported")

#Copy inbox folder to the PST
$inbox.CopyTo($exportFolder)
$calendar.CopyTo($exportFolder)
$contacts.CopyTo($exportFolder)

#Removestore
$namespace.RemoveStore($exportRootFolder)

#kill outlook
get-process | where { $_.Name -like "Outlook" }| kill
