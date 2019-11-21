' ADAPTER SOURCEPATH A LA MACHINE 

' Pour Generation du nom de fichier incluant la date + time
dim WshShell
dim strSafeDate, strSafeTime, strDateTime

' use a default source path or get one from the command line parameters
dim sourcepath: sourcepath = "C:\Users\docto\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets\"
if WScript.Arguments.Named.Exists("source") then
    sourcepath = WScript.Arguments.Named("source")
end if

' use a default destination path or get one from the command line
dim destinationpath: destinationpath = "C:\Wallpaper\"
if WScript.Arguments.Named.Exists("destination") then
    destinationpath = WScript.Arguments.Named("destination")
end if

' use a default file size limit or get one from the command line
' we accept in kbytes so we convert this to bytes
dim sizelimit: sizelimit = 1000 * 1024 ' default 6000 kbytes
if WScript.Arguments.Named.Exists("sizelimit") then
    sizelimit = WScript.Arguments.Named("sizelimit")
end if

' use a Scripting.FileSystemObject to get the file objects of each file
' in the source directory. The file object has a Size property, which
' has the file size in bytes
dim fso: set fso = CreateObject("Scripting.FileSystemObject")
dim sourcefolder: set sourcefolder = fso.GetFolder(sourcepath)
if not fso.FolderExists(destinationpath) then
     ' we'll throw an error if the path is not found but you could instead
     ' create the directory automatically
     err.raise 1,,destinationpath & " not found"
end if

' Creation de la base du nom de fichier avec date et time
Set WshShell = CreateObject("WScript.Shell")
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
'Set strDateTime equal to a string representation of the current date and time, for use as part of a valid Windows filename
strDateTime = strSafeDate & "-" & strSafeTime

' Creation d'un objet image pour recuperer la resolution
Set objImage = CreateObject("WIA.ImageFile")

' loop through each file in the directory, compare size property against
' the limit and copy as appropriate
dim oldFileName
dim file, count: count = 0
for each file in sourcefolder.Files

		if file.size > sizelimit then
			
			objImage.LoadFile (sourcefolder & "\" & file.name)
			
			' Verification de la largeur
    	if objImage.Width >= 1920 Then
    		' Verification de l'absence d'extension
    		if fso.GetExtensionName(file.Name) = "" then
    			'WScript.Echo(file.Name) 'DEBUG
    			oldFilename = file.name ' Sauvegarde nom d'origine
  				file.name = strDateTime & "-" & (count+1) &".jpg" ' Nouveau nom avec date et index
  			end if
  			count = count + 1
  		
  			' Copie le fichier vers le dossier de destination
  			file.Copy destinationpath
  			
  			' Restaure le nom de fichier original
  			file.Move (sourcefolder & "\" & oldFilename)
			     	
   	 	end if
   	 	
  	end if
  	
next
' Message de fin
msgBox ("Terminé: " & count & " fichier(s) déplacés."),vbInformation,"Fin de traitement"