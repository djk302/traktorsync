'create a public variable to store the tracks from the collection 
'this can be accessed by all subroutines and means we can minimise database calls
Public trackCollection
Public nmlLastUpdate
Public playlistProgress
Public nmlLastModified
Public mmPlaylists
Public traktorCollectionRoot

Sub traktorsync
	
	Dim traktorCollection, bkupLocation, mmRootTitle
	
	
	traktorCollection = "D:\Dave\Documents\Native Instruments\Traktor 2.11.3\collection.nml"
	bkupLocation = "D:\Backups\collection.nml"
	mmRootTitle = "Traktor Collection"
	
	'check if Traktor is currently running
	If IsProcessRunning("Traktor.EXE") Then
		Dim mess : mess = "Traktor is currently running, it is recommended you close Traktor before continuing"
		Select Case SDB.MessageBox(mess,mtWarning,Array(mbIgnore,mbCancel))
				Case mrIgnore
						'do nothing
				Case Else
						Exit Sub
		End Select
	End If
	
	'take a backup of the Traktor collection
	Call CopyFile(traktorCollection, bkupLocation)
	
	Dim i, itm, songPath, traktorList
	
	'create an XML object and load the traktor collection file
	Set xmlDoc = _
	CreateObject("Msxml2.DOMDocument.6.0")
    xmlDoc.setProperty "SelectionLanguage", "XPath"
	xmlDoc.Async = "False"
	xmlDoc.Load(traktorCollection)
	
	Set objFSO = Createobject ("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(traktorCollection)
	nmlLastModified = objFile.DateLastModified
	
	
	'handle an error loading the file
	if xmlDoc.parseError <> 0 Then
		msgbox "there was an error loading Traktor collection"
	end if
	
	'get the location of the playlist node, all playlists are children of this node
	set traktorPlaylistsRoot = xmlDoc.SelectSingleNode("//NML/PLAYLISTS/NODE[@NAME='$ROOT']")
	set traktorCollectionRoot = xmlDoc.SelectSingleNode("//NML/COLLECTION")
	

	'top level node for a track is ENTRY so load all ENTRY nodes into an array so we can iterate through them
	set traktorCollectionTracks = xmlDoc.SelectNodes("//NML/COLLECTION/ENTRY")

	'load all the tracks from the Traktor collection so we can do stuff with them without calling the database each time
	Call LoadTracksFromTraktor(traktorCollectionTracks)
	
	'create the progress object for tracking how far through processing playlists we are 
	Set playlistProgress = SDB.Progress
		
	
	'ask the user if they want to update playlists from Traktor to MediaMonkey.
	mess = "Do you want to add new playlists and update tracks in static playlists from Traktor"
	Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo,mbCancel))
		Case mrYes
			'first delete any playlists from MM that no longer exist in Traktor
			Call removeMMPlaylists(traktorPlaylistsRoot)
			'initialise variables to display progress bar
			set traktorLists = traktorPlaylistsRoot.SelectNodes("//NODE[@TYPE='PLAYLIST']")
			playlistProgress.MaxValue = traktorLists.Length
			playlistProgress.Value = 0
			playlistProgress.Text = SDB.Localize("Processing playlists...  " & playlistProgress.Value & "/" & playlistProgress.MaxValue)
			'run the MapNodes sub and pass the root playlist node 
			Call MapNodes(traktorPlaylistsRoot, "")
			playlistProgress.Value = traktorLists.Length
		Case mrNo
			'do nothing
		Case Else
			Exit Sub
	End Select

	'ask the user if they want to update playlists in Traktor
	mess = "Do you want to add new playlists and update tracks to Traktor?"
	Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo,mbCancel))
		Case mrYes
			set mmRoot = SDB.PlaylistByTitle(mmRootTitle)
			'delet any playlists from Traktor that no longer exist in MM
			Call removeTraktorPlaylists(traktorPlaylistsRoot)
			'initialise variables to display progres bar
			playlistProgress.MaxValue = 0 
			Call countMMPlaylists(mmRoot)
			Call MapPlaylists(mmRoot, traktorPlaylistsRoot, "")
			'save the collection NML
			xmlDoc.Save(traktorCollection)			
		Case mrNo
			'do nothing
		Case Else
			Exit Sub
	End Select
End Sub

Sub countMMPlaylists(mmPlaylist)
	''''Subroutine to count the number of playlists in the Traktor Collection in MM'''''

	If mmPlaylist.ChildPlaylists.Count > 0 Then
		playlistProgress.MaxValue = playlistProgress.MaxValue + mmPlaylist.ChildPlaylists.Count
		playlistProgress.Text = SDB.Localize("Processing playlists...  " & playlistProgress.Value & "/" & playlistProgress.MaxValue)
		For i = 0 to mmPlaylist.ChildPlaylists.Count-1
			Call countMMPlaylists(mmPlaylist.ChildPlaylists.Item(i))
		Next
	End If
End Sub

Function CreateGUID()

	'this function generates the 36 character random string for the playlist UUID

	Dim TypeLib
	
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
	CreateGUID = Mid(TypeLib.Guid, 2, 36)
	CreateGUID = replace(CreateGUID,"-","")
	CreateGUID = LCase(CreateGUID)
	
End Function
	

Function IsProcessRunning(process)

	'this function checks if a process is running on the machine

	Dim objList
	
    Set objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where name='" & process & "'")

    If objList.Count > 0 Then
        IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If
	
End Function


Sub LoadTracksFromTraktor(traktorEntries)
	
	'initialise variables so we can show a progress bar for the operation
	Set Progress = SDB.Progress
	Progress.MaxValue = traktorEntries.Length
	Progress.Value = 0
	Progress.Text = SDB.Localize("Loading Traktor Collection...  " & Progress.Value & "/" & Progress.MaxValue)
	
	'create the database object so we can make database calls
	Dim dat : Set dat = SDB.Database
	Dim sql 
	Dim locationNode
	Dim trax 
	Dim notFoundAction
	Dim notFoundTrack : Set notFoundTrack = SDB.NewSongData
	
	'Initialise a varibale that will allow us to deal with multiple missing tracks through one user response if neccesary
	notFoundAction = 0
	'create the dictionary to store the tracks in using the public variable initialised at the top of the script
	set trackCollection = CreateObject("Scripting.Dictionary")
	trackCollection.CompareMode = vbTextCompare
	
	'create a playlist to add missing tracks from Traktor
	Set notFoundPL = SDB.PlaylistByID(-1).CreateChildPlaylist("Stuff Added From Traktor")
	
	'cycle through all of the tracks in the traktor collection
	for each entry in traktorEntries
		
		'file location data is stored in the location subnode of the entry in the nml
		set locationNode = entry.SelectSingleNode("LOCATION")
		
		'/: is used as a delimiter in nml and the path and file name are stored in seperate attributes so here we construct the windows file path 
		location = replace(locationNode.getAttribute("DIR"),"/:","\") & locationNode.getAttribute("FILE")
		
		'we will use the full file path as stored in nml as our unique key in the dictionary as it should save operations later
		traktorLocation = locationNode.getAttribute("VOLUME") & locationNode.getAttribute("DIR") & locationNode.getAttribute("FILE")
		
		'make sure that any previous database operations have been closed
		Call dat.Commit()
		
		'construct the SQL query to find the song by its windows file path and run the search
		sql = "Songs.SongPath = "":" & location & """"
		Set trax = dat.QuerySongs(sql)
		
		notFound = 0
		
		'MM database query returns a "SongIterator" object containing the Song objects found by the search. As we are searching by file path the result can only be one song
		'If the EOF property is true then we have reached the last Song in the iterator, or in this case, not found any songs as the list is empty
		If trax.EOF Then
			
			'Its possible for the NML to contain tracks that no longer exist so we need to check if the file exits in the file system 
			If SDB.Tools.FileSystem.FileExists(locationNode.getAttribute("VOLUME") & location) Then
				Select Case notFoundAction
					Case 0
						'Ask the user if they want to add the track to Media Monkey 
						mess = "The following track was not found in the MediaMonkey database: " & location & sql & vbCrLf & "Would you like to add this track to MediaMonkey?"						
						Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbYesToAll,mbNo,mbNoToAll))
						Case mrYes
							'get the path for the track not found in MM
							notFoundTrack.Path = locationNode.getAttribute("VOLUME") & location
							'add the new track to MM
							Call updateSongData(notFoundTrack)
							Set trax = dat.QuerySongs(sql)
						Case mrYesToAll
							'set the not found action to 1 so we can add any more tracks without asking the user
							notFoundAction = 1
							'get the path for the track not found in MM
							notFoundTrack.Path = locationNode.getAttribute("VOLUME") & location
							'add the new track to MM
							Call updateSongData(notFoundTrack)
							Set trax = dat.QuerySongs(sql)
						Case mrNo
							'Do nothing
						Case mrNoToAll
							'set the not found action to -1 so we can skip adding any more tracks without asking the user
							notFoundAction = -1
						Case Else
							call dat.commit()
							Exit Sub
						End Select	
					Case 1 'yes to all
						'get the path for the track not found in MM
						notFoundTrack.Path = locationNode.getAttribute("VOLUME") & location
						'add the new track to MM
						Call updateSongData(notFoundTrack)
						Set trax = dat.QuerySongs(sql)
					Case -1 'no to all
						'Do Nothing
					Case Else
						call dat.commit()
						Exit Sub
				End Select	
			Else
				SDB.MessageBox "File does not exist!: " & locationNode.getAttribute("VOLUME") & location & vbCrLf & "The file may have been deleted or moved, you should run a consistency check in Traktor",mtWarning,mbOK
				notFound = 1
			End If
		End If
		
		If notFound = 0 Then
		
			'get the song object out of the songiterator returned by the database search
			set songUpdate = trax.item
							
			'add the track to the dictionary using the full path stored in nml as the key and the MM song ID as the value
			trackCollection.Add traktorLocation, songUpdate.ID
			
			'date added, playcount etc is tored in the INFO node of the ENTRY in NML
			set traktorInfo = entry.selectsinglenode("INFO")
			
			'get the infor we want to update
			ratingString = traktorInfo.getAttribute("RATING")
			playCount = traktorInfo.getAttribute("PLAYCOUNT")
			lastPlay = traktorInfo.getAttribute("LAST_PLAYED")
			traktorAdded = traktorInfo.getAttribute("IMPORT_DATE")
			
			updateNo = 0
			
			if Len(ratingString) > 0 then
				
				IF songUpdate.Custom2 <> ratingString Then
					'	mess = "SongUpdate.Custom2 = " & songUpdate.Custom2 & vbCrLf & "ratingString = " & ratingString				
					'	Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo))
					'	Case mrYes
					'	Case mrNo
					'		Exit Sub
					'	End Select
					songUpdate.Custom2 = ratingString
					updateNo = 1
				End If
				
			end if
			
			
			If NOT IsNull(playCount) Then

				If songUpdate.PlayCounter <> playCount Then
					'	mess = "songUpdate.PlayCounter = " & songUpdate.PlayCounter & vbCrLf & "playCount = " & playCount				
					'	Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo))
					'	Case mrYes
					'	Case mrNo
					'		Exit Sub
					'	End Select
					songUpdate.PlayCounter = playCount
					updateNo = 1
				End If
			
			Else
				If songUpdate.PlayCounter <> 0 Then
					'mess = "songUpdate.PlayCounter = " & songUpdate.PlayCounter & vbCrLf & "playCount = " & playCount				
					'	Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo))
					'	Case mrYes
					'	Case mrNo
					'		Exit Sub
					'	End Select
					songUpdate.PlayCounter = 0
					updateNo = 1
				End If
			
			End If 
			
			If NOT IsNull(lastPlay) Then
			
				lastPlayYear = left(lastPlay,4)
				lastPlayMonth = mid(lastPlay, 6, instr(6,lastPlay,"/")-6)
				lastPlayDay = mid(lastPlay, instr(6,lastPlay,"/")+1)
				If songUpdate.LastPlayed <> lastPlayDay & "/" & lastPlayMonth & "/" & lastPlayYear Then
						'mess = "songUpdate.LastPlayed = " & songUpdate.LastPlayed & vbCrLf & "lastplay = " &  lastPlayDay & "/" & lastPlayMonth & "/" & lastPlayYear			
						'Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo))
						'Case mrYes
						'Case mrNo
						'	Exit Sub
						'End Select
					songUpdate.LastPlayed = lastPlayDay & "/" & lastPlayMonth & "/" & lastPlayYear
					updateNo = 1
				End If
			End If
			
			'update the date the file added to the collection
			If NOT IsNull(traktorAdded) Then
				
				traktorAddedYear = left(traktorAdded,4)
				traktorAddedMonth = mid(traktorAdded, 6, instr(6,traktorAdded,"/")-6)
				traktorAddedDay = mid(traktorAdded, instr(6,traktorAdded,"/")+1)
				If songUpdate.DateAdded <> traktorAddedDay & "/" & traktorAddedMonth & "/" & traktorAddedYear Then
						'mess = "songUpdate.DateAdded = " & songUpdate.DateAdded & vbCrLf & "lastplay = " & traktorAddedDay & "/" & traktorAddedMonth & "/" & traktorAddedYear		
						'Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo))
						'Case mrYes
						'Case mrNo
						'	Exit Sub
						'End Select
					songUpdate.DateAdded = traktorAddedDay & "/" & traktorAddedMonth & "/" & traktorAddedYear
					updateNo = 1
				End If
					
			End If
			
			'update the MM database and write the data to file tags
			If updateNo > 0 Then
				'songUpdate.UpdateDB
				'songUpdate.WriteTags
			End If
		End If
		'update the progress bar
		Progress.Value = Progress.Value + 1
		Progress.Text = SDB.Localize("Loading Traktor Collection...  " & Progress.Value & "/" & Progress.MaxValue)
		If Progress.Terminate then
			Exit For
		End if
	
	Next
	
	'make sure we close any database operation
	Call dat.commit()
	
	

	
End Sub

Function updateSongData(songDataObject)

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''Function to copy song information from file tags into mm database'''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	set tagSongData = songDataObject.GetCopy
	songDataObject.ReadTags
	tagSongData.Title = songDataObject.Title
	tagSongData.ArtistName = songDataObject.ArtistName
	tagSongData.AlbumName = songDataObject.AlbumName
	tagSongData.AlbumArtistName = songDataObject.AlbumArtistName
	tagSongData.Bitrate = songDataObject.Bitrate
	tagSongData.BPM = songDataObject.BPM
	tagSongData.Channels = songDataObject.Channels
	tagSongData.Comment = songDataObject.Comment
	tagSongData.Date = songDataObject.Date
	tagSongData.DiscNumber = songDataObject.DiscNumber
	tagSongData.Genre = songDataObject.Genre
	tagSongData.Lyrics = songDataObject.Lyrics
	tagSongData.Publisher = songDataObject.Publisher
	tagSongData.SampleRate = songDataObject.SampleRate
	tagSongData.SongLength = songDataObject.SongLength
	tagSongData.TrackOrderStr = songDataObject.TrackOrderStr
	tagSongData.Year = songDataObject.Year
	tagSongData.UpdateDB
	
End Function

Function getMMplaylistID(playlistPath)

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''Function to get the MediaMonkey Playlist ID of a Traktor playlist from its path.''''
	'''''Returns a 0 if the ID  does not exist in the config file						 ''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'create the inifile object
	Dim cfg : Set cfg = SDB.IniFile
	Dim list
	
	'check that the value exists in the config file
	
	If cfg.ValueExists("TraktorSync",playlistPath) Then
		
		'get the value of the playlist ID from the config file
		getMMplaylistID = cfg.IntValue("TraktorSync",playlistPath)
			
		'check that the playlist still exists in MediaMonkey, if not we set to 0 as if the ID did not exist in the config file
		'list.id may be -1 if the mediamonkey root playlist is returned so we need to check fo that also
		Set list = SDB.PlaylistByID(getMMplaylistID)
		If list.id <= 0 Then
			getMMplaylistID = 0
		End If
		
	Else
		
		'return a 0 if the ID did not exist in the config file 
		getMMplaylistID = 0
	
	End If
		
End Function	

Function createMMplaylistID(playlistPath, parentID)

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''Function to create an entry in the MediaMonkey config file to store  the MM'''''
	'''''playlist ID of a Traktor playlist by the full path to the Traktor playlist.'''''      
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'create the inifile object
	Dim cfg : Set cfg = SDB.IniFile
	Dim list
	Dim lastNodePos, nameStartPos, nameLen, lastNode
	
	'we need to extract the playlist name from the playlist path we have been passed
	
	lastNodePos = InStrRev(playlistPath,"/")+1
	lastNode = mid(playlistPath,lastNodePos)
	
	
	nameStartPos = InStr(lastNode, "'") +1
	nameLen = len(lastNode) - nameStartPos -1
	
	
	playlistName = mid(lastNode,nameStartPos,nameLen)
	
	
	If playlistName = "$ROOT" Then
		playlistName = "Traktor Collection"
	End If
	
	
	'if the ID does not exist in the config file, create a new playlist and store its ID in the config file
	Set list = SDB.PlaylistByID(parentID).CreateChildPlaylist(playlistName)
	cfg.IntValue("TraktorSync",playlistPath) = list.ID
		
	'return the new playlist ID
	createMMplaylistID = list.ID

End Function

Sub removeMMPlaylists(traktorRootNode)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''' This sub looks through the playlists in MediaMonkey, checks if they still '''''
	''''' exist in Traktor and will remove from MediaMonkey if they have been       '''''
	''''' deleted from Traktor														'''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Set Progress = SDB.Progress

	
	'get the list of playlist paths and associated MediaMonky IDs from the config file
	Dim cfg : Set cfg = SDB.IniFile
	Dim KeyValueList : Set KeyValueList = cfg.Keys("TraktorSync")
	
	Progress.MaxValue = KeyValueList.Count
	Progress.Value = 0
	Progress.Text = SDB.Localize("Checking for playlists that have been removed from Traktor... " & Progress.Value & "/" & Progress.MaxValue)
	
	'Initialise a variable to deal with deleting multiple lists without confirming each time
	Dim delAction : delAction = 0
	
	'cycle through all of the values we have retireved from the config
	For i = 0 To KeyValueList.Count - 1
	
		KeyValue = KeyValueList.Item(i)
		
		'the key/value pair is a single string so we need to extract the the key and value to seperate variables
		playlistPath = Left(KeyValue, InStrRev(KeyValue, "=") - 1)
		playlistID = Mid(KeyValue, InStrRev(KeyValue, "=") + 1)
		
		'the playlist path is stored as the key and this doesnt work properly if there are equals signs
		'so we need to add the = signs back in to the Node name part of the path
		playlistPath = replace(playlistPath, "NODE[@NAME'", "NODE[@NAME='")
		
		'we already select the playlist root elswhere, so can remove that part of the string here
		playlistPathShort = replace(playlistPath, "/NODE[@NAME='$ROOT']" , "")
		
		'if the playlist pah is empty after removing the root section, then we are at the root and dont need to do anything
		If playlistPathShort <> "" Then
			
			'searchthe nml document for a playlist node with the path we have
			set listnode = traktorRootNode.SelectSingleNode("/" & playlistPathShort)
			'if the playlist has been deleted the search will return nothing
			If listnode is Nothing Then
				'use the corresponding MM ID of the playlist we couldnt find to select the playlist in mediamonkey
				Set listToDelete = SDB.PlaylistByID(playlistID)
				'check that the mediamonky root playlist hasnt been returned (we definately dont want to delete that!)
				If listToDelete.ID > 0 Then
					If delAction = 0 Then
						'Ask the user if they want to delete the playlist from MM and if they want to delete all other playlists that have been removed from Traktor 
						mess = "Playlist " & listToDelete.Title & " has been removed from Traktor, do you want to delete it from MediaMonkey?"
						Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbYesToAll,mbNo,mbNoToAll))
							Case mrYes
								delAction =  -1
							Case mrYesToAll
								delAction =  -2
							Case mrNo
								'do nothing
							Case mrNoToAll
								Exit Sub
						End Select
					End If
					
					If delAction < 0 Then
						'delete the playlist and its associated key in the config file
						call listToDelete.Delete
						cfg.DeleteKey "TraktorSync", playlistPath
						'if the user didnt select Yes To All reset the delete action variable
						If delAction = -1 Then
							delAction = 0
						End if
					End if
				End if
			End If
		End If 
	Progress.Value = Progress.Value + 1
	Progress.Text = SDB.Localize("Checking for playlists that have been removed from Traktor... " & Progress.Value & "/" & Progress.MaxValue)
	Next 
End Sub

Sub removeTraktorPlaylists(traktorPlaylistsRoot)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''' This sub removes playlists deleted from MediaMonkey from the Traktor NML  '''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Set Progress = SDB.Progress

	
	'get the list of playlist paths and associated MediaMonky IDs from the config file
	Dim cfg : Set cfg = SDB.IniFile
	Dim KeyValueList : Set KeyValueList = cfg.Keys("TraktorSync")
	
	Progress.MaxValue = KeyValueList.Count
	Progress.Value = 0
	Progress.Text = SDB.Localize("Checking for playlists that have been removed from MediaMonkey... " & Progress.Value & "/" & Progress.MaxValue)
	
	'Initialise a variable to deal with deleting multiple lists without confirming each time
	Dim delAction : delAction = 0
	
	'cycle through all of the values we have retireved from the config
	For i = 0 To KeyValueList.Count - 1
	
		KeyValue = KeyValueList.Item(i)
		
		'the key/value pair is a single string so we need to extract the the key and value to seperate variables
		playlistID = Mid(KeyValue, InStrRev(KeyValue, "=") + 1)
		playlistPath = Left(KeyValue, InStrRev(KeyValue, "=") - 1)
		'try to find the playlist by its ID
		Set list = SDB.PlaylistByID(playlistID)
		If list.id <= 0 Then
			If delAction = 0 Then
				'Ask the user if they want to delete the playlist from MM and if they want to delete all other playlists that have been removed from Traktor 
				mess = "Playlist " & lplaylistPath & " has been removed from Traktor, do you want to delete it from MediaMonkey?"
				Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbYesToAll,mbNo,mbNoToAll))
					Case mrYes
						delAction =  -1
					Case mrYesToAll
						delAction =  -2
					Case mrNo
						'do nothing
					Case mrNoToAll
						Exit Sub
				End Select
			End If
			
			If delAction < 0 Then
				'delete the playlist and its associated key in the config file
				'the playlist path is stored as the key and this doesnt work properly if there are equals signs
				'so we need to add the = signs back in to the Node name part of the path
				playlistPath = replace(playlistPath, "NODE[@NAME'", "NODE[@NAME='")
	
				'we already select the playlist root elswhere, so can remove that part of the string here
				playlistPathShort = replace(playlistPath, "/NODE[@NAME='$ROOT']" , "")
					
				set nodeToDel = traktorPlaylistsRoot.SelectSingleNode(playlistPathShort)
				if nodeToDel is Nothing Then
					'do nothing
				Else
					set nodeToDelParent = nodeToDel.parentNode
					Call nodeToDelParent.removeChild(nodeToDel)
					noToDelParent.setAttribute("COUNT") = noToDelParent.getAttribute("COUNT") - 1
				End If
				cfg.DeleteKey "TraktorSync", playlistPath
				'if the user didnt select Yes To All reset the delete action variable
				If delAction = -1 Then
					delAction = 0
				End if
			End if		
		End If
		
	Progress.Value = Progress.Value + 1
	Progress.Text = SDB.Localize("Checking for playlists that have been removed from MediaMonkey... " & Progress.Value & "/" & Progress.MaxValue)
		
	Next

End Sub
	
Sub updateTraktorToMM(nmlNode, parentPath, currentPath)
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''Subroutine that reads playlists in Traktor NML   '''''' 
	'''''and creates or updates playlists in MediaMonkey  ''''''                      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	
	'Loops and Recordings are special hidden playlists in the Traktor collection and we need to ignore them 
	Select Case nmlNode.getAttribute("NAME")
			Case "_LOOPS"
				Exit Sub
			Case "_RECORDINGS"
				Exit Sub
	End Select
	
	'get the MediaMonky playlist ID from the config file, will return a 0 if it doesnt exist
	currentPlaylistID = getMMplaylistID(currentPath)

	'If the playlist doesnt exist then create it 
	If currentPlaylistID = 0 Then
		'get the ID of the parent playlist
		parentPlaylistID = getMMplaylistID(parentPath)
		'create the playlist as a child of the parent
		currentPlaylistID = createMMplaylistID(currentPath, parentPlaylistID)
	End If
	
	'if the curent playlist is a Folder in Traktor then it will contain no tracks and we can exit here
	If nmlNode.getAttribute("TYPE") = "FOLDER" Then
		Exit Sub	
	End If
	
	'load the playlist object
	Dim list : Set list = SDB.PlaylistByID(currentPlaylistID)
	
	'update progress variables
	playlistProgress.Value = playlistProgress.Value + 1
	playlistProgress.Text = SDB.Localize("Processing playlists...  " & list.Title & " : " & playlistProgress.Value & "/" & playlistProgress.MaxValue)
	
	
	'if the playlist is an AutoPlaylist then we cant add tracks to it from Traktor
	If list.isAutoplaylist Then
		Exit Sub
	End If
	
	'select the Playlist node in the nml structure
	Dim traktorTrackList : set traktorTrackList = nmlNode.SelectSingleNode("./PLAYLIST")
	
	'initialise variables to handle the progress bar
	Set Progress = SDB.Progress
	Progress.MaxValue = traktorTracklist.getAttribute("ENTRIES")
	Progress.Value = 0
	
	'check if the playlist in Traktor has tracks in it
	If traktorTrackList.HasChildNodes() Then
	
		'Check if the playlist in MediaMonkey is empty, if so we need to build from scratch
		Set mmTrackList = list.Tracks		
		If mmTrackList.Count = 0 Then
			'Tracks are stored by their path in a node called primary key in nml
			For Each Track in traktorTrackList.SelectNodes("./ENTRY/PRIMARYKEY")
				Progress.Text = SDB.Localize("Creatinng " & list.Title & "... " & Progress.Value & "/" & Progress.MaxValue)
				'get the path of the track
				trackPath = Track.getAttribute("KEY")
				'check if the track exists in the MM database using the dictionary of tracks we built when loading the traktor collection
				If trackCollection.Exists(trackPath) Then
					'get the corresponding MM ID
					trackID = trackCollection.Item(trackPath)
					'add the track to the playlist
					list.AddTrackById(trackID)
				Else
					'handle non existant tracks
					msgbox "Couldnt Find " & trackPath
				End If
				
				'update progress value
				Progress.Value = Progress.Value + 1
			Next
		Else
			
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''If the playlist in MM is not empty then we need to work out any differences between the lists using dictionary objects ''
			''We create a dictionary to contain the tracks in traktor, one dictionary to contain the tracks in MM and one to contain ''
			''the differences. First run through the Traktor tracks and add all items to the Traktor and Difference dictionaries.    ''
			''Then we run through the tracks in MM adding to the MM dictionary and REMOVING from the Difference dictionary if they   ''
			''also exist in Traktor and removing straight from the MM playlist if they dont exist in Traktor. Any tracks that remain ''
			''in the Difference dictionary then get added to the MM playlist														 ''
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			'initalise the dictionaries making sure we use textual, not binary comparison to avoid issues with case matching
			Dim traktorDict : Set traktorDict = CreateObject( "Scripting.Dictionary" )
			traktorDict.CompareMode = vbTextCompare
			Dim differenceDict : Set differenceDict = CreateObject( "Scripting.Dictionary" )
			differenceDict.CompareMode = vbTextCompare
			Dim mmDict : Set mmDict = CreateObject( "Scripting.Dictionary" )
			mmDict.CompareMode = vbTextCompare
			
			'We need an array to store the position of each trak in the list so we can update any changes to the order
			Dim traktorTLLength  
			traktorTLLength = traktorTrackList.SelectNodes("./ENTRY/PRIMARYKEY").Length
			Dim posArray()
			ReDim posArray(traktorTLLength)
			playlistPos = 0
			
			'Here we run through all of the tracks in Traktor and add them to both the Traktor and Difference dictionaries
			For Each Track in traktorTrackList.SelectNodes("./ENTRY/PRIMARYKEY")
				
				Progress.Text = SDB.Localize("Updating " & list.Title)
				'get the path of the track
				trackPath = Track.getAttribute("KEY")
				
				If trackCollection.Exists(trackPath) Then
					'get the corresponding MM ID
					trackID = trackCollection.Item(trackPath)
					'add the track to the array, playlists can have duplicates but dictionaries cant so have to deal with that
					If Not traktorDict.Exists(trackID) Then
						traktorDict.Add trackID, 1
						differenceDict.Add trackID, 1
						posArray(playlistPos) = trackID
					Else
						traktorDict.Item(trackID) = traktorDict.Item(trackID) + 1
						differenceDict.Item(trackID) = differenceDict.Item(trackID) + 1
						posArray(playlistPos) = trackID
					End If
				Else
					msgbox "Couldnt Find " & trackPath
				End If
				
				playlistPos = playlistPos + 1
				
			Next
			
			
			'Now we run through the MM playlist
			For i = 0 to mmTrackList.Count - 1
				Set mmSong = mmTrackList.Item(i)
				'if the track doesnt exist in Traktor remove it from the MM playlist
				If Not traktorDict.Exists(mmSong.ID) Then 
					list.RemoveTrackNoConfirmation(mmSong)
				'if there are more copies of the track in the MM playlist than in Traktor, remove 1.
				ElseIf mmDict.Item(mmSong.ID) + 1 > traktorDict(mmSong.ID) Then
					list.RemoveTrackNoConfirmation(mmSong)
				'Otherwise we need to add it to the MM Dictionary and remove from the difference dictionary
				'We are also handling the fact that playlists can have duplicate entries
				ElseIf Not mmDict.Exists(mmSong.ID) Then
					mmDict.Add mmSong.ID,1
					If differenceDict.Item(mmSong.ID) > 1 Then
						differenceDict.Item(mmSong.ID) = differenceDict.Item(mmSong.ID) - 1
					Else
						differenceDict.Remove(mmSong.ID)
					End If	
				Else
					mmDict.Item(mmSong.ID) = mmDict.Item(mmSong.ID) + 1
					If differenceDict.Item(mmSong.ID) > 1 Then
						differenceDict.Item(mmSong.ID) = differenceDict.Item(mmSong.ID) - 1
					Else
						differenceDict.Remove(mmSong.ID)
					End If
				End If
			Next
				
			'Now we run through any tracks in the difference dictionary and add them to the MM playlist	
			If differenceDict.Count > 0 Then
				For each item in differenceDict.keys
					If differenceDict.Item(item) > 1 Then
						For i = 1 to differenceDict.Item(item)
							list.AddTrackByID(item)
						Next
					Else	
						list.AddTrackByID(item)
					End If
				Next
			End If
			
			'Ensure that mmTrackList variable conatins the updates we have just made
			Set mmTrackList = list.Tracks
			
			'Now we run through the tracks in the playlist, check their position and move if neccesary
			For i = 0 to mmTrackList.Count - 1
				set mmSong = mmTrackList.Item(i)
				If mmSong.ID <> posArray(i) Then
					If mmSong.ID <> posArray(i+1) Then
						For j = i to mmTrackList.Count - 1
							set mmSong = mmTrackList.Item(j)
							if mmSong.ID = posArray(i+1) Then
								Exit For
							End If
						Next
					End If
					For j = i to mmTrackList.Count - 1
						set mmSong2 = mmTrackList.Item(j)
						if mmSong2.ID = posArray(i) Then
							list.MoveTrack mmSong2, mmSong
							Exit For
						End If
					Next
				End If
				Progress.Value = Progress.Value + 1
			Next
		End if
	
	End If
	
	
End Sub
	
Sub MapNodes(CurrentNode, playlistPath)
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''Recursive sub that takes a Traktor collection node as input'''''   
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
  
	Dim Node, strText, parNodeName, cnt
	
	
	'check that we have been passed an object!	
	If IsObject(CurrentNode) Then
	
	
		cnt = 1
	
		'In Traktor nml structure if a playlist is a folder it has an child node called SUBNODES
		'which just contains the count of the number of subnodes so we need to deal with that as  
		'it sits between our parent and child playlists
		If CurrentNode.nodeName = "SUBNODES" Then
			'get the name of the parent playlist of the SUBNODE
			parNodeName = CurrentNode.parentNode.getAttribute("NAME")
		
			'In traktor nml the top level playlist is called $ROOT so we need to map this to the top level 
			'playlist in MediaMonkey
			if parNodeName = "$ROOT" Then
				parNodeName = "Traktor Collection2"
			end if	
		
			'Get the number of child playlists from the COUNT attribute of the SUBNODES node
			cnt = CurrentNode.getAttribute("COUNT")
		
			newPlaylistPath = playlistPath & "/SUBNODES"
			'if there are any child playlists then select the first child, otherwise we can exit the sub here
			If cnt > 0 Then
				set CurrentNode = CurrentNode.firstChild
			Else
				Exit Sub
			End If
		
		End If
	
		'If we are at the top level of the Traktor playlist structure then we can set the parent node to be empty
		If CurrentNode.getAttribute("NAME") = "$ROOT" Then
			parNodeName = ""
		End If
	
		'if there is only one playlist then run through this once otherwise do for all playlists
		If cnt = 1 then
		
			thisPlaylistPath = newPlaylistPath & "/NODE[@NAME'" & CurrentNode.getAttribute("NAME") & "']"
			call updateTraktorToMM(CurrentNode, playlistPath, thisPlaylistPath)
		
			'If the current node is a folder then we need to call this sub again to move down the nml structure.
			If CurrentNode.getAttribute("TYPE") = "FOLDER" Then
				call MapNodes(CurrentNode.firstChild, thisPlaylistPath)
			End If
		
	
		Else 
		
			For i = 1 To cnt
	 
				thisPlaylistPath = newPlaylistPath & "/NODE[@NAME'" & CurrentNode.getAttribute("NAME") & "']"
				call updateTraktorToMM(CurrentNode, playlistPath, thisPlaylistPath)
			
				If CurrentNode.getAttribute("TYPE") = "FOLDER" Then
					call MapNodes(CurrentNode.firstChild, thisPlaylistPath)
				End If
						
				set CurrentNode = CurrentNode.nextSibling
				thisPlaylistPath = newPlaylistPath
		
			Next
	
		End If
	    
	End If
  
End Sub

Sub MapPlaylists(mmpList, nmlNode, playlistPath)
		
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	''' Subroutine that looks through MediaMonkey   ''''
	''' playlist structure and pushes it to Traktor ''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'initialise variable to store our NML node
	Dim newNode : Set newNode = Nothing
	
	'handle the root of the Traktor Collection
	If mmpList.Title = "Traktor Collection" Then
		playlistPath = ""
	End If
	
	'Check if the current playlist has children
	If mmpList.ChildPlaylists.Count > 0 Then
	
		'run through all the child playlists
		For i=0 to mmpList.ChildPlaylists.Count-1
			
			set pList = mmpList.ChildPlaylists.Item(i)
			'construct the NML path the the playlist entry in collection
			thisPlaylistPath = playlistPath & "/SUBNODES/NODE[@NAME='" & replace(pList.Title,"'","") & "']"

			'select the playlist in NML using the path
			set newNode = nmlNode.SelectSingleNode("/" & thisPlaylistPath)
			
			'if the playlist doesnt exist then it wont be found and newNode will be nothing
			If newNode is Nothing Then
				'call the subroutine to create the playlist
				call createTraktorPlaylist(pList.Title, pList.ChildPlaylists.Count, nmlNode)
				'store the NML path and corresponding ID in the config file
				SDB.Inifile.IntValue("TraktorSync","/NODE[@NAME'$ROOT']" & replace(thisPlaylistPath,"=","")) = pList.ID
				'select the newly created playlist in the NML
				set newNode = nmlNode.SelectSingleNode("/" & thisPlaylistPath)
				'just in case something goes wrong
				If newNode is Nothing Then
					msgbox "Error occured creating playlist: " & pList.Title
				End If
			End If
			
			'if the NML node is a playlist (rather than a folder) then we need to call the sub that updates the tracks
			If newNode.getAttribute("TYPE")="PLAYLIST" Then
				call updateMMtoTraktor(pList, newNode)
			End If
			
			'If the current playlist has child lists then we need to call this sub again
			If pList.ChildPlaylists.Count > 0 Then			
				call MapPlaylists(pList, newNode, thisPlaylistPath)		
			end if
		Next
	
	'if the current playlist has no children	
	Else
		'construct the NML path the the playlist entry in collection
		thisPlaylistPath = playlistPath & "/SUBNODES/NODE[@NAME='" & mmpList.Title & "']"
		
		'select the playlist in NML using the path
		set newNode = nmlNode.SelectSingleNode(thisPlaylistPath)
		
		'if the playlist doesnt exist then it wont be found and newNode will be nothing
		If newNode = Nothing Then
				'call the subroutine to create the playlist
				call createTraktorPlaylist(pList.Title, pList.ChildPlaylists.Count, nmlNode)
				'store the NML path and corresponding ID in the config file
				SDB.Inifile.IntValue("TraktorSync","/NODE[@NAME'$ROOT']" & replace(thisPlaylistPath,"=","")) = pList.ID
				set newNode = nmlNode.SelectSingleNode("/" & thisPlaylistPath)
				'select the newly created playlist in the NML
				set newNode = nmlNode.SelectSingleNode("/" & thisPlaylistPath)
				'just in case something goes wrong
				If newNode is Nothing Then
					msgbox "Error occured creating playlist: " & pList.Title
				End If
		End If
		
		'if the NML node is a playlist (rather than a folder) then we need to call the sub that updates the tracks
		If newNode.getAttribute("TYPE")="PLAYLIST" Then
				call updateMMtoTraktor(pList, newNode)
		End If
		
	End If	
	
End Sub

sub createTraktorPlaylist(listName, noChildLists, parNode)
	
	''''''''''''''''''''''''''''''''''''''''''''''
	''' Subroutine that creates an empty 	   '''
	''' playlist in NML and adds it to Traktor '''
	''''''''''''''''''''''''''''''''''''''''''''''
	
	Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")
	'remove special characters
	listName = replace(listName,"'","")	
	
	'if the parent is currently a playlist not a folder then we need to convert, 
	'will remove all the tracks as Traktor cannot have tracks in a playlist folder (mediamonkey can)
	parNodeType = parNode.getAttribute("TYPE")
	if parNodeType = "PLAYLIST" Then
		set nodeToDel = parNode.SelectSingleNode("PLAYLIST")
		if nodeToDel is nothing then
		'do something sensible if things do go so well!
		 msgbox "there was an error creating the playlist"
		 Exit Sub
		end if
		
		'removechildnodes
		'set node type to be folder
		'add subnodes node, set count attribute to be 0
		call parNode.removeChild(nodeToDel)
		call parNode.setAttribute("TYPE","FOLDER")
		set newNode = xmlDoc.createElement("SUBNODES")
		call newNode.setAttribute("COUNT",0)
		call parNode.appendChild(newNode)
		
	end if
	
	'create the nml entry for the playlist
	set subNode = parNode.SelectSingleNode("SUBNODES")
	call subNode.setAttribute("COUNT",subNode.getAttribute("COUNT")+1)
	set newNode =  xmlDoc.createElement("NODE")
	'if there are no child playlists then we create a PLAYLIST, else we need to create a FOLDER
	if noChildLists > 0 Then
		call newNode.setAttribute("TYPE","FOLDER")
		call newNode.setAttribute("NAME",listName)
		call subNode.appendChild(newNode)
		set newNode = subNode.SelectSingleNode("NODE[@NAME='" & listName & "']")
		set newSubNode = xmlDoc.createElement("SUBNODES")
		call newSubNode.setAttribute("COUNT",0)
		call newNode.appendChild(newSubNode)
	else
		call newNode.setAttribute("TYPE","PLAYLIST")
		call newNode.setAttribute("NAME",listName)
		call subNode.appendChild(newNode)
		set newNode = subNode.SelectSingleNode("NODE[@NAME='" & listName & "']")
		set newSubNode = xmlDoc.createElement("PLAYLIST")
		call newSubNode.setAttribute("ENTRIES",0)
		call newSubNode.setAttribute("TYPE","LIST")
		call newSubNode.setAttribute("UUID",CreateGUID())
		call newNode.appendChild(newSubNode)
	end if
	
end sub

Sub updateMMtoTraktor(mmPlist, nmlNode)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Subroutine to update a Traktor playlist from MediaMonkey '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	Set xmlDoc = _
	CreateObject("Msxml2.DOMDocument.6.0")
	
	'update global progress variables
	playlistProgress.Value = playlistProgress.Value + 1
	playlistProgress.Text = SDB.Localize("Processing playlists...  " & mmPlist.Title & " : " & playlistProgress.Value & "/" & playlistProgress.MaxValue)
	
	'get the tracks already in the Traktor playlist
	set traktorList = nmlNode.SelectSingleNode("PLAYLIST")
	set traktorTracks = traktorList.SelectNodes("ENTRY")
	
	'get the tracks in the media monkey playlist
	set mmTracks = mmPlist.Tracks
	
	'if the playlist in mediamonkey has no tracks we can remove all tracks in traktor and exit here
	If mmTracks.Count = 0 Then
		If traktorTracks.length > 0 Then
			For Each node In traktorTracks
				traktorList.removeChild(node)
			Next
		End If
		Exit Sub
	End If
	
	'if the traktor playlist is empty then we can build it from scratch and exit here
	If traktorTracks.length = 0 Then
		
		'iterate through all the tracks in mediamonkey
		for j = 0 to mmTracks.Count - 1
			
			Set itm = mmTracks.Item(j)
			songPath = itm.Path
			songPath = replace(songPath,"\","/:")
			'check if the current track from MM exists in the Traktor collection
			If Not trackCollection.Exists(songPath) Then
				'if the track isnt in the collection we need to add it otherwise it will be lost from the playlist when traktor loads the nml
				Call addTrackToTraktor(itm)
			End If
			'set the nml for the new track entry and add to the playlist
			set newEntry = xmlDoc.createElement("ENTRY")
			set newPrimaryKey = xmlDoc.createElement("PRIMARYKEY")
			call newPrimaryKey.setAttribute("TYPE","TRACK")
			call newPrimaryKey.setAttribute("KEY",songPath)		
			call newEntry.appendChild(newPrimaryKey)
			call traktorList.appendChild(newEntry)
			'increase the entries count
			traktorList.setAttribute("ENTRIES") = traktorList.getAttribute("ENTRIES") + 1
		
		next
		
		Exit Sub
	
	End If
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' If the playlist in Traktor is not empty then we need to work out any differences between the lists using dictionary objects.   ''
	'' We create a dictionary to contain the tracks in MM, one dictionary to contain the tracks in Traktor and one to contain the     ''
	'' differences. First run through the MM tracks and add all items to the MM and Difference dictionaries. Then we run through the  '' 
	'' tracks in Traktor adding to the Traktor dictionary and REMOVING from the Difference dictionary if they  also exist in MM and   ''
	'' removing straight from the Traktor playlist if they dont exist in MM. Any tracks that remain in the Difference dictionary then '' 
	'' get added to the Traktor playlist														 									  ''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	'initalise the dictionaries making sure we use textual, not binary comparison to avoid issues with case matching
	Dim traktorDict : Set traktorDict = CreateObject( "Scripting.Dictionary" )
	traktorDict.CompareMode = vbTextCompare
	Dim differenceDict : Set differenceDict = CreateObject( "Scripting.Dictionary" )
	differenceDict.CompareMode = vbTextCompare
	Dim mmDict : Set mmDict = CreateObject( "Scripting.Dictionary" )
	mmDict.CompareMode = vbTextCompare
	
	'We need an array to store the position of each trak in the list so we can update any changes to the order
	Dim posArray()
	ReDim posArray(mmTracks.Count)		
	playlistPos = 0
	
	'Here we run through all of the tracks in MM and add them to both the MM and Difference dictionaries
	For i = 0 to mmTracks.Count - 1
		
		'get the path to the file
		Set itm = mmTracks.Item(i)
		songPath = itm.Path
		'translate to the NML format
		songPath = replace(songPath,"\","/:")
		
		'if the track doesnt exist in Traktor collection we need to add it
		If Not trackCollection.Exists(songPath) Then
			Call addTrackToTraktor(itm)
		End If
		
		'add the track to the array, playlists can have duplicates but dictionaries cant so have to deal with that
		If Not mmDict.Exists(songPath) Then
			mmDict.Add songPath, 1
			differenceDict.Add songPath, 1
			posArray(playlistPos) = songPath
		Else
			mmDict.Item(songPath) = mmDict.Item(songPath) + 1
			differenceDict.Item(songPath) = differenceDict.Item(songPath) + 1
			posArray(playlistPos) = songPath
		End If
		
		playlistPos = playlistPos + 1
	
	Next
	
	
	'Here we run through the tracks in Traktor playlist
	For Each entry In traktorTracks
		
		songPath = entry.SelectSingleNode("PRIMARYKEY").getAttribute("KEY")
		
		'if the track doesnt exist in the MM dictionary, remove it from the Traktor playlist
		If Not mmDict.Exists(songPath) Then
			traktorList.removeChild(entry)
		'if there are more copies of the track in Traktor than MM, remove 1 from Traktor
		ElseIf traktorDict.Item(songPath) + 1 > mmDict(songPath) Then
			traktorList.removeChild(entry)
		'Otherwise add to the Traktor dictionary and remove from the differences
		'we are also handling the fact that playlists can have duplicates here 
		ElseIf Not traktorDict.Exists(songPath) Then
			traktorDict.Add songPath,1
			If differenceDict.Item(songPath) > 1 Then
				differenceDict.Item(songPath) = differenceDict.Item(songPath) - 1
			Else
				differenceDict.Remove(songPath)
			End If
		Else
			traktorDict.Item(songPath) = traktorDict.Item(songPath) + 1
			If differenceDict.Item(songPath) > 1 Then
				differenceDict.Item(songPath) = differenceDict.Item(songPath) - 1
			Else
				differenceDict.Remove(songPath)
			End If
		End If

	Next


	'here we run through the differences and add them to the Traktor playlist
	If differenceDict.Count > 0 Then
		For each item in differenceDict.keys
			If differenceDict.Item(item) > 1 Then
				For i = 1 to differenceDict.Item(item)
					set newEntry = xmlDoc.createElement("ENTRY")
					set newPrimaryKey = xmlDoc.createElement("PRIMARYKEY")
					call newPrimaryKey.setAttribute("TYPE","TRACK")
					call newPrimaryKey.setAttribute("KEY",item)		
					call newEntry.appendChild(newPrimaryKey)
					call traktorList.appendChild(newEntry)
					traktorList.setAttribute("ENTRIES") = traktorList.getAttribute("ENTRIES") + 1
					If tracktorDict.Exists(item) Then
						tracktorDict.ITem(item) = tracktorDict.ITem(item) + 1
					Else
						tracktorDict.Add item, 1
					End If
				Next
			Else	
					set newEntry = xmlDoc.createElement("ENTRY")
					set newPrimaryKey = xmlDoc.createElement("PRIMARYKEY")
					call newPrimaryKey.setAttribute("TYPE","TRACK")
					call newPrimaryKey.setAttribute("KEY",item)		
					call newEntry.appendChild(newPrimaryKey)
					call traktorList.appendChild(newEntry)
					traktorList.setAttribute("ENTRIES") = traktorList.getAttribute("ENTRIES") + 1
					traktorDict.Add item, 1
			End If
		Next
	End If
	
	'reselect the list of tracks to ensure we have all the changes
	set traktorTracks = traktorList.SelectNodes("ENTRY")
	
	'now we run through the Traktor playlist and compare track positions with the array of track positions created from the MM playlist
	'if any tracks are in the wrong position move them to the right place
	For i = 0 to traktorTracks.length - 1
	
		set track = traktorTracks.Item(i)
		set pKey = track.selectSingleNode("PRIMARYKEY")
		
		If pKey.getAttribute("KEY") <> posArray(i) Then
			For j = i to traktorTracks.length - 1
				set track2 = traktorTracks.Item(j)
				set pKey2 = track2.selectSingleNode("PRIMARYKEY")
				if pKey2.getAttribute("KEY") = posArray(i) Then
					call traktorList.insertBefore(track2, track)
					Exit For
				End If
			Next
		End If
	Next

End Sub

Sub addTrackToTraktor(track)
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''
	'''Subroutine to create a new track entry in nml'''
	'''''''''''''''''''''''''''''''''''''''''''''''''''

	Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")
	'Create the new track entry
	Set newTrackEntry = xmlDoc.createElement("ENTRY")


	'first data in NML Entry is file modified info, stored in a different format than windows file system
	Createobject ("Scripting.FileSystemObject")
	Set fso = Createobject ("Scripting.FileSystemObject")          
	Set objFile = fso.GetFile(track.Path)
	fileModified = objFile.DateLastModified
	modified_date = year(fileModified) & "/" & month(fileModified) & "/" & day(fileModified)
	'time stored in seconds from midnight UTC in NML so calculate that
	modified_time = (hour(fileModified)*3600) + (minute(fileModified)*60) + second(fileModified)
	
	'Set the attributes for the ENTRY element
	Call newTrackEntry.setAttribute("MODIFIED_DATE",modified_date)
	Call newTrackEntry.setAttribute("MODIFIED_TIME",modified_time)
	Call newTrackEntry.setAttribute("TITLE",track.Title)
	Call newTrackEntry.setAttribute("ARTIST",track.ArtistName)
	
	'First child element is track location details
	Set newTrackLocation =xmlDoc.createElement("LOCATION")
	
	'we need various details about the file to set the attributes
	driveLetter = left(track.Path,2)
	filePos = InstrRev(track.Path,"\")+1
	location = replace(mid(track.Path,3,filePos-3),"\","/:")
	fileName = mid(track.Path,filePos)

	'set attributes
	Call newTrackLocation.setAttribute("DIR",location)
	Call newTrackLocation.setAttribute("FILE",fileName)
	Call newTrackLocation.setAttribute("VOLUME",driveLetter)
	'add the element as a child of the main ENTRY element
	Call newTrackEntry.AppendChild(newTrackLocation)
	
	'next child element is Album details
	Set newTrackAlbum = xmlDoc.createElement("ALBUM")
	
	'set the attributes
	Call newTrackAlbum.setAttribute("TRACK",track.TrackOrderStr)
	Call newTrackAlbum.setAttribute("TITLE",track.AlbumName)
	'add the element as a child of the main ENTRY element
	Call newTrackEntry.AppendChild(newTrackAlbum)
	
	'next element is modification data
	Set newTrackModInfo = xmlDoc.createElement("MODIFICATION_INFO")
	'set the one attribute
	Call newTrackModInfo.setAttribute("AUTHOR_TYPE", "Importer")
	'add element
	Call newTrackEntry.AppendChild(newTrackModInfo)
	
	'next element is track INFO
	Set newTrackInfo = xmlDoc.createElement("INFO")
	'set attributes
	Call newTrackInfo.setAttribute("BITRATE",track.Bitrate)
	Call newTrackInfo.setAttribute("GENRE",track.Genre)
	Call newTrackInfo.setAttribute("LABEL",track.Publisher)
	Call newTrackInfo.setAttribute("COMMENT",track.Comment)
	Call newTrackInfo.setAttribute("PLAYCOUNT",track.Playcounter)
	Call newTrackInfo.setAttribute("PLAYTIME",track.SongLength)
	Call newTrackInfo.setAttribute("RANKING",round((255/100)*track.Rating))
	Call newTrackInfo.setAttribute("IMPORT_DATE",Year(Now) & "/" & Month(Now) & "/" & Day(Now))
	If track.Playcounter > 0 Then
		Call newTrackInfo.setAttribute("LAST_PLAYED",Year(track.LastPlayed) & "/" & Month(track.LastPlayed) & "/" & Day(track.LastPlayed))
	End If
	Call newTrackInfo.setAttribute("RELEASE_DATE",track.Year & "/" & track.Month & "/" & track.Day)
	'add element as child of main ENTRY
	Call newTrackEntry.AppendChild(newTrackInfo)

	'increase total tracks by 1
	totalEntries = traktorCollectionRoot.getAttribute("ENTRIES")
	Call traktorCollectionRoot.setAttribute("ENTRIES",totalEntries+1)
	'add new ENTRY to the collection
	Call traktorCollectionRoot.AppendChild(newTrackEntry)
	
End Sub



Sub UpdateTracksFromTraktor(traktorEntries)
	
	Set Progress = SDB.Progress
	Progress.MaxValue = traktorEntries.Length
	Progress.Value = 0
	Progress.Text = SDB.Localize("Updating Song Data...")
	for each entry in traktorEntries
	
	set locationNode = entry.selectsinglenode("LOCATION")
	
	location = replace(locationNode.getAttribute("DIR"),"/:","\") & locationNode.getAttribute("FILE")
	
	
	Dim indx : Set indx = SDB.NewSongList
	Set dat = SDB.Database
	Call dat.commit()
	
	Dim sql : sql = "Songs.SongPath = "":" & location & """"
	Dim trax : Set trax = dat.QuerySongs(sql)
	
	If trax.EOF Then
		mess = "Not Found: " & location & sql
		
		Select Case SDB.MessageBox(mess,mtConfirmation,Array(mbYes,mbNo,mbCancel))
			Case mrYes
	'		
		Case mrNo
			Exit Sub
		Case Else
			Exit Sub
		End Select
		
	Else
		
		set songUpdate = trax.item
		
		set traktorInfo = entry.selectsinglenode("INFO")
		
		ratingString = traktorInfo.getAttribute("RATING")
		playCount = traktorInfo.getAttribute("PLAYCOUNT")
		lastPlay = traktorInfo.getAttribute("LAST_PLAYED")
		traktorAdded = traktorInfo.getAttribute("IMPORT_DATE")
		
		
		if Len(ratingString) > 0 then
			
			songUpdate.Custom2 = ratingString
			
		end if
		
		
		If NOT IsNull(playCount) Then

			songUpdate.PlayCounter = playCount
		
		Else
		
			songUpdate.PlayCounter = 0
		
		End If 
		
		If NOT IsNull(lastPlay) Then
		
			lastPlayYear = left(lastPlay,4)
			lastPlayMonth = mid(lastPlay, 6, instr(6,lastPlay,"/")-6)
			lastPlayDay = mid(lastPlay, instr(6,lastPlay,"/")+1)
			songUpdate.LastPlayed = lastPlayDay & "/" & lastPlayMonth & "/" & lastPlayYear & " 00:00:00"
			
		End If
		
		If NOT IsNull(traktorAdded) Then
			
			traktorAddedYear = left(traktorAdded,4)
			traktorAddedMonth = mid(traktorAdded, 6, instr(6,traktorAdded,"/")-6)
			traktorAddedDay = mid(traktorAdded, instr(6,traktorAdded,"/")+1)
			songUpdate.DateAdded = traktorAddedDay & "/" & traktorAddedMonth & "/" & traktorAddedYear & " 00:00:00"
				
		End If
		
	End If
	
	'call indx.Add(trax.item)
	Progress.Value = Progress.Value + 1
	songUpdate.UpdateDB
	songUpdate.WriteTags
	
	next
	
	
	call dat.commit()
	
End Sub


Sub CopyFile(SourceFile, DestinationFile)

	Set fso = CreateObject("Scripting.FileSystemObject")

		'Check to see if the file already exists in the destination folder
		Dim wasReadOnly
		wasReadOnly = False
		If fso.FileExists(DestinationFile) Then
        'Check to see if the file is read-only
        If fso.GetFile(DestinationFile).Attributes And 1 Then 
            'The file exists and is read-only.
            'WScript.Echo "Removing the read-only attribute"
            'Remove the read-only attribute
            fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes - 1
            wasReadOnly = True
        End If

        'WScript.Echo "Deleting the file"
        fso.DeleteFile DestinationFile, True
    End If

    'Copy the file
    'WScript.Echo "Copying " & SourceFile & " to " & DestinationFile
    fso.CopyFile SourceFile, DestinationFile, True

    If wasReadOnly Then
        'Reapply the read-only attribute
        fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes + 1
    End If

    Set fso = Nothing

End Sub
