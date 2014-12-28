' ====================
' AutoRating
' ====================
' Version 2.0.0.2 - September 2014
' Copyright © Sven Wilkens 2014
' https://plus.google.com/u/0/+SvenWilkens

' =======
' Licence
' =======
' This program is free software: you can redistribute it and/or modify it under the terms
' of the GNU General Public License as published by the Free Software Foundation, either
' version 3 of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License for more details.

' Please visit http://www.gnu.org/licenses/gpl-3.0-standalone.html to view the GNU GPLv3 licence.

' ===========
' Description
' ===========
' Automatically sets the rating for all tracks in your iTunes library according to how you listen to them.
' It also backup the Play and Skip Counts in the Comment field of the mp3 and disable iTunes album rating

' Related scripts: https://code.google.com/p/autorate/

' =========
' ChangeLog 
' =========
' Version 1.0.0.1 - Initial version
' Version 1.0.0.2 - Added darrating
' Version 2.0.0.1 - new algorithm
' Version 2.0.0.2 - Bugfix

'#########Variables#########
'General
Dim objApp,objLibrary,colTracks
'Counter
Dim theTrackCount,numAnalysed,updated,up,down,equal
'Track
Dim playCount,skipCount,trackLength,theDateAdded,theOldRating,theRating
'Calculation
Dim binLimits,binLimitIndex,minBin,maxBin,binIncrement,bin,minRatingPercent,maxRatingPercent,durationTmp
'Comment
Dim commentDivider,commentRating,commentPlayCount,commentSkipCount,commentValue,commentDateAdded,commentPlayedDate,commentSkippedDate
'Configuration
Dim minRating,maxRating,rateUnratedTracksOnly,ratingMemory,useHalfStarForItemsWithMoreSkipsThanPlays,wholeStarRatings,restoreNeeded,updateNeeded,backupComments,createPlaylist,playlistName,simulate

'#########Logfile#########
Dim strPath,strFolder
Set objShell = CreateObject("Wscript.Shell")
strPath = Wscript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile) 
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.CreateTextFile(strFolder & "\autorate.log")

'#########Init#########
Set objApp = CreateObject("iTunes.Application")
Set objLibrary = objApp.LibraryPlaylist

'#########Playlist selection#########
Set playlists = objApp.LibrarySource.Playlists
'Set sourcePlaylist = playlists.ItemByName("Testplaylist")
'Set sourcePlaylist = playlists.ItemByName("Musik") 'or Music means whole music library with music videos
Set sourcePlaylist = playlists.ItemByName("MusicOnly") 'Smart Playlist with music only

set colTracks = sourcePlaylist.Tracks 'use single playlist
'Set colTracks = objLibrary.Tracks 'use whole library

'###############################
'#########Configuration#########
'###############################
restoreComments = true 'default:true
backupComments = true 'default:true
simulate = false 'default:false
wholeStarRatings = false 'default:false
rateUnratedTracksOnly = false 'default:false
useHalfStarForItemsWithMoreSkipsThanPlays = false 'default:false
'Playlist
createPlaylist = true 'default:true
playlistName = "LastAutoRated"
'Rating
ratingMemory = 0.0 'Percentage of how much of the old rating should take into account
minRating = 1.0
maxRating = 5.0
'                       1           2            3          4           5  Stars
'				  10 ,  20,   30  ,  40  , 50 ,  60 , 70  , 80  , 90    100 Percentage
binLimits = Array(0.33, 0.34, 0.53, 0.54, 0.70, 0.71, 0.84, 0.85, 0.95, 0.96)
'binLimits = Array(0.0, 0.01, 0.04, 0.11, 0.23, 0.4, 0.6, 0.77, 0.89, 0.96) 'Cumulative normal density for each bin
'###############################
'###############################
'###############################

'Time
Dim theNow,atb,offsetMin,theNowUTC
theNow = Now
atb = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\" &_ 
        "Control\TimeZoneInformation\ActiveTimeBias" 
offsetMin = objShell.RegRead(atb) 
theNowUTC = dateadd("n", offsetMin, theNow)

'#########Functions#########

'Get backup values from track comment
Sub GetCommentValues(C)
	Dim s,r
	commentRating = 0
	commentPlayCount = 0
	commentSkipCount = 0
	commentDateAdded = 0
	commentPlayedDate = 0
	commentSkippedDate = 0
	if InStr(C,"Rating#") <> 0 and InStr(C,"PlayCount#") <> 0 and InStr(C,"SkipCount#") <> 0 and InStr(C,"AddedDate#") <> 0 and InStr(C,"PlayedDate#") <> 0 and InStr(C,"SkippedDate#") <> 0 then
		s = Split(C,",")
		For Each v in s
			r = Split(v,"#")
			if StrComp(r(0),"Rating") = 0 then
				commentRating = Int(r(1))
			elseif StrComp(r(0),"PlayCount") = 0 then
				commentPlayCount = Int(r(1))
			elseif StrComp(r(0),"SkipCount") = 0 then
				commentSkipCount = Int(r(1))
			elseif StrComp(r(0),"AddedDate") = 0 then
				commentDateAdded = r(1)
			elseif StrComp(r(0),"PlayedDate") = 0 then
				commentPlayedDate = r(1)
			elseif StrComp(r(0),"SkippedDate") = 0 then
				commentSkippedDate = r(1)
			end if
		Next
	end if
End Sub

'Calculate Score
Function getScore(objTrack)
	Dim daysSinceImported,daysSinceLastPlayed,daysSinceLastSkipped,oTrackLength,score,playedTime,lastPlayed,lastSkipped
	playCount = Int(objTrack.PlayedCount)
	skipCount = Int(objTrack.SkippedCount)
	trackLength = objTrack.Finish - objTrack.Start '1 '(the finish of theTrack) - (the start of theTrack)
	lastPlayed = objTrack.PlayedDate
	if lastPlayed = 0 then 
		lastPlayed = theNow
	end if
	lastSkipped = objTrack.SkippedDate
	if lastSkipped = 0 then 
		lastSkipped = theNow
	end if
	daysSinceLastPlayed = DateDiff("d",lastPlayed,theNow)
	daysSinceLastSkipped = DateDiff("d",lastSkipped,theNow)
	daysSinceImported = DateDiff("d",objTrack.DateAdded,theNow) + 1 'tracks added today???
	
	'Duration optimizer: boost short tracks
	if trackLength > 3599 then 
		durationTmp = Round((6000*trackLength)/3600)
	else
		durationTmp = 6000
	end if
	oTrackLength = Round((trackLength+360)/3) + Round((trackLength*trackLength)/durationTmp)
	playedTime = playCount * oTrackLength
	
	'Public Const Big_Berny_Formula_1 = "(10000000 * (7+OptPlayed-(Skip*0.98^(SongLength/60))^1.7)^3 / (10+DaysSinceFirstPlayed)^0.5) / (1+DaysSinceLastPlayed/365)"
	'Public Const Big_Berny_Formula_2 = "(10000000 * (7+OptPlayed-(Skip*0.98^(SongLength/60))^1.7)^3 / (10+DaysSinceFirstPlayed)^0.3) / (1+DaysSinceLastPlayed/730)"
	'Public Const Big_Berny_Formula_4 = "(10000000 * (7+Played-(Skip*0.98^(SongLength/60))^1.7)^3 / (10+DaysSinceFirstPlayed)^0.5) / (1+DaysSinceLastPlayed/365)"
	'Public Const Big_Berny_Formula_5 = "7+OptPlayed-(Skip*0.98^(SongLength/60))"
	'Public Const BerniPi_Formula_1 = "(500000000000+10000000000*(Played*0.999^((10+DaysSinceLastPlayed)/(Played/3+1))-Skip^1.7))/((10+DaysSinceFirstPlayed)/(Played^2+1))"
	'score = Int((10000000 * (7 + playedTime + (daysSinceLastSkipped / 365)^1.2 -(skipCount*0.98^(otrackLength/60))^1.7)^3 / (10 + daysSinceImported)^0.5) / ((daysSinceLastPlayed / 365) + 1))
	
	score = Int((10000000 * (7 + playedTime + (daysSinceLastSkipped*0.5^(otrackLength/60))^1.2 -(skipCount*0.98^(otrackLength/60))^3)^3 / (10 + daysSinceImported)^0.5) / ((daysSinceLastPlayed / 365) + 1))
	
	If score < 0 Then
        score = 0.0
    End If
	getScore = score
End Function

'#########Start script#########
objLog.WriteLine "AutoRate (C) 2014 Sven Wilkens | Runtime: " & theNow 
Wscript.Echo "AutoRate (C) 2014 Sven Wilkens"
'Init Playlist
if createPlaylist then
	set playlist = playlists.ItemByName(playlistName)
	if NOT playlist is Nothing then
		playlist.Delete
	end if
	set playlist = objApp.CreatePlaylist(playlistName)
end if

'Initialise statistical analysis temp values
set sortedFrequencyList = CreateObject("System.Collections.ArrayList")
set sortedCountList = CreateObject("System.Collections.ArrayList")
set sortedScoreList = CreateObject("System.Collections.ArrayList")
theTrackCount = 0
numTracksToRate = colTracks.count
up = 0
down = 0
equal = 0

'#########Restore from comments#########
updated = 0
if restoreComments then
	objLog.WriteLine "----------Restore from comments----------"
	Wscript.Echo "----------Restore from comments----------"
	Wscript.Stdout.Write "Restore counts from comments if needed: ["
	WScript.Stdout.Write numTracksToRate
	WScript.Stdout.Write "/"
	Wscript.Stdout.Write chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32)
	For Each objTrack in colTracks
		restoreNeeded = false
		GetCommentValues(objTrack.Comment)
		
		if objTrack.PlayedCount < commentPlayCount then
			objTrack.PlayedCount = commentPlayCount
			objTrack.PlayedDate = commentPlayedDate
			restoreNeeded = true
		end if
		if objTrack.SkippedCount < commentSkipCount then
			objTrack.SkippedCount = commentSkipCount
			objTrack.SkippedDate = commentSkippedDate
			restoreNeeded = true
		end if

		'Date Added is read only
		'if DateDiff("d",objTrack.DateAdded,commentDateAdded) = 0 then
		'	objTrack.DateAdded = commentDateAdded
		'	restoreNeeded = true
		'end if
		
		theTrackCount = theTrackCount + 1
				
		'objEx.WriteLine objTrack.trackDatabaseID & "," & objTrack.PlayedCount & "," & objTrack.SkippedCount & "," & objTrack.Finish - objTrack.Start & "," & objTrack.DateAdded & "," & objTrack.PlayedDate & "," & objTrack.SkippedDate
		if restoreNeeded then
			objLog.WriteLine Mid("------------------------------" & updated & "------------------------------",1,61)
			objLog.WriteLine chr(9) & "ID: " & chr(9) & chr(9) & objTrack.trackDatabaseID
			objLog.WriteLine chr(9) & "Title: " & chr(9) & chr(9) & objTrack.Name
			objLog.WriteLine chr(9) & "Artist: " & chr(9) & objTrack.Artist
			objLog.WriteLine chr(9) & "Length: " & chr(9) & objTrack.Time
			objLog.WriteLine chr(9) & "Played: " & chr(9) & objTrack.PlayedCount
			objLog.WriteLine chr(9) & "Last Played: " & chr(9) & objTrack.PlayedDate
			objLog.WriteLine chr(9) & "Skipped: " & chr(9) & objTrack.SkippedCount
			objLog.WriteLine chr(9) & "Last Skipped: " & chr(9) & objTrack.SkippedDate
			objLog.WriteLine chr(9) & "Date added: " & chr(9) & objTrack.DateAdded
			updated = updated + 1
		end if
		
		WScript.Stdout.Write chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8)
		Wscript.Stdout.Write theTrackCount
		Wscript.Stdout.Write "]"
		For i = 1 to (9 - Len(theTrackCount))
			Wscript.Stdout.Write chr(32)
		next
	Next
	Wscript.Stdout.WriteBlankLines(1)
	WScript.Echo updated & " Files restored from comment."

	objLog.WriteLine
	objLog.WriteLine "#"
	objLog.WriteLine "# " & updated & " Files restored from comment."
	objLog.WriteLine "#"
	objLog.WriteLine
end if

'#########Analyse tracks#########
Wscript.Echo "----------Analyse tracks----------"
updated = 0
theTrackCount = 0
numAnalysed = 0
Wscript.Stdout.Write "Create statistics: ["
WScript.Stdout.Write numTracksToRate
WScript.Stdout.Write "/"
Wscript.Stdout.Write chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32)

For Each objTrack in colTracks
	theTrackCount = theTrackCount + 1
	score = getScore(objTrack)
	sortedScoreList.Add score
	numAnalysed = numAnalysed + 1
	
	WScript.Stdout.Write chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8)
	Wscript.Stdout.Write theTrackCount
	Wscript.Stdout.Write "]"
	For i = 1 to (9 - Len(theTrackCount))
		Wscript.Stdout.Write chr(32)
	next
Next

Wscript.Stdout.WriteBlankLines(1)

if sortedScoreList.count() > 0 then
	'sort the lists so we can find the item at lower and upper percentiles and bin the values in a histogram.
	sortedScoreList.sort()
	set binLimitScore = CreateObject("System.Collections.ArrayList")

	For Each binLimit in binLimits
		binLimitIndex = Int(numAnalysed * binLimit)
		if binLimitIndex < 1 then
			binLimitIndex = 1
		elseif binLimitIndex > numAnalysed then
			binLimitIndex = numAnalysed
		end if
		binLimitScore.Add sortedScoreList(binLimitIndex-1)
	Next
	
	objLog.WriteLine "Rating | Score Border"
	objLog.WriteLine "-------|-------------"
	Wscript.Echo "Rating | Score Border"
	Wscript.Echo "-------|-------------"
	Dim ratingBorder
	ratingBorder = 0.0
	For Each scoreLimit in binLimitScore
		objLog.WriteLine "   " & FormatNumber(ratingBorder,1) & " | " & scoreLimit
		Wscript.Echo "   " & FormatNumber(ratingBorder,1) & " | < " & scoreLimit
		ratingBorder = ratingBorder + 0.5
	Next

	'Left analysis loop
	minRatingPercent = minRating * 20
	maxRatingPercent = maxRating * 20

	'#########Assign ratings#########
	objLog.WriteLine "----------Assign Rating----------"
	Wscript.Echo "----------Assign Rating----------"
	
	'Correct minimum rating value if user selects whole-star ratings or to reserve 1/2 star for disliked songs
	'0 star ratings are always reserved for songs with no skips and no plays
	if (wholeStarRatings or useHalfStarForItemsWithMoreSkipsThanPlays) and (minRatingPercent < 20) then
		minRatingPercent = 20 'ie 1 star
	elseif minRatingPercent < 10 then
		minRatingPercent = 10 'ie 1/2 star
	end if

	if wholeStarRatings then
		minRatingPercent = Int(minRatingPercent / 20) * 20
		maxRatingPercent = Int(maxRatingPercent / 20) * 20
	end if

	theTrackCount = 0
	minBin = Int(minRatingPercent / 10)
	maxBin = Int(maxRatingPercent / 10)

	if wholeStarRatings then
		binIncrement = 2
	else
		binIncrement = 1
	end if
	
	Wscript.Stdout.Write "Assign Rating: ["
	WScript.Stdout.Write numTracksToRate
	WScript.Stdout.Write "/"
	Wscript.Stdout.Write chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32)			
	For Each objTrack in colTracks
		updateNeeded = false
		theTrackCount = theTrackCount + 1
		if (not rateUnratedTracksOnly) or (objTrack.Rating = 0) then
			playCount = Int(objTrack.PlayedCount)
			skipCount = Int(objTrack.SkippedCount)
			theOldRating = objTrack.Rating
			score = getScore(objTrack)
			if playCount = 0 and skipCount = 0 then
				theRating = 0
				'Override calculated rating if the weighted skip count is greater than the play count and ignores rating memory
			elseif useHalfStarForItemsWithMoreSkipsThanPlays and (playCount < skipCount) then
				theRating = 10
			else
				'Score method
				bin = maxBin
				while score < binLimitScore(bin-1) and bin >= minBin
					bin = bin - binIncrement
				wend
				theRating = bin * 10.0
				'Factor in previous rating memory
				if ratingMemory > 0.0 then
					theRating = ((theOldRating) * ratingMemory) + (theRating * (1.0 - ratingMemory))
				end if
			end if

			'Round to whole stars if requested to
			if wholeStarRatings then
				theRating = Round(theRating / 20) * 20
			else
				theRating = Round(theRating / 10) * 10
			end if
			if theRating = 0 then 
				theRating = 1
			end if

			'Save to track
			'Wscript.Echo theTrackCount & " | Name: " & objTrack.Name & " | Rating: " & theRating
			WScript.Stdout.Write chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8)
			Wscript.Stdout.Write theTrackCount
			Wscript.Stdout.Write "]"
			For i = 1 to (9 - Len(theTrackCount))
				Wscript.Stdout.Write chr(32)
			next
			
			commentValue = "PlayCount#" & objTrack.PlayedCount & ",SkipCount#" & objTrack.SkippedCount & ",Rating#" & theRating & ",AddedDate#" & objTrack.DateAdded & ",PlayedDate#" & objTrack.PlayedDate & ",SkippedDate#" & objTrack.SkippedDate
			
			'Set rating
			if (theOldRating <> theRating) And NOT simulate then
				objTrack.Rating = theRating
				updateNeeded = true
				if createPlaylist then
					playlist.AddTrack(objTrack)
				end if
				'rating set successfully	
			end if
			
			'Backup Values to comment
			if StrComp(objTrack.Comment,commentValue) <> 0 then
				updateNeeded = true
				if backupComments then
					objTrack.Comment = commentValue
				end if

			end if
			
			'Log if changed
			if updateNeeded then 
				updated = updated + 1
				On Error Resume Next
				objLog.WriteLine Mid("------------------------------" & updated & "------------------------------",1,61)
				objLog.WriteLine chr(9) & "ID: " & chr(9) & chr(9) & objTrack.trackDatabaseID
				objLog.WriteLine chr(9) & "Title: " & chr(9) & chr(9) & objTrack.Name
				objLog.WriteLine chr(9) & "Artist: " & chr(9) & objTrack.Artist
				objLog.WriteLine chr(9) & "Length: " & chr(9) & objTrack.Time
				objLog.WriteLine chr(9) & "Played: " & chr(9) & objTrack.PlayedCount
				objLog.WriteLine chr(9) & "Last Played: " & chr(9) & objTrack.PlayedDate
				objLog.WriteLine chr(9) & "Skipped: " & chr(9) & objTrack.SkippedCount
				objLog.WriteLine chr(9) & "Last Skipped: " & chr(9) & objTrack.SkippedDate
				objLog.WriteLine chr(9) & "Date added: " & chr(9) & objTrack.DateAdded
				objLog.WriteLine chr(9) & "Old Rating: " & chr(9) & theOldRating
				objLog.WriteLine chr(9) & "New Rating: " & chr(9) & theRating
				objLog.WriteLine
				if theRating > theOldRating then 
					objLog.WriteLine chr(9) & chr(9) & chr(94) & " Rating goes up!"
					up = up + 1
				elseif theRating < theOldRating then
					objLog.WriteLine chr(9) & chr(9) & chr(118) & " Rating goes down!"
					down = down + 1
				else
					objLog.WriteLine chr(9) & chr(9) & chr(61) & " Rating keeps equal!"
					equal = equal + 1
				end if
				objLog.WriteLine
			end if
			
			'Disable iTunes auto rating
			if objTrack.AlbumRating <> 1 then
				objTrack.AlbumRating = 1
			end if
		end if
	Next
	Wscript.Stdout.WriteBlankLines(1)				
	WScript.Echo updated & " File ratings updated."
	
	objLog.WriteLine
	objLog.WriteLine "#"
	objLog.WriteLine "# " & updated & " File ratings updated."
	objLog.WriteLine "# " & up & " File ratings goes up."
	objLog.WriteLine "# " & down & " File ratings goes down."
	objLog.WriteLine "# " & equal & " File ratings keeps equal."
	objLog.WriteLine "#"
	objLog.WriteLine
	WScript.Echo "Done!"
	objShell.run "notepad.exe " & strFolder & "\autorate.log"
else
	WScript.Echo "Script aborded because no tracks are available."
	objLog.WriteLine "Script aborded because no tracks are available."
end if
