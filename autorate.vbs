' ====================
' AutoRating
' ====================
' Version 1.0.0.1 - September 2014
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

Dim objApp
Dim objLibrary
Dim colTracks
Dim wholeStarRatings
Dim rateUnratedTracksOnly
Dim ratingBias
Dim ratingMemory
Dim useHalfStarForItemsWithMoreSkipsThanPlays
Dim minRating
Dim maxRating
Dim skipCountFactor
Dim binLimitFrequencies
Dim binLimitCounts
Dim theNow
Dim theTrackCount
Dim numAnalysed
Dim numTracksToAnalyse
Dim playCount
Dim skipCount
Dim trackLength
Dim theDateAdded
Dim combinedCount
Dim combinedFrequency
Dim binLimits
Dim binLimitIndex
Dim minRatingPercent
Dim maxRatingPercent
Dim ratingScale
Dim minBin
Dim maxBin
Dim binIncrement
Dim theOldRating
Dim theRating
Dim bin
Dim frequencyMethodRating
Dim countMethodRating
Dim updated
Dim commentRating
Dim commentPlayCount
Dim commentSkipCount
Dim commentValue
Dim restoreNeeded
Dim backupComments
Dim restoreCounts

'Logfile
Dim strPath
Dim strFolder
Set objShell = CreateObject("Wscript.Shell")
strPath = Wscript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile) 
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.CreateTextFile(strFolder & "\autorate.log")
Set objApp = CreateObject("iTunes.Application")
Set objLibrary = objApp.LibraryPlaylist

'Playlist selection
Set playlists = objApp.LibrarySource.Playlists
Set playlist = playlists.ItemByName("Testplaylist")

'Set colTracks = playlist.Tracks 'use single playlist
Set colTracks = objLibrary.Tracks 'use whole library

'Globals
backupComments = true
restoreComments = true
wholeStarRatings = false
rateUnratedTracksOnly = false
ratingBias = 0.8 'Percentage if the rating should look after recent played or max played. Bigger value means more count oriented.
ratingMemory = 0.0 'Percentage of how much of the old rating should take into account
skipCountFactor = 1 'Multiplier for skips

useHalfStarForItemsWithMoreSkipsThanPlays = true
minRating = 1.0
maxRating = 5.0
binLimitFrequencies = Array(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
binLimitCounts = Array(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
updated = 0

theNow = Now

'Get backup values from track comment
Sub GetCommentValues(C)
	Dim s,r
	commentRating = 0
	commentPlayCount = 0
	commentSkipCount = 0
	if InStr(C,"Rating:") <> 0 and InStr(C,"PlayCount:") <> 0 and InStr(C,"SkipCount:") <> 0 then
		s = Split(C,",")
		For Each v in s
			r = Split(v,":")
			if StrComp(r(0),"Rating") = 0 then
				commentRating = Int(r(1))
			elseif StrComp(r(0),"PlayCount") = 0 then
				commentPlayCount = Int(r(1))
			elseif StrComp(r(0),"SkipCount") = 0 then
				commentSkipCount = Int(r(1))
			end if
		Next
	end if
End Sub

'Start script
objLog.WriteLine "AutoRate (C) 2014 Sven Wilkens | Runtime: " & theNow 
Wscript.Echo "AutoRate (C) 2014 Sven Wilkens"
'Initialise statistical analysis temp values
set sortedFrequencyList = CreateObject("System.Collections.ArrayList")
set sortedCountList = CreateObject("System.Collections.ArrayList")
theTrackCount = 0
numTracksToRate = colTracks.count

if restoreComments then
	'Restore from comments
	Wscript.Stdout.Write "Restore counts from comments if needed: ["
	WScript.Stdout.Write numTracksToRate
	WScript.Stdout.Write "/"
	Wscript.Stdout.Write chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32)
	For Each objTrack in colTracks
		restoreNeeded = false
		GetCommentValues(objTrack.Comment)
		
		if objTrack.PlayedCount < commentPlayCount then
			objTrack.PlayedCount = commentPlayCount
			restoreNeeded = true
		end if
		if objTrack.SkippedCount < commentSkipCount then
			objTrack.SkippedCount = commentSkipCount
			restoreNeeded = true
		end if
		
		theTrackCount = theTrackCount + 1
		
		if restoreNeeded then
			objLog.WriteLine "##Updated## | Title: " & objTrack.Name & " | Artist: " & objTrack.Artist & " | PlayCount: " & objTrack.PlayedCount & " | SkipCount: " & objTrack.SkippedCount & " | Rating: " & theRating
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

'Analyse tracks
updated = 0
theTrackCount = 0
numAnalysed = 0
Wscript.Stdout.Write "Create statistics: ["
WScript.Stdout.Write numTracksToRate
WScript.Stdout.Write "/"
Wscript.Stdout.Write chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32) & chr(32)
For Each objTrack in colTracks
	theTrackCount = theTrackCount + 1
	playCount = Int(objTrack.PlayedCount)
	skipCount = objTrack.SkippedCount * skipCountFactor
	trackLength = 1 '(the finish of theTrack) - (the start of theTrack)
							
	if playCount > skipCount then
		numAnalysed = numAnalysed + 1
		theDateAdded = objTrack.DateAdded
		combinedCount = Int((playCount - skipCount) * trackLength)
		if combinedCount <= 0 then
			combinedCount = 0
			combinedFrequency = 0.0
		else
			combinedFrequency = (combinedCount / (theNow - theDateAdded))
		end if
		sortedCountList.Add combinedCount
		sortedFrequencyList.Add combinedFrequency
		
	end if
	WScript.Stdout.Write chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8)
	Wscript.Stdout.Write theTrackCount
	Wscript.Stdout.Write "]"
	For i = 1 to (9 - Len(theTrackCount))
		Wscript.Stdout.Write chr(32)
	next
Next
Wscript.Stdout.WriteBlankLines(1)

if sortedFrequencyList.count() > 0 and sortedCountList.count() > 0 then
	'sort the lists so we can find the item at lower and upper percentiles and bin the values in a histogram.
	sortedFrequencyList.sort()
	sortedCountList.sort()
	'                       1           2            3          4           5  Stars
	'				  10 ,  20,   30  ,  40  , 50 ,  60 , 70  , 80  , 90    100 Percentage
	binLimits = Array(0.33, 0.34, 0.53, 0.54, 0.70, 0.71, 0.84, 0.85, 0.95, 0.96)
	'binLimits = Array(0.0, 0.01, 0.04, 0.11, 0.23, 0.4, 0.6, 0.77, 0.89, 0.96) 'Cumulative normal density for each bin
	set binLimitFrequencies = CreateObject("System.Collections.ArrayList")
	set binLimitCounts = CreateObject("System.Collections.ArrayList")
	
	For Each binLimit in binLimits
		binLimitIndex = Int(numAnalysed * binLimit)
		if binLimitIndex < 1 then
			binLimitIndex = 1
		elseif binLimitIndex > numAnalysed then
			binLimitIndex = numAnalysed
		end if
		binLimitFrequencies.Add sortedFrequencyList(binLimitIndex-1)
		binLimitCounts.Add sortedCountList(binLimitIndex-1)
	Next

	'Left analysis loop
	minRatingPercent = minRating * 20
	maxRatingPercent = maxRating * 20

	'Assign ratings

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
	ratingScale = maxRatingPercent - minRatingPercent

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
		theTrackCount = theTrackCount + 1
			
		if (not rateUnratedTracksOnly) or (objTrack.Rating = 0) then
		
			playCount = Int(objTrack.PlayedCount)
			skipCount = Int(objTrack.SkippedCount * skipCountFactor) 'weighted skips relative to plays
			
			theDateAdded = objTrack.DateAdded

			combinedCount = Int(playCount - skipCount)
			
			if combinedCount <= 0 then
				combinedCount = 0
				combinedFrequency = 0.0
			else
				combinedFrequency = (combinedCount / (theNow - theDateAdded))
			end if
			
			theOldRating = objTrack.Rating
			if playCount = 0 and skipCount = 0 then
				theRating = 0
				'Override calculated rating if the weighted skip count is greater than the play count and ignores rating memory
			elseif useHalfStarForItemsWithMoreSkipsThanPlays and (playCount < skipCount) then
				theRating = 10
			else
				'Frequency method
				bin = maxBin

				while combinedFrequency < binLimitFrequencies(bin-1) and bin > minBin
					bin = bin - binIncrement
				wend
				frequencyMethodRating = bin * 10.0
				
				'Count method
				bin = maxBin
				while combinedCount < binLimitCounts(bin-1) and bin > minBin
					bin = bin - binIncrement
				wend
				countMethodRating = bin * 10.0
				
				'Combine ratings according to user preferences
				theRating = (frequencyMethodRating * (1.0 - ratingBias)) + (countMethodRating * ratingBias)
				
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
			
			'Save to track
			'Wscript.Echo theTrackCount & " | Name: " & objTrack.Name & " | Rating: " & theRating
			WScript.Stdout.Write chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8) & chr(8)
			Wscript.Stdout.Write theTrackCount
			Wscript.Stdout.Write "]"
			For i = 1 to (9 - Len(theTrackCount))
				Wscript.Stdout.Write chr(32)
			next
			
			commentValue = "PlayCount:" & objTrack.PlayedCount & ",SkipCount:" & objTrack.SkippedCount & ",Rating:" & theRating
			
			'Set rating
			if theOldRating <> theRating then
				objTrack.Rating = theRating
				'rating set successfully	
			end if
			'Backup Values to comment
			if StrComp(objTrack.Comment,commentValue) <> 0 then
				if backupComments then
					objTrack.Comment = commentValue
				end if
				objLog.WriteLine "#Updated | ID: " & objTrack.Trackid & " | Played: " & objTrack.PlayedCount & " | Skipped: " & objTrack.SkippedCount & " | ORating: " & theOldRating & " | NRating: " & theRating
				updated = updated + 1
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
	objLog.WriteLine "#"
	objLog.WriteLine
	WScript.Echo "Done!"
else
	WScript.Echo "Script aborded because no tracks are played or all songs have a higher/equal SkipCount than PlayCount."
	objLog.WriteLine "Script aborded because no tracks are played or all songs have a higher/equal SkipCount than PlayCount."
end if
