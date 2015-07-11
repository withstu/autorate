__author__ = 'Sven'

import datetime
import math
import win32com.client
import logging
import pylast
import operator
import timeit
import os

#Switch to script directory
abspath = os.path.abspath(__file__)
dname = os.path.dirname(abspath)
os.chdir(dname)

start = timeit.default_timer()

##########Variables#########
#Logfile
logging.basicConfig(filename='autorate-py.log', filemode='w', level=logging.INFO)

################################
##########Configuration#########
################################
API_KEY = ""
API_SECRET = ""

restoreComments = True #default:true
backupComments = True #default:true
simulate = False #default:false
wholeStarRatings = False #default:false
rateUnratedTracksOnly = False #default:false
useHalfStarForItemsWithMoreSkipsThanPlays = False #default:false
durationEffect = "megaweak" #default:default;options: full,strong,moderate,default,weak,veryweak,extremelyweak,superweak,megaweak,supermegaweak,ignore
#Playlist
createPlaylist = True #default:true
playlistName = "LastAutoRated"
#Top Track Playlist
topplaylistName = "Top1000"
topcount = 1000
#Updated Playlist
playCounterPlaylistName = "PlayCounterUpdated"
skipCounterPlaylistName = "SkipCounterUpdated"
#Rating
ratingMemory = 0.0 #Percentage of how much of the old rating should take into account
minRating = 1.0
maxRating = 5.0
#                       1           2            3          4           5  Stars
#				  10 ,  20,   30  ,  40  , 50 ,  60 , 70  , 80  , 90    100 Percentage
binLimits = [0.33, 0.34, 0.53, 0.54, 0.70, 0.71, 0.84, 0.85, 0.95, 0.96]
#binLimits = Array(0.0, 0.01, 0.04, 0.11, 0.23, 0.4, 0.6, 0.77, 0.89, 0.96) 'Cumulative normal density for each bin
################################
################################
################################
#Time
theNow = datetime.datetime.now()
nullDate = datetime.datetime(1970, 1, 1, 0, 0)

##########Functions#########

def getItunes():
    return win32com.client.Dispatch("iTunes.Application")

def getItunesPlaylist(playlist):
    # First, we create an instance of the iTunes Application
    itunes = win32com.client.Dispatch("iTunes.Application")

    playlists = itunes.LibrarySource.Playlists
    sourcePlaylist = playlists.ItemByName(playlist)

    return sourcePlaylist.Tracks

def getItunesPlaylists():
    # First, we create an instance of the iTunes Application
    itunes = win32com.client.Dispatch("iTunes.Application")

    playlists = itunes.LibrarySource.Playlists

    return playlists

def getLastFmScore(objTrack):
    lfm = pylast.LastFMNetwork(api_key=API_KEY, api_secret=API_SECRET)
    track = lfm.get_track(objTrack.Artist, objTrack.Name)
    playcount = track.get_playcount()
    listenercount = track.get_listener_count()
    userloved = track.get_userloved()
    score = 1
    return score

#Get backup values from track comment
def GetCommentValues(track):
    comment = track['Comment']
    commentValues = {"commentRating" : 0,
                     "commentPlayCount" : 0,
                     "commentSkipCount" : 0,
                     "commentDateAdded" : 0,
                     "commentPlayedDate" : 0,
                     "commentSkippedDate" : 0
                     }

    if (
                                    comment.find("Rating#") != -1
                            and comment.find("PlayCount#") != -1
                        and comment.find("SkipCount#") != -1
                    and comment.find("AddedDate#") != -1
                and comment.find("PlayedDate#") != -1
            and comment.find("SkippedDate#") != -1
    ):
        s = comment.split(",")
        for v in s:
            r = v.split("#")
            if r[0] == "Rating":
                commentValues['commentRating'] = int(r[1])
            elif r[0] == "PlayCount":
                commentValues['commentPlayCount'] = int(r[1])
            elif r[0] == "SkipCount":
                commentValues['commentSkipCount'] = int(r[1])
            elif r[0] == "AddedDate":
                commentValues['commentDateAdded'] = r[1]
            elif r[0] == "PlayedDate":
                commentValues['commentPlayedDate'] = r[1]
            elif r[0] == "SkippedDate":
                commentValues['commentSkippedDate'] = r[1]

    return commentValues


# Full Duration Effect
def getDurationFull(trackLength):
    return trackLength


# Strong Duration Effect
def getDurationStrong(trackLength):
    if trackLength > 3599:
        durationTmp = round((6000 * trackLength) / 3600)
    else:
        durationTmp = 6000
    return round((trackLength + 360) / 3) + round((trackLength * trackLength) / durationTmp)


# Moderate Duration Effect
def getDurationModerate(trackLength):
    if trackLength > 3599:
        durationTmp = round((10000 * trackLength) / 3600)
    else:
        durationTmp = 10000
    return round((trackLength + 360) / 3) + round((trackLength * trackLength) / durationTmp)


# Default Duration Effect
def getDurationDefault(trackLength):
    if trackLength > 3599:
        durationTmp = round((9000 * trackLength) / 3600)
    else:
        durationTmp = 9000
    return round((trackLength + 540) / 4) + round((trackLength * trackLength) / durationTmp)


# Weak Duration Effect
def getDurationWeak(trackLength):
    if trackLength > 3599:
        durationTmp = round((10000 * trackLength) / 3600)
    else:
        durationTmp = 10000
    return round((trackLength + 720) / 5) + round((trackLength * trackLength) / durationTmp)


# Very Weak Duration Effect
def getDurationVeryWeak(trackLength):
    if trackLength > 3599:
        durationTmp = round((20000 * trackLength) / 3600)
    else:
        durationTmp = 20000
    return round((trackLength + 720) / 5) + round((trackLength * trackLength) / durationTmp)


# Extremely Weak Duration Effect
def getDurationExtremelyWeak(trackLength):
    if trackLength > 3599:
        durationTmp = round((20000 * trackLength) / 3600)
    else:
        durationTmp = 20000
    return round((trackLength + 900) / 6) + round((trackLength * trackLength) / durationTmp)


# Super Weak Duration Effect
def getDurationSuperWeak(trackLength):
    if trackLength > 3599:
        durationTmp = round((20000 * trackLength) / 3600)
    else:
        durationTmp = 20000
    return round((trackLength + 1080) / 7) + round((trackLength * trackLength) / durationTmp)


# Mega Weak Duration Effect
def getDurationMegaWeak(trackLength):
    if trackLength > 3599:
        durationTmp = round((20000 * trackLength) / 3600)
    else:
        durationTmp = 20000
    return round((trackLength + 1260) / 8) + round((trackLength * trackLength) / durationTmp)


# Super Mega Weak Duration Effect
def getDurationSuperMegaWeak(trackLength):
    if trackLength > 3599:
        durationTmp = round((20000 * trackLength) / 3600)
    else:
        durationTmp = 20000
    return round((trackLength + 3564) / 18) + round((trackLength * trackLength) / durationTmp)


# Full Duration Effect
def getDurationIgnore(trackLength):
    return 180

def getScore(track):
    playCount = track['PlayedCount']
    skipCount = track['SkippedCount']
    trackLength = track['Finish'] - track['Start']
    lastPlayed = track['PlayedDate']
    lastSkipped = track['SkippedDate']
    dateAdded = track['DateAdded']

    daysSinceLastPlayed = (theNow - lastPlayed).days
    if daysSinceLastPlayed < 0:
        daysSinceLastPlayed = 0
    daysSinceLastSkipped = (theNow - lastSkipped).days
    if daysSinceLastSkipped < 0:
        daysSinceLastSkipped = 0
    daysSinceImported = (theNow - dateAdded).days
    if daysSinceImported < 0:
        daysSinceImported = 0

    if durationEffect == 'full':
        oTrackLength = getDurationFull(trackLength)
    elif durationEffect == 'strong':
        oTrackLength = getDurationStrong(trackLength)
    elif durationEffect == "moderate" :
        oTrackLength = getDurationModerate(trackLength)
    elif durationEffect == "weak" :
        oTrackLength = getDurationWeak(trackLength)
    elif durationEffect == "veryweak" :
        oTrackLength = getDurationVeryWeak(trackLength)
    elif durationEffect == "extremelyweak" :
        oTrackLength = getDurationExtremelyWeak(trackLength)
    elif durationEffect == "superweak" :
        oTrackLength = getDurationSuperWeak(trackLength)
    elif durationEffect == "megaweak" :
        oTrackLength = getDurationMegaWeak(trackLength)
    elif durationEffect == "supermegaweak" :
        oTrackLength = getDurationSuperMegaWeak(trackLength)
    elif durationEffect == "ignore" :
        oTrackLength = getDurationIgnore(trackLength)
    else:
        oTrackLength = getDurationDefault(trackLength)
    playedTime = math.sqrt(playCount) * oTrackLength

    score = (((playedTime - (math.sqrt(skipCount)*oTrackLength*0.9**(oTrackLength/60)*0.6**(daysSinceLastSkipped / 365))) / (30 + daysSinceImported)**0.2)*100) / ((daysSinceLastPlayed**1.2 / 730) + 1)

    return score


def main():
    start = timeit.default_timer()
    logging.info("Autorate (C) " + str(theNow.year) + " Sven Wilkens | Runtime: " + str(theNow))
    #Init Itunes
    objApp = getItunes()
    playlists = getItunesPlaylists()
    #Init Playlist
    if createPlaylist:
        playlistfolder = playlists.ItemByName("AutoRate")
        if playlistfolder is None:
            playlistfolder = objApp.CreateFolder("AutoRate")

        playlist = playlistfolder.Source.Playlists.ItemByName(playlistName)
        if playlist is not None:
            playlist.Delete()
        playlist = playlistfolder.CreatePlaylist(playlistName)

        #topplaylist
        topplaylist = playlistfolder.Source.Playlists.ItemByName(topplaylistName)
        if topplaylist is not None:
            topplaylist.Delete()
        topplaylist = playlistfolder.CreatePlaylist(topplaylistName)

        #playCounterPlaylist
        playCounterPlaylist = playlistfolder.Source.Playlists.ItemByName(playCounterPlaylistName)
        if playCounterPlaylist is not None:
            playCounterPlaylist.Delete()
        playCounterPlaylist = playlistfolder.CreatePlaylist(playCounterPlaylistName)

        #skipCounterPlaylist
        skipCounterPlaylist = playlistfolder.Source.Playlists.ItemByName(skipCounterPlaylistName)
        if topplaylist is not None:
            skipCounterPlaylist.Delete()
        skipCounterPlaylist = playlistfolder.CreatePlaylist(skipCounterPlaylistName)

    #Init temp values
    tracks = getItunesPlaylist("MusicOnly")
    #tracks = getItunesPlaylist("Testplaylist")
    sortedFrequencyList, sortedCountList, sortedScoreList, scoreList = [], [], [], []
    up = 0
    down = 0
    equal = 0

    logging.info("----------Analyse tracks----------")
    updated = 0

    for objTrack in tracks:
        track = {'objTrack': objTrack,
                 'trackDatabaseID': objTrack.trackDatabaseID,
                 'Name': objTrack.Name,
                 'Artist': objTrack.Artist,
                 'Time': objTrack.Time,
                 'PlayedCount': int(objTrack.PlayedCount),
                 'PlayedDate': objTrack.PlayedDate.replace(tzinfo=None),
                 'SkippedCount': int(objTrack.SkippedCount),
                 'SkippedDate': objTrack.SkippedDate.replace(tzinfo=None),
                 'DateAdded': objTrack.DateAdded.replace(tzinfo=None),
                 'Start': objTrack.Start,
                 'Finish': objTrack.Finish,
                 'Rating': objTrack.Rating,
                 'Comment': objTrack.Comment,
                 'score': 0,
                 }

        #########Restore from comments#########
        restored = 0
        if restoreComments:
            restoreNeeded = False
            commentValues = GetCommentValues(track)

            if track['PlayedCount'] < commentValues['commentPlayCount']:
                objTrack.PlayedCount = commentValues['commentPlayCount']
                track['PlayedCount'] = commentValues['commentPlayCount']
                objTrack.PlayedDate = commentValues['commentPlayedDate']
                track['PlayedDate'] = commentValues['commentPlayedDate']
                restoreNeeded = True
            if track['SkippedCount'] < commentValues['commentSkipCount']:
                objTrack.SkippedCount = commentValues['commentSkipCount']
                track['SkippedCount'] = commentValues['commentSkipCount']
                objTrack.SkippedDate = commentValues['commentSkippedDate']
                track['SkippedDate'] = commentValues['commentSkippedDate']
                restoreNeeded = True

            if createPlaylist and track['PlayedCount'] > commentValues['commentPlayCount']:
                playCounterPlaylist.AddTrack(objTrack)

            if createPlaylist and track['SkippedCount'] > commentValues['commentSkipCount']:
                skipCounterPlaylist.AddTrack(objTrack)

            #Date Added is read only
            #if DateDiff("d",objTrack.DateAdded,commentValues['commentDateAdded']') = 0 then
            #	objTrack.DateAdded = commentDateAdded
            #	restoreNeeded = true
            #end if

            #objEx.WriteLine objTrack.trackDatabaseID & "," & objTrack.PlayedCount & "," & objTrack.SkippedCount & "," & objTrack.Finish - objTrack.Start & "," & objTrack.DateAdded & "," & objTrack.PlayedDate & "," & objTrack.SkippedDate
            if restoreNeeded:
                logging.debug("------------------------------" + str(restored) + "------------------------------")
                logging.debug(chr(9) + "ID: " + chr(9) + chr(9) + str(track['trackDatabaseID']))
                logging.debug(chr(9) + "Title: " + chr(9) + chr(9) + track['Name'])
                logging.debug(chr(9) + "Artist: " + chr(9) + track['Artist'])
                logging.debug(chr(9) + "Length: " + chr(9) + str(track['Time']))
                logging.debug(chr(9) + "Played: " + chr(9) + str(track['PlayedCount']))
                logging.debug(chr(9) + "Last Played: " + chr(9) + str(track['PlayedDate']))
                logging.debug(chr(9) + "Skipped: " + chr(9) + str(track['SkippedCount']))
                logging.debug(chr(9) + "Last Skipped: " + chr(9) + str(track['SkippedDate']))
                logging.debug(chr(9) + "Date added: " + chr(9) + str(track['DateAdded']))
                restored += 1

        score = getScore(track)
        if isinstance(score, complex):
            msg = "Score is Complex - Title: " + track['Name'] + ", Artist: " + track['Artist']
            logging.error(msg)
        else:
            track['score'] = score
            scoreList.append(track)
    if restoreComments:
        logging.info(str(restored) + " Files restored from comment.")

    if len(scoreList) > 0:
        numTracksToRate = len(scoreList)
        sortedScoreList = sorted(scoreList, key=operator.itemgetter('score'), reverse=True)
        binLimitScore = []
        for binLimit in binLimits:
            binLimitIndex = int(numTracksToRate * (1 - binLimit))
            if binLimitIndex < 1:
                binLimitIndex = 1
            elif binLimitIndex > numTracksToRate:
                binLimitIndex = numTracksToRate
            binLimitScore.append(sortedScoreList[binLimitIndex-1]['score'])
        logging.info("Rating | Score Border")
        logging.info("-------|-------------")
        ratingBorder = 0.0
        for scoreLimit in binLimitScore:
            logging.info("   " + str(ratingBorder) + " | " + str(scoreLimit))
            ratingBorder = ratingBorder + 0.5

        #Left analysis loop
        minRatingPercent = minRating * 20
        maxRatingPercent = maxRating * 20

        logging.info("----------Assign Rating----------")

        if (wholeStarRatings or useHalfStarForItemsWithMoreSkipsThanPlays) and (minRatingPercent < 20):
            minRatingPercent = 20 #ie 1 star
        elif minRatingPercent < 10:
            minRatingPercent = 10 #ie 1/2 star

        if wholeStarRatings:
            minRatingPercent = int(minRatingPercent / 20) * 20
            maxRatingPercent = int(maxRatingPercent / 20) * 20

        minBin = int(minRatingPercent / 10)
        maxBin = int(maxRatingPercent / 10)

        if wholeStarRatings:
            binIncrement = 2
        else:
            binIncrement = 1

        topplaylistcount = 0
        for track in sortedScoreList:
            objTrack = track['objTrack']
            updateNeeded = False
            if (not rateUnratedTracksOnly) or (track['Rating'] == 0):
                playCount = track['PlayedCount']
                skipCount = track['SkippedCount']
                theOldRating = track['Rating']
                score = track['score']
                if (playCount == 0 and skipCount == 0):
                    theRating = 0
                elif useHalfStarForItemsWithMoreSkipsThanPlays and (playCount < skipCount):
                    # Override calculated rating if the weighted skip count is greater than the play count and ignores rating memory
                    theRating = 10
                else:
                    # Score method
                    bin = maxBin
                    while score < binLimitScore[bin-1] and bin >= minBin:
                        bin = bin - binIncrement
                    theRating = bin * 10.0
                    #Factor in previous rating memory
                    if ratingMemory > 0.0:
                        theRating = (theOldRating * ratingMemory) + (theRating * (1.0 - ratingMemory))
                    #lastfm = getLastFmScore(objTrack)

                #Round to whole stars if requested to
                if wholeStarRatings:
                    theRating = round(theRating / 20) * 20
                else:
                    theRating = round(theRating / 10) * 10

                #Disable Album Rating
                if theRating == 0:
                    theRating = 1

                commentValue = "PlayCount#" + str(track['PlayedCount']) \
                               + ",SkipCount#"\
                               + str(track['SkippedCount'])\
                               + ",Rating#"\
                               + str(theRating)\
                               + ",AddedDate#"\
                               + str(track['DateAdded'])\
                               + ",PlayedDate#"\
                               + str(track['PlayedDate'])\
                               + ",SkippedDate#"\
                               + str(track['SkippedDate'])

                #Set rating
                if (theOldRating != theRating) and not simulate:
                    objTrack.Rating = theRating
                    updateNeeded = True
                    if createPlaylist:
                        playlist.AddTrack(objTrack)

                if createPlaylist and topplaylistcount < topcount:
                    topplaylist.AddTrack(objTrack)
                    topplaylistcount += 1

                #Backup Values to comment
                if objTrack.Comment != commentValue:
                    updateNeeded = True
                    if backupComments:
                        objTrack.Comment = commentValue

                #Logging
                if updateNeeded:
                    updated = updated + 1
                    logging.debug("------------------------------" + str(updated) + "------------------------------")
                    logging.debug(chr(9) + "ID: " + chr(9) + chr(9) + str(track['trackDatabaseID']))
                    logging.debug(chr(9) + "Title: " + chr(9) + chr(9) + track['Name'])
                    logging.debug(chr(9) + "Artist: " + chr(9) + track['Artist'])
                    logging.debug(chr(9) + "Length: " + chr(9) + str(track['Time']))
                    logging.debug(chr(9) + "Played: " + chr(9) + str(track['PlayedCount']))
                    logging.debug(chr(9) + "Last Played: " + chr(9) + str(track['PlayedDate']))
                    logging.debug(chr(9) + "Skipped: " + chr(9) + str(track['SkippedCount']))
                    logging.debug(chr(9) + "Last Skipped: " + chr(9) + str(track['SkippedDate']))
                    logging.debug(chr(9) + "Date added: " + chr(9) + str(track['DateAdded']))
                    logging.debug(chr(9) + "Old Rating: " + chr(9) + str(theOldRating))
                    logging.debug(chr(9) + "New Rating: " + chr(9) + str(theRating))
                    if theRating > theOldRating:
                        up = up + 1
                    elif theRating < theOldRating:
                        down = down + 1
                    else:
                        equal = equal + 1

                if objTrack.AlbumRating != 1:
                    objTrack.AlbumRating = 1

        logging.info(str(updated) + " File ratings updated.")
        logging.info(str(up) + " File ratings goes up.")
        logging.info(str(down) + " File ratings goes down.")
        logging.info(str(equal) + " File ratings keeps equal.")

        logging.info("Done!")

    else:
        logging.info("Script aborded because no tracks are available.")

    stop = timeit.default_timer()
    runtime = stop - start
    logging.info("Script runtime: " + str(runtime))


if __name__ == '__main__':
    main()
