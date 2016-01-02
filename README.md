autorate (Windows)
==================

Windows version of the program autorate (https://code.google.com/p/autorate/) with an improved algorithm.

iTunes rating made easy: Automated track rating from play statistics

This application automatically sets the rating for all tracks in your iTunes library according to how you listen to them.

Tracks you listen to more will have a higher rating; tracks you neglect will have a lower rating.

This is particularly useful for automatically creating playlists, or transferring only your favourite tracks onto your iPod.

Install instruction
==================
Create a smart playlist with the name "MusicOnly", which only contains music on your computer ("Media Kind" is "Music" and "Location" is "on this computer").
Double click autorate.cmd. If you prefer python (script runs faster) just call autorate.py (C:\Python34\pythonw.exe C:\autorate.py).

You can also configure the script to run every day/week when you use the windows task scheduler.

If you want to use half-star rating you have to activate it first:
- Make sure you are closed out of the iTunes application.
- Hold the Windows Key and press “R” to bring up the Run dialog box.
- Type “C:\Program Files\iTunes\iTunes.exe” /setPrefInt allow-half-stars 1 then press “Enter“. If that doesn’t work, try “C:\Program Files (x86)\iTunes\iTunes.exe” /setPrefInt allow-half-stars 1 then press “Enter“.
