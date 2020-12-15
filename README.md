
# TorrentScript
qBittorrent post processing script with import to Radarr and Medusa by using a temporary directory
Caveat: I'm not a programmer, so use at own risk :)

## Description
Script will do the following:  
1. Unpack or copy to Temporary Processing folder defined in config.json (copy from config-sample.json)
2. If TV show as defined in JSON Label
    - Will strip all audio and subtitle languages not in JSON WantedLanguages, currently still using external Python script [MKVstrip](https://github.com/jobrien2001/mkvstrip)
    - Extract all SRT subtitles in JSON WantedLanguages
    - Rename all srt files from alpha3 to alpha2 codes as defined in JSON LanguageCodes
	    - File.3.en.srt
	    - File.en.srt
	    - File.nl.srt
    - Start [Subliminal](https://github.com/Diaoul/subliminal) to see if there are missing subtitles we can download 
    - Clean up the srt subtitles using [Subtitle Edit](https://github.com/SubtitleEdit/subtitleedit)
	    - Remove Hearing Impaired.
	    - Fix Common errors.
     - Start Medusa Import
 - If Movie as defined in JSON Label
    - Extract all SRT subtitles in JSON WantedLanguages
    - Rename all srt files from alpha3 to alpha2 codes as defined in JSON LanguageCodes
	    - File.3.en.srt
	    - File.en.srt
	    - File.nl.srt
    - Start [Subliminal](https://github.com/Diaoul/subliminal) to see if there are missing subtitles we can download 
    - Clean up the srt subtitles using [Subtitle Edit](https://github.com/SubtitleEdit/subtitleedit)
	    - Remove Hearing Impaired.
	    - Fix Common errors.
    - Start Radarr Import
3. Clean up Temporary folder


## Installation
Need to create a config.json in the root folder of the script, you can copy the config-sample.json
Script needs to be called from qBittorrent after download is finished with the following command
```
powershell "C:\Scripts\TorrentScript\TorrentScript.ps1" -DownloadPath '%D' -DownloadName '%N' -DownloadLabel '%L' -TorrentHash '%I'
```
![qBittorrent settings page](https://i.imgur.com/8TWZyEY.png)

