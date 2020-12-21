
# TorrentScript
qBittorrent post processing script with import to Radarr and Medusa by using a temporary directory.  
Caveat: I'm not a programmer, so use at own risk :)

## Description
Only 1 instance of the script will be running and if other downloads complete during execution they will wait on the first script to finish. This is to prevent a system overload.

Script will do the following:  
1. Unpack or copy to Temporary Processing folder defined in config.json (copy from config-sample.json)
2. If torrent label is TV show as defined in JSON Label
    - Will strip all subtitle languages not in JSON WantedLanguages
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
3. If torrent label is Movie as defined in JSON Label
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
4. If there is another Label will just unpack or extract
5. Clean up Temporary folder
6. If there is no Label set or label is `NoProcess` script will exit
7. Send an email with result

**Success**  
![Success](https://i.imgur.com/Bjp5ggF.png)  
**Medusa wrong host**  
![Medusa wrong host](https://i.imgur.com/9BrtJ6z.png)  
**Unrar error**  
![Unrar error](https://i.imgur.com/TYvRUXL.png)  

## Installation
Need to create a config.json in the root folder of the script, you can copy the config-sample.json.  
Script needs to be called from qBittorrent after download is finished with the following command
```
powershell "C:\Scripts\TorrentScript\TorrentScript.ps1" -DownloadPath '%D' -DownloadName '%N' -DownloadLabel '%L' -TorrentHash '%I'
```
![qBittorrent settings page](https://i.imgur.com/8TWZyEY.png)

The following external tools need to be available and the path defined in the `config.json`
 - [WinRar](https://www.rarlab.com/download.htm)
 - [MKVMerge](https://mkvtoolnix.download/)
 - [MKVExtract](https://mkvtoolnix.download/)
 - [Subtitle Edit](https://github.com/SubtitleEdit/subtitleedit)
 - [Subliminal](https://github.com/Diaoul/subliminal)
 - [MailSend](https://github.com/muquit/mailsend-go)
