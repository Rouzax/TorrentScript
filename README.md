# TorrentScript
qBittorrent post proccessing script with import to Radarr and Medusa by using a tempory directory

## Description
Script will do the following  
1. Unpack or copy to Temporary Proccessing folder defined in config.json (copy from condig-sample.json)
2. 

## Installation
Script needs to be called from qBittorrent after download is finished with the following command
```
powershell "C:\Scripts\TorrentScript\TorrentScript.ps1" -DownloadPath '%D' -DownloadName '%N' -DownloadLabel '%L' -TorrentHash '%I'
```

