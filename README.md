
# TorrentScript
A qBittorrent post-processing script designed to import media into Radarr (Movies) and Medusa or Sonarr (TV Shows) while handling subtitles, renaming, and cleanup. 
Caveat: I'm not a programmer, so use at own risk :)

## Description
The script processes torrent downloads by:
1. Unpacking or copying files to a **Temporary Processing folder** defined in `config.json`.
2. If the torrent label is a **TV show** (as defined in `config.json` under `Label.TV`):
    - Removes unwanted subtitle languages from MKV files.
    - Extracts subtitles in the **wanted languages** from MKV files.
    - Attempts to download missing subtitles from **OpenSubtitles**.
    - Converts subtitle language codes from **ISO 639-2 (3-letter)** to **ISO 639-1 (2-letter)** as per `LanguageCodes.json`.
    - Cleans up SRT subtitles using [Subtitle Edit](https://github.com/SubtitleEdit/subtitleedit):
        - Removes **Hearing Impaired** subtitles.
        - Fixes **common errors**.
    - **Starts Medusa or Sonarr import** (as defined in `config.json` under `ImportPrograms.TV`).
3. If the torrent label is a **Movie** (as defined in `config.json` under `Label.Movie`):
    - Performs the same subtitle processing as above.
    - **Starts Radarr import**.
4. If there is another **unrecognized label**, the script will simply unpack/extract the files without further processing.
5. Cleans up the **temporary folder** after processing.
6. If no label is set or if the label is `NoProcess`, the script exits without processing.
7. Sends an **email notification** with the result.

**Success**  
![Success](https://i.imgur.com/Bjp5ggF.png)  

**Medusa wrong host**  
![Medusa wrong host](https://i.imgur.com/9BrtJ6z.png)  

**Unrar error**  
![Unrar error](https://i.imgur.com/TYvRUXL.png)  

---

## Installation

1. **Create a `config.json`** file in the root directory of the script.  
   You can copy the provided `config-sample.json` and modify it as needed.

2. **Set up qBittorrent to trigger the script after downloads finish**  
   The script should be triggered with the following command:
   ```sh
   powershell "C:\Scripts\TorrentScript\TorrentScript.ps1" -DownloadPath '%R' -DownloadLabel '%L' -TorrentHash '%I'
   ```
   ![qBittorrent settings page](https://i.imgur.com/8TWZyEY.png)

3. **Enable per-torrent folders in qBittorrent**  
   This setting ensures that each torrent is placed inside its own folder.
   ![qBittorrent Folder settings page](https://i.imgur.com/Uq6bOBP.png)

4. **Configure Remote Path Mapping in Radarr/Sonarr**  
   Radarr and Sonarr should **not** directly access qBittorrent's completed downloads folder.
   
   Example:
   - qBittorrent root download path: `C:\Torrent\Downloads\`
   - Radarr **Remote Path Mapping**: `C:\Torrent\Radarr\`
   
   ![Radarr Remote Path Mapping settings page](https://i.imgur.com/qL0aOKl.png)

   This ensures that Radarr and Sonarr only imports processed media via this script.

---

## Understanding `RemotePath` in `config.json`

The `RemotePath` setting in `config.json` ensures that Medusa, Radarr, and Sonarr receive the **correct file paths** when the script is executed from a different machine than where these applications are running.

### **How It Works**
- If the script is triggered from a **jumphost** or another system different from the one hosting Medusa, Radarr, or Sonarr, the local file paths may not match what these applications expect.
- The `RemotePath` setting provides the **correct path** that Medusa, Radarr, or Sonarr should use to locate the processed files, even when the script itself is running from another machine.
- This is especially useful in setups using **shared storage** where multiple machines have access to the same files but reference them with different paths.

### **Example Scenario**
Assume your downloads are stored on a **network share** at `\\NAS\Downloads\`, but:
- The **jumphost** (where the script runs) mounts this as `D:\Downloads\`
- The **server running Medusa** accesses the same location as `C:\TEMP\Torrent\Medusa`
- The **server running Radarr** expects `C:\TEMP\Torrent\Radarr`


# Required External Tools

This script relies on the following external tools, which need to be installed and **defined in `config.json`**:

| Tool                                                          | Purpose                         |
| ------------------------------------------------------------- | ------------------------------- |
| [WinRAR](https://www.rarlab.com/download.htm)                 | Extracting archives             |
| [MKVMerge](https://mkvtoolnix.download/)                      | Merging MKV streams             |
| [MKVExtract](https://mkvtoolnix.download/)                    | Extracting subtitles from MKV   |
| [Subtitle Edit](https://github.com/SubtitleEdit/subtitleedit) | Cleaning and renaming subtitles |

---

## Advanced Features & Additional Notes

### **1. Multi-Instance Handling**
- The script ensures that only **one instance** runs at a time.
- If multiple torrents complete while a process is already running, they **wait** until the first script execution is finished.

### **2. Subtitle Processing**
- Extracts, cleans, and renames subtitles for supported languages.
- Uses `LanguageCodes.json` to convert 3-letter language codes to 2-letter codes.

### **3. Logging & Error Handling**
- Logs are stored in a defined path (`LogArchivePath` in `config.json`).
- Errors are logged and can be sent via email notifications.

### **4. Medusa/Radarr/Sonarr Import**
- The script interacts with Medusa or Sonarr and Radarr via their **API** using the provided API keys.
- It **ensures that only properly processed files** are imported.