'::::::::::::::::::::::::::::::::::::::::::::::::::
OperatingSystem: WIN7 an above
PS Version:      2.0 and above
OSEdition:       x86;x64
Description:     RedWall - Reddit Wallpapers
                 Script was created to download images from subreddit frontpages.
Scriptversion:   v1.1
Autor:           Bela.Sajtos@gmail.com
RebootRequired:  NO
Arguments:       NO
Comments:        Only downloads images with direct image links
                 Valid image resolution is added to the filenames of the downloaded images, like "[1920x1080] Art.jpg" to ease manual selection later
                 NSFW subreddits will not work
                 According to feedbacks: script will not work properly if proxy is used. I could not test and solve it.
Usage:           Before run: modify the source code for your need (manually in the PERSONALIZATION section)
                 It is rcommended to create a daily scheduled task executing the script
                 
::::::::::::::::::::::::::::::::::::::::::::::::::'

#--------------------------------------------------------------------
#PERSONALIZATION - modify these parameters for your needs!
#--------------------------------------------------------------------

#Pictures will be saved into this destination folder. Path must end with "\"
$Wallpapers = "C:\temp\wallpapers\"

#Specifies how many of the top posts from subreddit frontpages will be collected. Max 25.
$Limit = 10

#Comment out or delete not needed subreddits, or add more which are not listed here
$subreddits = @()

$subreddits += "https://www.reddit.com/r/wallpaper/"
$subreddits += "https://www.reddit.com/r/wallpapers/"
$subreddits += "https://www.reddit.com/r/MinimalWallpaper/" #Minimalist style
$subreddits += "https://www.reddit.com/r/WidescreenWallpaper/" #3440x1440, 2560x1080 & 21:9 wallpapers.
$subreddits += "https://www.reddit.com/r/WQHD_Wallpaper/" #high quality wallpaper
$subreddits += "https://www.reddit.com/r/EarthPorn/"
$subreddits += "https://www.reddit.com/r/carporn/"
$subreddits += "https://www.reddit.com/r/CityPorn/"
$subreddits += "https://www.reddit.com/r/ExposurePorn/" #Long Exposure Photography
$subreddits += "https://www.reddit.com/r/HumanPorn/" #well, actually this is NOT porn, SFW
#$subreddits += "https://www.reddit.com/r/iWallpaper"           #i
#$subreddits += "https://www.reddit.com/r/VerticalWallpapers"   #for phones


#--------------------------------------------------------------------
# SCRIPT
#--------------------------------------------------------------------


#Check if the Destination folder exist. If not, create it 
if (!(Test-Path $Wallpapers))
{
    New-Item -Path $Wallpapers -ItemType directory -Force | Out-Null
    Write-Host "Destination folder was created:"$Wallpapers
}

$FolderContent = Get-ChildItem $Wallpapers | Select-Object name
$Links = @()

Write-Host "`n***********LOADING SUBREDDITS***********`n"

foreach ($subreddit in $subreddits)
{
    Write-Host "Subreddit "$subreddit "is loading..."
    
    #Loading the the defined subreddits
    $ie = New-Object -ComObject "InternetExplorer.Application"
    $ie.Navigate($subreddit)
    while ($ie.busy) { Start-Sleep -Milliseconds 100 }
    
    #Collecting the first [$Limit] links from the actual subreddit
    $Links += $ie.Document.getElementsByTagName("a") | Where-Object { $_.className -like "title*may-blank*" } | Select-Object tabIndex, nameprop, host, className, innerhtml, href | Select-Object -First $Limit
    
    $ie.Quit()
}


#Remove duplicated entries from different subreddits, but with the same target image link
$Links = $Links | Sort-Object -Property href -Unique

Write-Host "`n***********DOWNLOADING IMAGES***********`n"

foreach ($Link in $Links)
{
    #check if the link is direct link to an image, or not
    if ($Link.nameProp.split(".").count -gt 1)
    {
        #directlink
        $URL = $link.href
        $FileExtension = "." + $Link.nameProp.split(".")[-1] #gets file extension, like ".jpg"
        if ($URL.EndsWith("?1")) #some link/filename ends with "?1", this has to be removed
        {
            $URL = $URL.Replace("?1", "")
            $FileExtension = $FileExtension.Replace("?1", "")
        }
        $FileName = $link.innerHTML.replace(" ", "_") -replace "([^a-zA-Z._])" #replaces spaces with "_", then removes all characters which are not letters, ".", or "_"
        $FileName += "_" + $Link.nameProp.split(".")[0] #reused titles could cause skipping files because of duplication, like several posts with title "Art", so the original gibberish filename is added in the end as a unique identifier
        $FileName = $Filename.Replace("_x", "") #in the link names image resolution (1920x1080) is often listed, but because of the earlier formatting there will be a "_x" leftover
        if ($FileName.Length -gt 100) { $FileName = $FileName.Substring(0, 100) } #long titles can cause exception, fully qualified name to a file must be less than 260 characters, we wont need more than a hundred
        $FileName += $FileExtension #adding the extension to the filename
        $Fullpath = $Wallpapers + $FileName
    }
    
    else
    {
        #if the reddit post is NOT a direct link post to an image, then it is skipped...
        Write-Host "Link is not a direct link, it will be skipped - ("$Link.href")"
        continue
    }
    
    
    #if the image was previously downloaded, then skip the loop and continue with the next link
    if ($FolderContent | Where-Object { $_ -like "*$FileName*" })
    {
        Write-Host "Image already exist from ("$Link.href")"
        continue
    }
    
    
    #Download the raw images. Valid image resolution is added to the filenames of the downloaded images, like "[1920x1080] Art.jpg" to ease manual selection later
    if (!(Test-Path $fullpath))
    {
        Write-Host "Image is being downloaded from ("$Link.href")"
        
        Invoke-WebRequest $URL -OutFile $fullpath -ErrorAction SilentlyContinue
        
        if (Test-Path $Fullpath)
        {
            #Get metadata from the downloaded image and rename the downloaded file with adding resolution to the filename, like "[1920x1080] Art.jpg"
            $img = New-Object -comObject WIA.ImageFile
            $img.LoadFile($fullpath)
            $newFileName = "[" + $img.Width + "x" + $img.Height + "] " + $fullpath.Split("\")[-1]
            
            Rename-Item -Path $fullpath -NewName $newFileName
        }
    }
}
