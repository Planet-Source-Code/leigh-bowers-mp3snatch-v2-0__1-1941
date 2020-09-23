<div align="center">

## MP3Snatch v2\.0


</div>

### Description

This revised code finally supports the MP3 "Genre" tag (WinAMP 2.22+ compliant).

Loads of you have emailed me requesting this facility - I think this is the first such VB code to support it! Easy to implement and the new genre routine is compact. It was a right pain in the arse collecting the Genre descriptions ;-) Note:- A demonstration app is availble from my homepage...
 
### More Info
 
Cut and paste this into a VB Class Module (clsMP3Snatch.cls for exmaple).

For demonstration code to acompany this class module, drop by my home page...

This "Split" instruction *may* be VB6 only - I'm not 100% sure.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Leigh Bowers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/leigh-bowers.md)
**Level**          |Unknown
**User Rating**    |4.2 (67 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/leigh-bowers-mp3snatch-v2-0__1-1941/archive/master.zip)





### Source Code

```
Option Explicit
' Title:  MP3 Snatch
' Author:  Leigh Bowers
' Version: 2.0
' Released: 1st June 1999
' WWW:   http://www.esheep.freeserve.co.uk/compulsion/index.html
' Email:  compulsion@esheep.freeserve.co.uk
' News:   Added "Genre" functionality (WinAMP compliant)
Private sFilename As String
Private Type Info
  sTitle As String * 30
  sArtist As String * 30
  sAlbum As String * 30
  sComment As String * 30
  sYear As String * 4
  sGenre As String * 21 ' NEW
End Type
Private MP3Info As Info
Public Property Get Filename() As String
  Filename = sFilename
End Property
Public Property Let Filename(ByVal sPassFilename As String)
  Dim iFreefile As Integer
  Dim lFilePos As Long
  Dim sData As String * 128
  Dim sGenreMatrix As String
  Dim sGenre() As String
  ' Genre
  sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
  ' Build the Genre array (VB6+ only)
  sGenre = Split(sGenreMatrix, "|")
  ' Store the filename (for "Get Filename" property)
  sFilename = sPassFilename
  ' Clear the info variables
  MP3Info.sTitle = ""
  MP3Info.sArtist = ""
  MP3Info.sAlbum = ""
  MP3Info.sYear = ""
  MP3Info.sComment = ""
  ' Ensure the MP3 file exists
  If Dir(sFilename) = "" Then Exit Property
  ' Retrieve the info data from the MP3
  iFreefile = FreeFile
  lFilePos = FileLen(sFilename) - 127
  Open sFilename For Binary As #iFreefile
    Get #iFreefile, lFilePos, sData
  Close #iFreefile
  ' Populate the info variables
  If Left(sData, 3) = "TAG" Then
    MP3Info.sTitle = Mid(sData, 4, 30)
    MP3Info.sArtist = Mid(sData, 34, 30)
    MP3Info.sAlbum = Mid(sData, 64, 30)
    MP3Info.sYear = Mid(sData, 94, 4)
    MP3Info.sComment = Mid(sData, 98, 30)
    MP3Info.sGenre = sGenre(Asc(Mid(sData, 128, 1)))
  End If
End Property
Public Property Get Title() As String
  Title = RTrim(MP3Info.sTitle)
End Property
Public Property Get Artist() As String
  Artist = RTrim(MP3Info.sArtist)
End Property
Public Property Get Genre() As String
  Genre = RTrim(MP3Info.sGenre)
End Property
Public Property Get Album() As String
  Album = RTrim(MP3Info.sAlbum)
End Property
Public Property Get Year() As String
  Year = MP3Info.sYear
End Property
Public Property Get Comment() As String
  Comment = RTrim(MP3Info.sComment)
End Property
```

