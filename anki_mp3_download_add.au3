#include <InetConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Json.au3>

; CTRL + M
HotKeySet("^m","PasteAudioIntoCard")
; CTRL + .
HotKeySet("^.","Quit")

Global $sAudioURL = ""
Global $jsonDictAudioURL = ""
Global $sWordGlobal = "" ; for DownloadAudioFile
Global $sAudioFilePath = ""

Main()
Func Main()
    Call("LoadDictionaryIntoRAM")

    ; start endless loop for hotkeys
    While 1
        Sleep(100)
    WEnd
EndFunc ;==> Main

Func PasteAudioIntoCard()
    ; Retrieve the data stored in the clipboard.
    Local $sWord = StringLower(ClipGet())
    $sWordGlobal = $sWord;
    Call("ExtractAudioURLFromDict", $sWord)
    Call("DownloadAudioFile")
    Call("SwitchFocusToAnkiAndAddSound")
    ; Create sound element & put it into clipboard
EndFunc ; ==> PasteAudioIntoCard

; ExtractAudioURLFromDict()
Func ExtractAudioURLFromDict($sWord)
    $sAudioURL = Json_Get($jsonDictAudioURL, '[' & $sWord & ']')
    
    ; check if word exists in dictionary
    Local $urlLength = StringLen($sAudioURL)
    If $urlLength = 0 Then
        MsgBox($MB_SYSTEMMODAL, "Title", "Didn't found that word in the dictionary!", 10)
    EndIf
EndFunc ; ==> ExtractAudioURLFromDict

; DownloadAudioFile()
Func DownloadAudioFile()
    ; Save the downloaded file to the temporary folder.
    Local $sPath = "C:\Users\kevin\Desktop\english\audio\"
    ; File Name should have space
    Local $sEscapedFileName = StringReplace($sWordGlobal, " ", "-") & ".mp3"
    Local $sFilePath = $sPath & $sEscapedFileName
    $sAudioFilePath = $sFilePath

    ; MsgBox($MB_SYSTEMMODAL, "Title", "Filepath: "  & $sFilePath, 10) ; for debugging

    ; Download the file by waiting for it to complete. The option of 'get the file from the local cache' has been selected.
    Local $iBytesSize = InetGet($sAudioURL, $sFilePath, $INET_FORCERELOAD)
EndFunc   ;==>DownloadAudioFile

Func LoadDictionaryIntoRAM()
    Local $sDict = "C:\Users\kevin\Desktop\english\english.json"
    Local $jsonFile = FileRead($sDict)
    Local $jsonObject = Json_Decode($jsonFile)
    $jsonDictAudioURL = $jsonObject
    MsgBox($MB_SYSTEMMODAL, "Title", "Successfully started application and loaded dictionary!", 10)
EndFunc ; ==> LoadDictionaryIntoRAM

Func SwitchFocusToAnkiAndAddSound()
    ; Wait 5 seconds for the Notepad window to appear.
    Local $hwndAnkiAdd =  WinWaitActive("Add", "", 3)
    ; open the add media dialogue
    Send("{F3}")
    Local $addMediaAnkiAdd = WinWaitActive("Add Media", "", 3)
    ; click on the filename input area & enter path
    ControlClick("", "", "[ClassnameNN:Edit1; INSTANCE:1; ID:1148]")
    Sleep(1000)
    Send($sAudioFilePath & "{ENTER}")
EndFunc ; ==> SwitchFocusToAnkiAndAddSound

Func Quit()
    Exit
EndFunc ; ==> Quit