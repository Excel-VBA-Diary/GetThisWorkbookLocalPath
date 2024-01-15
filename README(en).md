K# GetThisWorkbookLocalPath
# Resolve the problem of ThisWorkbook.Path returning a URL in OneDrive  
Created: December 11, 2023  
Last updated: January 15, 2024  
  
## Problem to be solved 
  
There is a problem with ThisWorkbook.Path returning a URL when Excel VBA runs on OneDrive. This is inconvenient because you cannot get your own local path and you cannot use the Dir function or even FileSystemObject.

Several methods have been proposed to solve this problem. For personal OneDrive, string conversion can be used, but when manipulating SharePoint files via OneDrive for Business, the method of converting URL paths to local paths using only string processing will not work. The conversion of the tenant code in the URL to a tenant name, for example, is required and cannot be solved by string processing.

There are two ways to use SharePoint and Teams files in OneDrive: "Synchronize" and "Add shortcut to OneDrive", but the paths on the local drive are different for each method, and furthermore, it is not possible to know from the URL path which method is used for synchronization. This is why OneDrive is not available.

For these reasons, it is virtually impossible in OneDrive for Business to convert the URL returned by ThisWorkbook.Path to a local path through string processing.
  
## Proposed Solutions

Four different methods are proposed here.    
(Part 1)  Use "Show Recently Opened Items"  
(Part 2)  Use "Open Explorer"  
(Part 3)  Use "System.Windows.Forms.SendKeys"  
(Part 4)  Use GetLocalPath function  
  
The source code for (1) through (3) is available in this repository. The files exported from the standard modules are posted as they are, so please import them or copy and paste the necessary parts.
(Part 1)  Module1.bas  
(Part 2)  Module2.bas  
(Part 3)  Module3.bas  
  
The source code of the GetLocalPath function in (4) can be found in the following repository.  
   [GetLocalPath](https://github.com/Excel-VBA-Diary/GetLocalPath)  
  
## Proposed Solution (Part 1)   
  
The source code is Module1.bas. The function to get the local path is GetThisWorkbookLocalPath1().  

This code uses the "Show Recently Opened Items" feature, which automatically records recently opened files and folders as linked files (LNK files) in the folder shown below.  
  
    C:\Users\<user-name>\AppData\Roaming\Microsoft\Windows\Recent  
  
You can get the path on the local drive by getting the link to this link file.  
  
To use "Show Recent Items", go to "Personalization" -> "Start" in Windows settings and switch "Show recently opened items .." option turn on.  

For Windows 11, "Show recently opened items in Start, Jump List, and File Explorer"  

For Windows 10, "Show recently opened items in Jump Lists on Start or the taskbar and in file explorer Quick Access"  

If this setting is off, the link is as described above. If this setting is off, GetThisWorkbookLocalPath1() returns an empty string (zero-length string) because the linked file (LNK file) described above is not recorded.  
  
### How to know if "Show Recently Opened Items" is enabled or disabled    
  
Before calling GetThisWorkbookLocalPath1(), you can check by reading the registry key to know if "Show Recently Opened Items" is turned on. The function for this is Is_Start_TrackDocs().  
  
This function reads the value of Start_TrackDocs in the registry key shown below, and returns its value.

    HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\  

## Proposed Solution (Part 2)   
  
The source code is Module2.bas. The function to get the local path is GetThisWorkbookLocalPath2().  

This code retrieves the local path from Explorer showing the folder where the currently open Excel file (i.e. ThisWorkbook) is located.  

Specifically, the Window object from the Explorer window is obtained, and the absolute path (URI) "file:///C:/Users/.../...//OneDrive.../..." is obtained with the LocationURL property.  

This absolute path (URI) is encoded and must be decoded; the DecodeURL_ASCII() function is for that purpose. Only certain ASCII characters are decoded by this function.  
  
The full-set version of the DecodeURL function is also written for reference. This is the inverse function of the ENCODEURL function, which is an Excel worksheet function. It is prepared in case it is encoded in the future.   

Since GetThisWorkbookLocalPath2() obtains information from the explorer in this way, the information will not be available if the corresponding explorer is closed. In this case, GetThisWorkbookLocalPath2() returns an empty string (zero-length string).  

Note that if the files are placed directly under OneDrive or OneDrive for Business (root folder), ThisWorkbook.Path returns a specific URL pattern for each, so even without obtaining information from Explorer, OneDrive will return Environ("OneDrive") for OneDrive and Environ("OneDriveCommercial") for OneDrive for Business to correspond to the local path.  

## Proposed Solution (Part 3)   
  
The source code is Module3.bas. The function to get the local path is GetThisWorkbookLocalPath3().  

This code sends keystrokes to the currently opened Excel file (i.e. ThisWorkbook) itself by SendKeys to get the local path.  

For an Excel file on OneDrive, you can get the local path to the clipboard by "File" tab -> "Info" -> "Copy Local Path".  

Keystroke is "Alt" -> "F" -> "I" -> "L". After this, "Alt" -> "H" -> "Up" -> "Enter" is sent to go back to the original home tab.  

So actual keystroke is "Alt" -> "F" -> "I" -> "L" -> "Alt" -> "H"-> "up" -> "Enter".  

SendKeys cannot use VBA's Application.SendKeys method. This is because the Application.SendKeys method does not work well for manipulating its own ribbon tab.  

This problem can be solved by sending keystrokes to Excel externally via PowerShell; the script that sends the keystrokes to be executed by PowerShell is embedded in the code.  

Actually, we would like to send the "Esc" key to return to the original home tab, but depending on the timing, VBA may be interrupted, so we avoid sending the "Esc" key.  

The timing for sending keystrokes is specified by the Start-Sleep cmdlet in the script. Although the timing is set to a reasonable level, it may be necessary to adjust the Start-Sleep timing depending on the Windows or Office environment.  

Please note that it is normal for the window to change when keystrokes are sent. If the keystroke submission fails, GetThisWorkbookLocalPath3() returns an empty string (zero-length string).  

If the original source code makes heavy use of ThisWorkbook.Path, simply replacing ThisWorkbook.Path with GetThisWorkbookLocalPath3() will result in frequent screen movement, so it is recommended to use a global variable such as It is recommended to minimize the number of calls to GetThisWorkbookLocalPath3() as much as possible.  

## Proposed Solution (Part 4)    
  
The solution uses the OneDrive mount information in the Windows registry. This mount information is located under the following subkeys:  
```
\HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive
```
A description and source code for the GetLocalPath function can be found [here](https://github.com/Excel-VBA-Diary/GetLocalPath).   
To convert the URL path returned by ThisWorkbook.Path to a local path using this function, use the following: 
```
Dim localPath As String
localPath = GetLocalPath(ThisWorkbook.Path)
```  

## Afterword 

OneDrive, OneDrive for Business, or Teams or SharePoint can be used as a local drive by "adding a shortcut to OneDrive". This has the advantage of being used without web access in mind.  
On the other hand, VBA is ineffective for these new mechanisms. This proposal is one way to compensate for that, but to begin with, VBA has not had any major updates since 2012, and it is hard to deny the feeling that it has been left behind in response to the new solutions Microsoft is proposing.  
  
Even if the problem of ThisWorkbook.Path returning URLs has been solved, the use of SharePoint files with "Add shortcut to OneDrive" may still require CheckOut/CheckIn exclusivity control is necessary in some cases.  
  
Of course, VBA has a CheckOut/CheckIn method, but it is not simple because it requires flow control including retry processing.
In this sense, this proposal should be regarded as a temporary measure in case there is no other solution.  

## LICENSE 

This code is available under the MIT License.  

[EOF]
