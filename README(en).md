K# GetThisWorkbookLocalPath
# Resolve the problem of ThisWorkbook.Path returning a URL in OneDrive  
Last update: December 15, 2023
  
## Problem to be solved 
  
We have a problem with ThisWorkbook.Path returning a URL when run Excel VBA on OneDrive. This is inconvenient because my own local path can not be gotten and even FileSystemObject is not available.  
  
Several methods have been proposed to solve this problem, but the method of converting URL paths to local paths by string processing may not work. In particular, OneDrive for Business requires converting tenant codes in URLs to tenant names, which cannot be solved by string processing.  
  
There are two ways to sync Teams and SharePoint files: "Sync Client" and "Add Shortcut to OneDrive", but each of them has a different path on the local drive, and it is not possible to know from the URL path which method is used to sync.  
  
For these reasons, it is virtually impossible to convert the URL returned by ThisWorkbook.Path to a local path by string processing.  
  
## Proposed Solutions

Three different methods are proposed here.  
The first is to use "Show Recently Opened Items," the second is to use the "Open Explorer", and the third is to use SendKeys.  
Each has its own prerequisites and should be used alone or in combination as needed.  
  
The source code is provided as-is in the file that exports the standard module, so you can either import it or copy and paste the necessary parts.  
  
The three different methods are listed in the following three files.  

(Part 1) Module1.bas : Method using "Show Recently Opened Items"  
(Part 2) Module2.bas : Method using "Open Explorer"  
(Part 3) Module3.bas : Method using [System.Windows.Forms.SendKeys]  

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

This absolute path (URI) is encoded and must be decoded using the DecodeURL() function. Since only certain ASCII characters are encoded, we have also written a simplified version of the DecodeURL_ASCII() function for reference.  

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

## Afterword 

OneDrive, OneDrive for Business, or Teams or SharePoint can be used as a local drive by "adding a shortcut to OneDrive". This has the advantage of being used without web access in mind.  
On the other hand, VBA is ineffective for these new mechanisms. This proposal is one way to compensate for that, but to begin with, VBA has not had any major updates since 2012, and it is hard to deny the feeling that it has been left behind in response to the new solutions Microsoft is proposing.  
  
Even if the problem of ThisWorkbook.Path returning URLs has been solved, the use of SharePoint files with "Add shortcut to OneDrive" may still require CheckOut/CheckIn exclusivity control is necessary in some cases.  
  
Of course, VBA has a CheckOut/CheckIn method, but it is not simple because it requires flow control including retry processing.
In this sense, this proposal should be regarded as a temporary measure in case there is no other solution.  

## LICENSE 

This code is available under the MIT License.  

[EOF]
