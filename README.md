# randomfilegenerator

## Project Description

This project contains a PowerShell script to generate files of various sizes with random GUID content for SharePoint 201x performance testing. 

This project contains several script to generate files with random content in it.
The idea is to create txt files for SharePoint performance testing and search purposes. A detailed reason behind the project can be found on these blog entries: 
- https://blogs.technet.microsoft.com/zsoltilles/2017/02/26/generate-random-content-for-sharepoint/
- https://blogs.technet.microsoft.com/zsoltilles/2017/03/05/generate-random-content-for-sharepoint-2/
- https://blogs.technet.microsoft.com/zsoltilles/2017/03/12/generate-random-content-for-sharepoint-3/
- https://blogs.technet.microsoft.com/zsoltilles/2017/03/28/generate-random-content-for-sharepoint-4/

## RandomGUIDFileGenerator.ps1 
- Generate simple text files.
- With content constructed with GUIDs.
- Good for upload tests.


## RandomWordFileGenerator.ps1 
- Generate Word files.
- Requires a dictionary that contains words
- Requires a file that contains separators based on their weight. (See GenerateSeparatorFile.ps1)
- Optionally uses a list of users to randomly change the Creator and LastModifiedBy fields of the document.
- Optionally uses a list of dates to randomly change the Created and Modified fields of the document.
- Optionally uses a template Word file.


## GenerateSeparatorFile.ps1 
- To generate a file that contains separators.
- Without parameter the file creates a file with separators weighted by their occurrence in the English language.


## RandomExcelFileGenerator.ps1 
- Requires a dictionary that contains words
- Optionally uses a list of users to randomly change the Creator and LastModifiedBy fields of the document.
- Optionally uses a list of dates to randomly change the Created and Modified fields of the document.
- Optionally uses a template Excel file.


## RandomPowerPointFileGenerator.ps1 
- Requires a dictionary that contains words
- Optionally uses a list of users to randomly change the Creator and LastModifiedBy fields of the presentation.
- Optionally uses a list of dates to randomly change the Created and Modified fields of the document.
- Optionally uses a template PowerPoint file.
