Attribute VB_Name = "URLDownload"

Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
ByVal pcaller As LongPtr, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As LongPtr, _
ByVal lpfnCB As LongPtr) As LongPtr





