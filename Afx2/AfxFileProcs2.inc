' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxFileProcs2.inc
' Purpose:  Windows file controls
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#pragma once
#INCLUDE ONCE "Afx2/DWSTRING.inc"
USING Afx2

NAMESPACE Afx2

' ########################################################################################
'                                     *** FILES ***
' ########################################################################################

' ========================================================================================
' Searches a directory for a file or subdirectory with a name that matches a specific name
' (or partial name if wildcards are used).
' Parameter:
' - pwszFileSpec: The directory or path, and the file name, which can include wildcard
'   characters, for example, an asterisk (*) or a question mark (?).
'   This parameter should not be NULL, an invalid string (for example, an empty string or a
'   string that is missing the terminating null character), or end in a trailing backslash (\).
'   If the string ends with a wildcard, period (.), or directory name, the user must have
'   access permissions to the root and all subdirectories on the path. To extend the limit
'   of MAX_PATH wide characters to 32,767 wide characters, prepend "\\?\" to the path.
' Return value:
'   Returns TRUE if the specified file exists or FALSE otherwise.
' Remarks:
'   Prepending the string "\\?\" does not allow access to the root directory.
'   On network shares, you can use a pwszFileSpec in the form of the following:
'   "\\server\service\*". However, you cannot use a pwszFileSpec that points to the share
'   itself; for example, "\\server\service" is not valid.
'   To examine a directory that is not a root directory, use the path to that directory,
'   without a trailing backslash. For example, an argument of "C:\Windows" returns information
'   about the directory "C:\Windows", not about a directory or file in "C:\Windows".
'   To examine the files and directories in "C:\Windows", use an pwszFileSpec of "C:\Windows\*".
'   Be aware that some other thread or process could create or delete a file with this name
'   between the time you query for the result and the time you act on the information.
'   If this is a potential concern for your application, one possible solution is to use
'   the CreateFile function with CREATE_NEW (which fails if the file exists) or OPEN_EXISTING
'   (which fails if the file does not exist).
' ========================================================================================
PRIVATE FUNCTION AfxFileExists (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
   DIM fd AS WIN32_FIND_DATAW
   IF pwszFileSpec = NULL THEN EXIT FUNCTION
   DIM hFind AS HANDLE = FindFirstFileW(pwszFileSpec, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   FindClose hFind
   ' // Make sure that it is not a directory or a temporary file
   IF (fd.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY AND _
      (fd.dwFileAttributes AND FILE_ATTRIBUTE_TEMPORARY) <> FILE_ATTRIBUTE_TEMPORARY THEN
      FUNCTION = TRUE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes the specified file.
' Parameter:
' - pwszFileSpec: The full path and name of the file to delete.
' Return value:
'   If the function succeeds, the return value is nonzero.
'   If the function fails, the return value is zero (0). To get extended error information, call GetLastError.
' Remarks:
'   If an application attempts to delete a file that does not exist, this function fails
'   with ERROR_FILE_NOT_FOUND. If the file is a read-only file, the function fails with
'   ERROR_ACCESS_DENIED.
' ========================================================================================
PRIVATE FUNCTION AfxDeleteFile (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
   IF pwszFileSpec = NULL THEN EXIT FUNCTION
   FUNCTION = DeleteFileW(pwszFileSpec)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Replacemet for Free Basic's Kill. Returns zero (0) on success, or non-zero on error.
' ========================================================================================
PRIVATE FUNCTION AfxKill (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
   IF pwszFileSpec = NULL THEN EXIT FUNCTION
   FUNCTION = NOT DeleteFileW(pwszFileSpec)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new directory.
' - lpPathName : The path of the directory to be created.
'   To extend the limit to 32,767 wide characters, prepend "\?" to the path.
' Return value:
'   If the function succeeds, the return value is TRUE.
'   If the function fails, the return value is FALSE.
'   To get extended error information, call GetLastError.
'   Possible errors include the following.
'   - ERROR_ALREADY_EXISTS
'     The specified directory already exists.
'   - ERROR_PATH_NOT_FOUND
'     One or more intermediate directories do not exist; this function will only create
'     the final directory in the path.
' ========================================================================================
PRIVATE FUNCTION AfxCreateDirectory (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
   RETURN CreateDirectoryW(lpPathName, NULL)
END FUNCTION
' ========================================================================================
PRIVATE FUNCTION AfxMakeDir (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
   RETURN CreateDirectoryW(lpPathName, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Unicode replacement for FreeBasic MkDir.
' Returns 0 on success, or -1 on failure.
' ========================================================================================
PRIVATE FUNCTION AfxMkDir (BYVAL lpPathName AS LPCWSTR) AS LONG
   DIM res AS LONG = CreateDirectoryW(lpPathName, NULL)
   IF res = 0 THEN RETURN -1 ELSE RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes an existing empty directory.
' - lpPathName : The path of the directory to be removed. This path must specify an empty
'   directory, and the calling process must have delete access to the directory.
' To extend the MAX_PATH limit to 32,767 wide characters prepend "\?" to the path.
' If the function succeeds, the return value is TRUE.
' If the function fails, the return value is FALSE.
' To get extended error information, call GetLastError.
' The AfxRemoveDirectory function marks a directory for deletion on close. Therefore, the
' directory is not removed until the last handle to the directory is closed.
' To recursively delete the files in a directory, use the SHFileOperation function.
' AfxRemoveDirectory removes a directory junction, even if the contents of the target are not
' empty; the function removes directory junctions regardless of the state of the target object.
' ========================================================================================
PRIVATE FUNCTION AfxRemoveDirectory (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
   RETURN RemoveDirectoryW(lpPathName)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxRemoveDir (BYVAL lpPathName AS LPCWSTR) AS BOOLEAN
   RETURN RemoveDirectoryW(lpPathName)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Unicode replacement for Free Basic's RmDir.
' Returns 0 on success and -1 on failure.
' ========================================================================================
PRIVATE FUNCTION AfxRmDir (BYVAL lpPathName AS LPCWSTR) AS LONG
   DIM res AS LONG = RemoveDirectoryW(lpPathName)
   IF res = 0 THEN RETURN -1 ELSE RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies an existing file to a new file.
' - lpExistingFileName : The name of an existing file.
'   To extend the limit of MAX_PATH characters to 32,767 wide characters prepend "\?" to the path.
'   If lpExistingFileName does not exist, CopyFile fails, and GetLastError returns ERROR_FILE_NOT_FOUND.
' - lpNewFileName : The name of the new file.
'   To extend the limit of MAX_PATH characters to 32,767 wide characters prepend "\?" to the path.
' - bFailIfExists
'   If this parameter is TRUE and the new file specified by lpNewFileName already exists, the
'   function fails. If this parameter is FALSE and the new file already exists, the function
'   overwrites the existing file and succeeds.
' Return value:
'   If the function succeeds, the return value is TRUE.
'   If the function fails, the return value is FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxCopyFile (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR, BYVAL bFailIfExists AS BOOLEAN = FALSE) AS BOOLEAN
   RETURN CopyFileW(lpExistingFileName, lpNewFileName, bFailIfExists)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Unicode replacement for Free Basic's FileCopy.
' Returns 0 on success, or 1 if an error occurred.
' ========================================================================================
PRIVATE FUNCTION AfxFileCopy (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR) AS LONG
   DIM nRet AS LONG = IIF(CopyFileW(lpExistingFileName, lpNewFileName, FALSE) = 0, 1, 0)
   RETURN nRet
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves an existing file or a directory, including its children.
' - lpExistingFileName : The name of an existing file.
'   To extend the limit of MAX_PATH characters to 32,767 wide characters prepend "\?" to the path.
'   If lpExistingFileName does not exist, AfxMoveFile fails, and GetLastError returns ERROR_FILE_NOT_FOUND.
' - lpNewFileName : The name of the new file.
'   To extend the limit of MAX_PATH characters to 32,767 wide characters prepend "\?" to the path.
' Return value:
'   If the function succeeds, the return value is TRUE.
'   If the function fails, the return value is FALSE. To get extended error information, call GetLastError.
' ========================================================================================
PRIVATE FUNCTION AfxMoveFile (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR) AS BOOLEAN
   RETURN MoveFileW(lpExistingFileName, lpNewFileName)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxRenameFile (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR) AS BOOLEAN
   RETURN MoveFileW(lpExistingFileName, lpNewFileName)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Unicode replacement for Free Basic's Name function.
' Returns 0 on success, or non-zero on failure.
' ========================================================================================
PRIVATE FUNCTION AfxName (BYVAL lpExistingFileName AS LPCWSTR, BYVAL lpNewFileName AS LPCWSTR) AS LONG
   DIM res AS LONG = MoveFileW(lpExistingFileName, lpNewFileName)
   IF res = 0 THEN RETURN -1 ELSE RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Searches a directory for a file or subdirectory with a name that matches a specific name
' (or partial name if wildcards are used).
' Parameter:
' - pwszFileSpec: The directory or path, and the file name, which can include wildcard
'   characters, for example, an asterisk (*) or a question mark (?).
'   This parameter should not be NULL, an invalid string (for example, an empty string or a
'   string that is missing the terminating null character), or end in a trailing backslash (\).
'   If the string ends with a wildcard, period (.), or directory name, the user must have
'   access permissions to the root and all subdirectories on the path. To extend the limit
'   of MAX_PATH wide characters to 32,767 wide characters, prepend "\\?\" to the path.
' Return value:
'   Returns TRUE if the specified folder exists or FALSE otherwise.
' Remarks
'   Prepending the string "\\?\" does not allow access to the root directory.
'   On network shares, you can use a file spec in the form of the following: "\\server\service\*".
'   However, you cannot use a file spec that points to the share itself; for example,
'   "\\server\service" is not valid.
'   To examine a directory that is not a root directory, use the path to that directory,
'   without a trailing backslash. For example, an argument of "C:\Windows" returns information
'   about the directory "C:\Windows", not about a directory or file in "C:\Windows".
'   To examine the files and directories in "C:\Windows", use a file spec of "C:\Windows\*".
'   Be aware that some other thread or process could create or delete a file with this name
'   between the time you query for the result and the time you act on the information.
'   If this is a potential concern for your application, one possible solution is to use
'   the CreateFile function with CREATE_NEW (which fails if the file exists) or OPEN_EXISTING
'   (which fails if the file does not exist).
' ========================================================================================
PRIVATE FUNCTION AfxFolderExists (BYVAL pwszFileSpec AS WSTRING PTR) AS BOOLEAN
   DIM fd AS WIN32_FIND_DATAW
   IF pwszFileSpec = NULL THEN EXIT FUNCTION
   DIM hFind AS HANDLE = FindFirstFileW(pwszFileSpec, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   FindClose hFind
   ' // Make sure that it is a directory
   IF (fd.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY THEN FUNCTION = TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a directory; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsFolder (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a system file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsSystemFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a hidden file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsHiddenFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a read only file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsReadOnlyFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a temporary file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsTemporaryFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a normal file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsNormalFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a compressed file or diectory; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsCompressedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is an encrypted file or diectory; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsEncryptedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_ENCRYPTED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified file or directory is not to be indexed by the content
' indexing service; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsNotContentIndexedFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_NOT_CONTENT_INDEXED) = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified file is not available immediately; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsOffLineFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_OFFLINE) = FILE_ATTRIBUTE_OFFLINE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a file or directory that has an associated reparse
' point, or a file that is a symbolic link; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsReparsePointFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_REPARSE_POINT) = FILE_ATTRIBUTE_REPARSE_POINT)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the specified path is a sparse file; FALSE, otherwise.
' To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' ========================================================================================
PRIVATE FUNCTION AfxIsSparseFile (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM dwAttributes AS DWORD = GetFileAttributesW(wszFileSpec)
   IF dwAttributes = INVALID_FILE_ATTRIBUTES THEN RETURN FALSE
   RETURN ((dwAttributes AND FILE_ATTRIBUTE_SPARSE_FILE) = FILE_ATTRIBUTE_SPARSE_FILE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Unicode replacement for Free Basic's FileDateTime.
' Returns the file's last modified date and time as Date Serial.
' wszFileName : Filename to retrieve date and time for.
' Return Value : Returns a Date Serial.
' Example:
' #include "windows.bi"
' #include "vbcompat.bi"
' #include "Afx/AfxWin.bi"
' DIM wszFileName AS WSTRING * MAX_PATH = ExePath & "\c2.bas"
' DIM dt AS DOUBLE = AfxFileDateTime(wszFileName)
' PRINT Format(dt, "yyyy-mm-dd hh:mm AM/PM")
' ========================================================================================
PRIVATE FUNCTION AfxFileDateTime (BYREF wszFileName AS WSTRING) AS DOUBLE
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileName, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN RETURN 0
   FindClose hFind
   DIM ft AS FILETIME
   FileTimeToLocalFileTime(@fd.ftLastWriteTime, @ft)
   DIM st AS SYSTEMTIME
   FileTimeToSystemTime(@ft, @st)
   DIM dt AS DOUBLE
   SystemTimeToVariantTime @st, @dt
   RETURN dt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Unicode replacement for Free Basic's FileLen.
' Returns the size in bytes of the specified file. Don't use it with folders.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileSize (BYREF wszFileSpec AS WSTRING) AS ULONGLONG
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind = INVALID_HANDLE_VALUE THEN RETURN 0
   FindClose hFind
   IF (fd.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY AND _
      (fd.dwFileAttributes AND FILE_ATTRIBUTE_TEMPORARY) <> FILE_ATTRIBUTE_TEMPORARY THEN
      DIM ullSize AS ULONGLONG = (fd.nFileSizeHigh * (&hFFFFFFFFull + 1)) + fd.nFileSizeLow
      RETURN ullSize
   END IF
END FUNCTION
' ========================================================================================
#define AfxFileLen AfxGetFileSize

' ========================================================================================
' Returns the time the file was created.
' - wszFileSpec: The directory or path, and the file name, which can include wildcard characters,
'   for example, an asterisk (*) or a question mark (?).
'   This parameter should not be NULL, an invalid string (for example, an empty string or a
'   string that is missing the terminating null character), or end in a trailing backslash (\).
'   If the string ends with a wildcard, period (.), or directory name, the user must have access
'   permissions to the root and all subdirectories on the path.
'   To extend the limit from MAX_PATH to 32,767 wide characters, prepend "\\?\" to the path.
' - bUTC: Pass FALSE if you want to get the time in local time (the NTFS file system stores time
'   values in UTC format, so they are not affected by changes in time zone or daylight saving time).
'   FileTimeToLocalFileTime uses the current settings for the time zone and daylight saving time.
'   Therefore, if it is daylight saving time, it takes daylight saving time into account, even
'   if the file time you are converting is in standard time.
' Usage: AfxFileTimeToDateStr(AfxGetFileCreationTime("C:\Tests\test.bas", FALSE), "dd/MM/yyyy")
' ========================================================================================
PRIVATE FUNCTION AfxGetFileCreationTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind <> INVALID_HANDLE_VALUE THEN
      FindClose hFind
      IF bUTC = TRUE THEN
         RETURN fd.ftCreationTime
      ELSE
         DIM FT AS FILETIME
         FileTimeToLocalFileTime(@fd.ftCreationTime, @FT)
         RETURN FT
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the time the file was last accessed. The NTFS file system delays updates to the
' last access time for a file by up to 1 hour after the last access.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileLastAccessTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind <> INVALID_HANDLE_VALUE THEN
      FindClose hFind
      IF bUTC = TRUE THEN
         RETURN fd.ftLastAccessTime
      ELSE
         DIM FT AS FILETIME
         FileTimeToLocalFileTime(@fd.ftLastAccessTime, @FT)
         RETURN FT
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the time the file was last written to, truncated, or overwritten.
' When writing to a file, the last write time is not fully updated until all handles that
' are used for writing are closed.
' ========================================================================================
PRIVATE FUNCTION AfxGetFileLastWriteTime (BYREF wszFileSpec AS WSTRING, BYVAL bUTC AS BOOLEAN = TRUE) AS FILETIME
   DIM fd AS WIN32_FIND_DATAW
   DIM hFind AS HANDLE = FindFirstFileW(wszFileSpec, @fd)
   IF hFind <> INVALID_HANDLE_VALUE THEN
      FindClose hFind
      IF bUTC = TRUE THEN
         RETURN fd.ftLastWriteTime
      ELSE
         DIM FT AS FILETIME
         FileTimeToLocalFileTime(@fd.ftLastWriteTime, @FT)
         RETURN FT
      END IF
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a file is likely to contain a form of unicode text.
' ========================================================================================
PRIVATE FUNCTION AfxIsFileUnicode (BYREF wszFileName AS WSTRING) AS BOOLEAN
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN RETURN FALSE
   DIM dwFileSize AS DWORD, dwHighSize AS DWORD
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   DIM dwBytesToRead AS DWORD = 1024
   IF dwFileSize < dwBytesToRead THEN dwBytesToRead = dwFileSize
   DIM pBuffer AS UBYTE PTR = CAllocate(1, dwBytesToRead)
   DIM bRes AS BOOLEAN, dwBytesRead AS DWORD
   bRes = ReadFile(hFile, pBuffer, dwBytesToRead, @dwBytesRead, NULL)
   IF bRes THEN bRes = IsTextUnicode(pBuffer, dwBytesRead, NULL)
   IF pBuffer THEN DeAllocate(pBuffer)
   IF hFile THEN CloseHandle(hFile)
   RETURN bRes
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scans a text file ans returns the number of occurrences of the specified delimiter.
' Default value is CHR(13, 10), which returns the number of lines.
' ========================================================================================
PRIVATE FUNCTION AfxFileScanA (BYREF wszFileName AS WSTRING, BYREF szDelimiter AS ZSTRING = CHR(13, 10)) AS DWORD
   DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   IF LEN(szDelimiter) = 0 THEN EXIT FUNCTION
   ' // Open the file
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   ' // Get the size of the file
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   DIM pBuffer AS UBYTE PTR
   pBuffer = CAllocate(1, dwFileSize)
   IF pBuffer = NULL THEN EXIT FUNCTION
   DIM bSuccess AS LONG = ReadFile(hFile, pBuffer, dwFileSize, @dwBytesRead, NULL)
   CloseHandle(hFile)
   IF bSuccess = FALSE THEN EXIT FUNCTION
   DIM nLen AS LONG = LEN(szDelimiter)
   DIM pstr AS ANY PTR = pBuffer
   DO
      pstr = strstr(pstr, szDelimiter)
      IF pstr = NULL THEN EXIT DO
      pstr += nLen
      dwCount += 1
   LOOP
   DeAllocate(pBuffer)
   FUNCTION = dwCount
END FUNCTION
' ========================================================================================
' ========================================================================================
' Version for unicode text files.
' ========================================================================================
PRIVATE FUNCTION AfxFileScanW (BYREF wszFileName AS WSTRING, BYREF wszDelimiter AS WSTRING = CHR(13, 10)) AS DWORD
   DIM dwCount AS DWORD, dwFileSize AS DWORD, dwHighSize AS DWORD, dwBytesRead AS DWORD
   IF LEN(wszDelimiter) = 0 THEN EXIT FUNCTION
   ' // Open the file
   DIM hFile AS HANDLE = CreateFileW(@wszFileName, GENERIC_READ, FILE_SHARE_READ, NULL, _
                         OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   ' // Get the size of the file
   dwFileSize = GetFileSize(hFile, @dwHighSize)
   DIM pBuffer AS UBYTE PTR
   pBuffer = CAllocate(1, dwFileSize)
   IF pBuffer = NULL THEN EXIT FUNCTION
   DIM bSuccess AS LONG = ReadFile(hFile, pBuffer, dwFileSize, @dwBytesRead, NULL)
   CloseHandle(hFile)
   IF bSuccess = FALSE THEN EXIT FUNCTION
   DIM nLen AS LONG = LEN(wszDelimiter) * 2
   DIM pstr AS ANY PTR = pBuffer
   DO
      pstr = wcsstr(pstr, @wszDelimiter)
      IF pstr = NULL THEN EXIT DO
      pstr += nLen
      dwCount += 1
   LOOP
   DeAllocate(pBuffer)
   FUNCTION = dwCount
END FUNCTION
' ========================================================================================

END NAMESPACE
