#include "stdafx.h"
#include "DBMS.h"
#include "common.h"

#pragma warning(disable : 4996)
#pragma comment(lib,"shlwapi")
#include <shlwapi.h>
#include <tchar.h>


//////////////////////////////////////////////////////////////////
// Function name    : ExtractFirstWord
// Description      : 입력된 문자열중 첫 워드를 제거한다.
// Return type      : CString  입력된 문자열중 제거된 첫 워드
// Argument         : CString *str, CString sep,BOOL bTrim
// 
// 1. str : Source String
// 2. sep : Delimiter
// 3. bTrim : 좌우 Space의 제거 Option


CString ExtractFirstWordcom(CString *str, CString sep,BOOL bTrim)
{
	CString csRtn;
	int nDelPos = str->Find(sep,0);
	if(nDelPos < 0)
	{
		csRtn = *str;
		*str = "";
	}
	else
	{
		csRtn = str->Left(nDelPos);
		*str = str->Right(str->GetLength()-(csRtn.GetLength()+sep.GetLength()));
	}
	
	if(bTrim)
	{
		csRtn.TrimLeft(VK_SPACE); csRtn.TrimLeft(VK_TAB);
		csRtn.TrimRight(VK_SPACE); csRtn.TrimRight(VK_TAB);
	}

	return csRtn;
}

CString ExtractFirstWordcom(CString& str, CString sep,BOOL bTrim)
{
	CString csRtn;
	int nDelPos = str.Find(sep,0);
	if(nDelPos < 0)
	{
		csRtn = str;
		str = "";
	}
	else
	{
		csRtn = str.Left(nDelPos);
		str = str.Right(str.GetLength()-(csRtn.GetLength()+sep.GetLength()));
	}
	
	if(bTrim)
	{
		csRtn.TrimLeft(VK_SPACE); csRtn.TrimLeft(VK_TAB);
		csRtn.TrimRight(VK_SPACE); csRtn.TrimRight(VK_TAB);
	}
	return csRtn;
}

//////////////////////////////////////////////////////////////////
// Function name    : ExtractLastWord
// Description      : 입력된 문자열중 마지막 워드를 제거한다.
// Return type      : CString  입력된 문자열중 제거된 마지막 워드
// Argument         : CString *str, CString sep,BOOL bTrim
// 
// 1. str : Source String
// 2. sep : Delimiter
// 3. bTrim : 좌우 Space의 제거 Option
CString ExtractLastWordcom(CString *str, char sep,BOOL bTrim)
{
	CString csRtn;
	int nDelPos = str->ReverseFind(sep);
	if(nDelPos < 0)
	{
		csRtn = *str;
		*str = "";
	}
	else
	{
		csRtn = str->Right(str->GetLength()-nDelPos-1);
		*str = str->Left(nDelPos);
	}
	
	if(bTrim)
	{
		csRtn.TrimLeft(VK_SPACE); csRtn.TrimLeft(VK_TAB);
		csRtn.TrimRight(VK_SPACE); csRtn.TrimRight(VK_TAB);
	}

	return csRtn;
}

CString ExtractLastWordcom(CString& str, char sep,BOOL bTrim)
{
	CString csRtn;
	int nDelPos = str.ReverseFind(sep);
	if(nDelPos < 0)
	{
		csRtn = str;
		str = "";
	}
	else
	{
		csRtn = str.Right(str.GetLength()-nDelPos-1);
		str = str.Left(nDelPos);
	}
	
	if(bTrim)
	{
		csRtn.TrimLeft(VK_SPACE); csRtn.TrimLeft(VK_TAB);
		csRtn.TrimRight(VK_SPACE); csRtn.TrimRight(VK_TAB);
	}
	return csRtn;
}


CCriticalSection g_log;
//#include "safelock.h"
void Log(CString strMsg)
{
	/*
	CSafeLock autolock(&g_log);
	TCHAR szFile[MAX_PATH]={0,};
	TCHAR szPath[MAX_PATH]={0,};
	TCHAR* pszPos= NULL;
	GetModuleFileName(NULL, szPath, MAX_PATH);
	pszPos = _tcsrchr(szPath, '\\');
	
	*pszPos = NULL;

	_stprintf( szPath, _T("C:\\GenaLog"));
	CreateDirectory( szPath, NULL);

//	sprintf( szPath, "%s\\%s", szPath,m_pszExeName);
//	CreateDirectory( szPath, NULL );

	_stprintf( szFile, _T("%s\\%s.txt"), szPath, COleDateTime::GetCurrentTime().Format( _T("%Y%m%d") ) );

	CString strDate;
	CString strContents;
	
	strDate.Format(_T("%s: "),COleDateTime::GetCurrentTime().Format( _T("%Y/%m/%d %H:%M:%S"))); 
	strContents = strDate; 
	strContents += strMsg;
	strContents += "\r\n";
	CFile file;
	if( file.Open( szFile, CFile::modeCreate| CFile::modeNoTruncate | CFile::modeWrite |  CFile::shareDenyNone ) == 0)
	return;   

	file.SeekToEnd();
	file.Write( strContents , strContents.GetLength() );   
	file.Flush();
	file.Close(); 
	*/
}

BOOL MakeDirectory(CString strPath)
{
	CString strMsg;
	int i=0;
	int nLen = strPath.GetLength();

	char ch;
	CString strTemp;
	strTemp = strPath.Left(2);
	if(strTemp.CompareNoCase(_T("\\\\"))==0)
	{
		//remote computer name
		i = strPath.Find(_T("\\"),2)+1;
		strTemp = strPath.Left(i);
	}
	else
		strTemp="";


    for (; i < nLen ; i++ )
    {
        ch = strPath.GetAt(i);
        strTemp += ch;
		if ( ch == _T('\\') || ch == _T('/') )
        {
            if (!PathFileExists(strTemp))
            {
                if (!::CreateDirectory(strTemp,NULL))
                {
                    if (GetLastError() != ERROR_ALREADY_EXISTS )
                        return FALSE;
                }
            }
        }
    }

    if (!PathFileExists(strTemp))
    {
        if (!::CreateDirectory(strTemp,NULL))
        {
            if (GetLastError() != ERROR_ALREADY_EXISTS )
                return FALSE;
        }
    }

	return TRUE;
}

BOOL DeleteDirectory(CString strDelDir)
{
	SHFILEOPSTRUCT shfo;

	TCHAR pszFrom[1024] = {0};
	TCHAR pszTo[1024] = {0};

	WORD wFlag = 0;

	_tcscpy(pszFrom, strDelDir);
	_tcscpy(pszTo, strDelDir);

	pszFrom[_tcslen(pszFrom) + 1] = 0;

	shfo.hwnd = NULL;
	shfo.wFunc = FO_DELETE;
	shfo.pFrom = pszFrom;
	shfo.pTo = pszTo;
	shfo.fFlags = FOF_NOCONFIRMATION;
	shfo.lpszProgressTitle = _T("Delete directory..");

	return SHFileOperation(&shfo); 
}

BOOL DeleteDirectory2(CString strDelDir)
{
	CString strPath;

	strPath = strDelDir;

	if(strPath.Right(1) != "\\")
		strPath += "\\";
	strPath += "*.*";
	
	CFileFind ff;
	BOOL bFind = ff.FindFile(strPath);
	if(!bFind)
		return FALSE;
	else
	{
		while(bFind) // until we find an entry
		{
			bFind = ff.FindNextFile();
			if(ff.IsDots())
				continue;

			if(ff.IsDirectory())
			{
				strPath = ff.GetFilePath();
				// we have found a directory, recurse
				if(!DeleteDirectory(strPath))
				{
					ff.Close();
					return FALSE;    // directory couldn't be deleted
				}
				// remove the empty directory
				RemoveDirectory(strPath);
			}
			else
			{
				CString strTemp;
				strTemp=ff.GetFilePath();
				if(ff.IsReadOnly())
				{
					CFileStatus fStatus;
					CFile::GetStatus(strTemp,fStatus);
					fStatus.m_attribute = 0;
					CFile::SetStatus(strTemp,fStatus);
					
				}
				if(!DeleteFile(strTemp))
				{
					ff.Close();
					return FALSE;
				}
			}
		}
		ff.Close();
	}
	return TRUE;
}

BOOL RenameDirectory(CString strFrom,CString strTo)
{
	SHFILEOPSTRUCT shfo;

	TCHAR pszFrom[1024] = {0};
	TCHAR pszTo[1024] = {0};

	WORD wFlag = 0;

	_tcscpy(pszFrom,strFrom);
	_tcscpy(pszTo, strTo);

	pszFrom[_tcslen(pszFrom) + 1] = 0;

	shfo.hwnd = NULL;
	shfo.wFunc = FO_RENAME;
	shfo.pFrom = pszFrom;
	shfo.pTo = pszTo;
	shfo.fFlags = FOF_NOCONFIRMATION;
	shfo.lpszProgressTitle = _T("File Rename..");

	return SHFileOperation(&shfo); 
}

LONGLONG GetFolderCapacity(CString& strFolder)
{
	LONGLONG lFileSize=0;
	
	CFileFind finder;

	CString sWildCard = strFolder + "\\*.*";

    BOOL bFound;
    BOOL bWorking = finder.FindFile(sWildCard);

    bFound = bWorking;

    while(bWorking)
    {
        bWorking = finder.FindNextFile();

        if(finder.IsDots())
			continue;

        // if dirrectory then recur call
        if(finder.IsDirectory())
        {
            CString strFolder2 = finder.GetFilePath();
            lFileSize += GetFolderCapacity(strFolder2);
            continue;
        }
		else
			lFileSize += finder.GetLength();

    }
    return lFileSize;
}

BOOL CopyFolder2(CString &strFrom, CString &strTo)
{
	SHFILEOPSTRUCT shfo;

    TCHAR pszFrom[1024] = {0};
    TCHAR pszTo[1024] = {0};	
	
	memset(pszTo,0,sizeof(pszTo));
	memset(pszFrom,0,sizeof(pszFrom));

	SetFileAttributes(strTo,FILE_ATTRIBUTE_ARCHIVE);
	
	_tcscpy(pszFrom, strFrom);
	_tcscpy(pszTo, strTo);
	pszFrom[_tcslen(pszFrom) + 1] = 0;

	ZeroMemory(&shfo, sizeof shfo);
	shfo.hwnd = NULL;
	shfo.wFunc = FO_COPY;
	shfo.pFrom = pszFrom;
	shfo.pTo = pszTo;
	shfo.fFlags = FOF_NOCONFIRMATION | FOF_MULTIDESTFILES | FOF_NOCONFIRMMKDIR;

	shfo.lpszProgressTitle = _T("File Copy..");

	int nRet = SHFileOperation(&shfo);

	return nRet;
}

BOOL CopyFolder(CString &strFrom, CString &strTo)
{
	CFileFind finder;

	CString sWildCard = strFrom + "\\*.*";

    BOOL bFound;
    BOOL bWorking = finder.FindFile(sWildCard);

    bFound = bWorking;

    if(!MakeDirectory(strTo))
	{
		AfxMessageBox(_T("Save Directory making Failure!"));
		return FALSE;
	}

    while(bWorking)
    {
        bWorking = finder.FindNextFile();

        // skip . and .. files; otherwise, we"d
        // recur infinitely!
        if(finder.IsDots()) continue;

        // 디렉토리이툈E재귀 호횁E
        if(finder.IsDirectory())
        {
            CString sFrom2 = finder.GetFilePath();
            CString sTo2 = strTo + "\\" + finder.GetFileName();
            if(::CreateDirectory(sTo2, NULL) != 0)
				CopyFolder(sFrom2, sTo2);
			else
			{
				int kk = 9;
			}	
            continue;
        }
        ExCopyFile(finder.GetFilePath(), strTo);
    }
    return bFound;
}

int ExCopyFile(CString strFrom, CString strTo)
{
    SHFILEOPSTRUCT shfo;

    TCHAR pszFrom[1024] = {0};
    TCHAR pszTo[1024] = {0};

    WORD wFlag = 0;

	memset(pszTo,0,sizeof(pszTo));
	memset(pszFrom,0,sizeof(pszFrom));
    
	_tcscpy(pszFrom, strFrom);
    _tcscpy(pszTo, strTo);
    pszFrom[_tcslen(pszFrom) + 1] = 0;

	ZeroMemory(&shfo, sizeof shfo);
    shfo.hwnd = AfxGetMainWnd()->m_hWnd; //hWnd;
    shfo.wFunc = FO_COPY;
    shfo.pFrom = pszFrom;
    shfo.pTo = pszTo;
    shfo.fFlags = FOF_NOCONFIRMATION;
    shfo.lpszProgressTitle = _T("File Copy..");

    return SHFileOperation(&shfo);
}

BOOL FindFolder(CString strRootPath,CString strFolder, CString &strRetPath)
{
	CFileFind finder;

	CString sWildCard = strRootPath + "\\*.*";
	CString strPath;

    BOOL bFound = finder.FindFile(sWildCard);

	while(bFound)
    {
        bFound = finder.FindNextFile();

        // skip . and .. files; otherwise, we"d
        // recur infinitely!
        if(finder.IsDots())
			continue;

        if(finder.IsDirectory())
        {
            strPath = finder.GetFilePath();
            CString strFolderName = finder.GetFileName();
			if(strFolderName.CompareNoCase(strFolder)==0)
			{
				strRetPath = strPath;
				return TRUE;
			}
			else
			{
				if(FindFolder(strPath, strFolder,strRetPath))
					return TRUE;
				else
					continue;
			}
        }
    }
    return FALSE;
}

BOOL FindFile(CString strRootPath,CString strFile, CString &strRetPath)
{
	CFileFind finder;

	CString strWildCard = strRootPath + "\\*.*";
	CString strPath;

    BOOL bFound = finder.FindFile(strWildCard);

	while(bFound)
    {
        bFound = finder.FindNextFile();

        // skip . and .. files; otherwise, we"d
        // recur infinitely!
        if(finder.IsDots())
			continue;

        if(!finder.IsDirectory())
        {
            strPath = finder.GetFilePath();
            CString strFileName = finder.GetFileName();
			if(strFileName.CompareNoCase(strFile)==0)
			{
				strRetPath = strPath;
				return TRUE;
			}
			else
			{
				if(FindFile(strPath, strFile,strRetPath))
					return TRUE;
				else
					continue;
			}
        }
    }
    return FALSE;
}

char* StringToChar(CString str)
{
	char*		szStr = NULL;
	wchar_t*	wszStr;
	int				nLenth;

	USES_CONVERSION;
	//1. CString to wchar_t* conversion
	wszStr = T2W(str.GetBuffer(str.GetLength()));

	//2. wchar_t* to char* conversion
	nLenth = WideCharToMultiByte(CP_ACP, 0, wszStr, -1, NULL, 0, NULL, NULL); //char* 형에 대한길이를 구함 
	szStr = new char[nLenth];  //메모리 할당 

							   //3. wchar_t* to char* conversion
	WideCharToMultiByte(CP_ACP, 0, wszStr, -1, szStr, nLenth, 0, 0);
	return szStr;
}
void StringToChar(CString str, char* pCh)
{
	wchar_t*	wszStr;
	int				nLenth;

	USES_CONVERSION;
	//1. CString to wchar_t* conversion
	wszStr = T2W(str.GetBuffer(str.GetLength()));

	//2. wchar_t* to char* conversion
	nLenth = WideCharToMultiByte(CP_ACP, 0, wszStr, -1, NULL, 0, NULL, NULL); //char* 형에 대한길이를 구함 

																			  //3. wchar_t* to char* conversion
	WideCharToMultiByte(CP_ACP, 0, wszStr, -1, pCh, nLenth, 0, 0);
	return;
}
char* StringToChar(CString str, int &nLen)
{
	char*		szStr = NULL;
	wchar_t*	wszStr;

	USES_CONVERSION;
	//1. CString to wchar_t* conversion
	wszStr = T2W(str.GetBuffer(str.GetLength()));

	//2. wchar_t* to char* conversion
	nLen = WideCharToMultiByte(CP_ACP, 0, wszStr, -1, NULL, 0, NULL, NULL); //char* 형에 대한길이를 구함 
	szStr = new char[nLen];  //메모리 할당 

							 //3. wchar_t* to char* conversion
	WideCharToMultiByte(CP_ACP, 0, wszStr, -1, szStr, nLen, 0, 0);
	return szStr;
}