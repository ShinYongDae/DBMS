// DataFile.cpp: implementation of the CDataFile class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DataFile.h"
#include "Common.h"

#pragma warning(disable : 4996)

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
CDataFile::CDataFile()
{
	m_str = "";
	m_bModified = FALSE;
	m_bOpened = FALSE;
	m_FilePath = "";
}

CDataFile::CDataFile(LPCTSTR lpszFileName)
{
	m_str = "";
	m_bModified = FALSE;
	m_bOpened = FALSE;
	m_FilePath = "";
	if(!Open(lpszFileName))
	{
		CFile file;
		CFileException pError;
		if(!file.Open(lpszFileName,CFile::modeCreate|CFile::modeReadWrite,&pError))
		{
			// 파일 오픈에 실패시 
			#ifdef _DEBUG
			   afxDump << "File could not be opened " << pError.m_cause << "\n";
			#endif
		}
		else
		{
			m_bModified = FALSE;
			m_bOpened = TRUE;
			// Open(lpszFileName);
		}
	}
}

CDataFile::~CDataFile()
{
	m_str.Empty();
//	if(m_bModified)
	//	Close();
}

BOOL CDataFile::Open(LPCTSTR lpszFileName)
{
	if(!m_FilePath.Compare(lpszFileName)) return TRUE; // 이미 열려 있는 경우 TRUE를 리턴한다.
	
	CFile file;
	CFileException pError;
	if(!file.Open(lpszFileName,CFile::modeRead,&pError))
	{
		// 파일 오픈에 실패시 
		#ifdef _DEBUG
		   afxDump << "File could not be opened " << pError.m_cause << "\n";
		#endif
		m_bOpened = FALSE;
		return FALSE;
	}
	else
	{
		m_str.Empty();
		file.SeekToBegin();
		DWORD dwFileLength = file.GetLength();

		// 파일 크기 만큼 버퍼 메모리를 확보한다.
		TCHAR *cpTemp;
		cpTemp = (TCHAR *)malloc(sizeof(TCHAR) * dwFileLength+1);
		
		//file의 내용을 버퍼로 읽어 들인다.
		UINT nBytesRead = file.Read(cpTemp,dwFileLength);
		
		file.Close();
		
		// 버퍼의 끝처리를 한다.
		*(cpTemp+dwFileLength) = NULL;
		// 버퍼의 내용을 CString으로 옮긴다.
		
		LPTSTR p;
		if(nBytesRead >0)
		{
			//m_str.GetBuffer(nBytesRead+1);
			p = m_str.GetBuffer(nBytesRead+1);
			_tcscpy(p,cpTemp);
			m_str.ReleaseBuffer();  // Surplus memory released, p is now invalid.
			ASSERT( m_str.GetLength() == (int)nBytesRead ); // Length still nBytesRead
		}
		
		// 버퍼 메모리를 릴리즈 한다.
		free(cpTemp);

		m_FilePath = lpszFileName;

		m_bModified = FALSE;
		m_bOpened = TRUE;
	}
	return TRUE;
}

BOOL CDataFile::Close()
{
	if(m_bOpened && m_bModified) // 파일이 수정되었을 경우 수정된 내용를 원본 파일에 반영한다.
	{
		CFile file;
		CFileException pError;
		if(!file.Open(m_FilePath,CFile::modeWrite,&pError))
		{
			// 파일 오픈에 실패시 
			#ifdef _DEBUG
			   afxDump << "File could not be opened " << pError.m_cause << "\n";
			#endif
			return FALSE;
		}
		else
		{	
			//버퍼의 내용을 file에 복사한다.
			file.SeekToBegin();
			file.Write(m_str,m_str.GetLength());
			file.Close();
			m_bModified = FALSE;
		}
	}
	
	return TRUE;
}

////////////////////////////////////////////////////////
// 파일의 크기(Size)를 리턴한다.
// 파일이 열려 있지 않으면 0를 리턴한다.
LONG CDataFile::GetSize()
{
	if(!m_bOpened) return 0;
	return (LONG)m_str.GetLength();
}

LONG CDataFile::Find(TCHAR ch)
{
	return (LONG)m_str.Find(ch);
}
LONG CDataFile::Find(LPCTSTR lpszSub)
{
	return (LONG)m_str.Find(lpszSub);
}
LONG CDataFile::Find(TCHAR ch, int nStart)
{
	return (LONG)m_str.Find(ch,nStart);
}
LONG CDataFile::Find(LPCTSTR pstr, int nStart)
{
	return (LONG)m_str.Find(pstr,nStart);
}

//////////////////////////////////////////////////////
// 파일의 끝에 문자열을 붙인다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않으면 0를 리턴한다.
LONG CDataFile::Append(LPCTSTR str)
{
	if(!m_bOpened) return 0;
	m_str.Insert(m_str.GetLength(),str);
	m_bModified = TRUE;
	return GetSize();
}

//////////////////////////////////////////////////////
// 파일의 시작을 기준으로 nIndex의 위치에 str의 문자열을 삽입한다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않으면 0를 리턴한다.
LONG CDataFile::Insert(int nIndex,LPCTSTR str)
{
	if(!m_bOpened) return 0;
	m_str.Insert(nIndex,str);
	m_bModified = TRUE;
	return GetSize();
}
//////////////////////////////////////////////////////
// 파일의 시작을 기준으로 nIndex의 위치로 부터 nCount의 문자를 삭제한다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않으면 0를 리턴한다.
LONG CDataFile::Delete(int nIndex, int nCount)
{
	if(!m_bOpened) return 0;
	m_str.Delete(nIndex,nCount);
	m_bModified = TRUE;
	return GetSize();
}

////////////////////////////////////////////////////////
// 파일에서 처음부터 lpszOld로 지정된 문자열을 모두 찾아 lpszNew로 바꾼다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않으면 0를 리턴한다.
LONG CDataFile::Replace(LPCTSTR lpszOld, LPCTSTR lpszNew)
{
	if(!m_bOpened) return 0;
	int nCount = m_str.Replace(lpszOld,lpszNew);
	m_bModified = TRUE;
	return GetSize();
}
////////////////////////////////////////////////////////
// 파일의 라인수를 리턴한다.
// 파일이 열려 있지 않으면 0를 리턴한다.
UINT CDataFile::GetTotalLines()
{
	if(!m_bOpened) return 0;

	UINT nLine = 0;
	int nCurPos,nPrevPos = 0;
	// '\n'을 이용하여 라인의 수를 읽는다.
	do{
		nCurPos = m_str.Find('\n',nPrevPos);
		if(nCurPos == -1) break;
		nLine++;
		nPrevPos = (nCurPos+1);
	}while(1);
	
	// 마지막 '\n'의 위치와 파일의 끝이 다르다면 라인수를 하나 증가시킨다.  
	nCurPos = m_str.GetLength();
	if(nPrevPos != m_str.GetLength())
		nLine++;
	return nLine;
}

////////////////////////////////////////////////////////
// 파일에서 nStart위치로 부터 str을 찾아 str의 라인 위치를 리턴한다.
// 파일이 열려 있지 않거나, 해당 스트링이 없으면 -1을 리턴한다.
int CDataFile::GetLineNumber(LPCTSTR str, int nStart)
{
	if(!m_bOpened) return -1;

	CString strLine;
	UINT nMaxLine = GetTotalLines();
	for(UINT nLine = nStart; nLine <= nMaxLine; nLine++)
	{
		// 라인의 스트링을 읽어 들인다.
		strLine = GetLineString(nLine);
		// str이 존재 하는 지를 조사 한다.
		if(strLine.Find(str,0) != -1)
			return nLine;
	}
	return -1;
}

////////////////////////////////////////////////////////
// 파일에서 nStart위치로 부터 Descriptor를 찾아 라인 위치를 리턴한다.
// 파일이 열려 있지 않거나, 해당 스트링이 없으면 -1을 리턴한다.
int CDataFile::GetLineNumOfDescriptor(LPCTSTR Descriptor, int nStart)
{
	if(!m_bOpened) return -1;

	CString strLine,strValue;
	UINT nMaxLine = GetTotalLines();
	for(UINT nLine = nStart; nLine <= nMaxLine; nLine++)
	{
		// 라인의 스트링을 읽어 들인다.
		strLine = GetLineString(nLine);
		// Descriptor부분을 추출한다.
		strValue = ExtractFirstWordcom(&strLine,_T("="),TRUE);
		// 찾고자 하는 Descriptor가 존재 하는 지를 조사 한다.
		if(!strValue.Compare(Descriptor))
			return nLine;
	}
	return -1;
}

////////////////////////////////////////////////////////
// 파일에서 nLine으로 지정된 라인의 시작 위치를 리턴한다.
// 파일이 열려 있지 않으면 -1을 리턴한다.
int CDataFile::GetLineStartPosition(UINT nLine)
{
	if(!m_bOpened) return -1;

	if(nLine == 1) return 0;

	int nCurPos,nPrevPos = 0;
	UINT  nCurLine = 1;
	// '\n'을 이용하여 라인의 수를 읽는다.
	do{
		nCurPos = m_str.Find('\n',nPrevPos); // 라인의 끝을 찾는다.
		if(nCurPos == -1) break;
		nCurLine++;
		nPrevPos = (nCurPos+1);
		char ch = m_str.GetAt(nCurPos+1);
		CString str = m_str.Mid(nCurPos+1,5);
		if(nCurLine == nLine) break;
	}while(1);
	
	// 마지막 '\n'의 위치와 파일의 끝이 다르다면 라인수를 하나 증가시킨다.  
	return nPrevPos;
}

////////////////////////////////////////////////////////
// 파일에서 nStartIndex위치에서 nEndIndex의 위치까지의 문자열을 반환한다.
// 이때 반환되는 문자열은 NULL로 종료되는 문자열이다.
// 파일이 열려 있지 않으면 NULL을 리턴한다.
CString CDataFile::GetString(int nStartIndex, int nEndIndex)
{
	if(!m_bOpened) return CString("");

	CString str;
	str = m_str.Mid(nStartIndex,nEndIndex-nStartIndex);
	return str;
}
////////////////////////////////////////////////////////
// 파일에서 nLine위치의 문자열을 반환한다.
// 파일이 열려 있지 않거나, nLine이 잘못 지정되면 NULL을 리턴한다.
CString CDataFile::GetLineString(UINT nLine)
{
	if(!m_bOpened) return CString("");

	// 파일의 라인수를 읽는다.
	UINT nMaxLine = GetTotalLines();
	
	// 읽고자 하는 라인번호가 적당한지를 조사한다.
	if (nLine > nMaxLine || nLine < 1)
	{
		AfxMessageBox(_T("Line string read error"));
		return CString("");
	}

	int nStartIndex,nEndIndex;
	
	// 라인의 시작 위치를 찾는다.
	nStartIndex = GetLineStartPosition(nLine);

	// 다음 라인의 시작 위치를 찾는다.
	if (nLine < nMaxLine)
		nEndIndex = GetLineStartPosition(nLine+1);
	else
		nEndIndex = m_str.GetLength();

	// 라인의 스트링을 읽어 들인다.
	CString strLine;
	strLine = GetString(nStartIndex,nEndIndex);

	// 라인에서 라인피드 문자가 있을 경우 자동 삭제 된다.
	strLine.Remove('\r');
	// 라인에서 뉴라인 문자가 있을 경우 자동 삭제 된다.
	strLine.Remove('\n');
	
	return strLine;
}

////////////////////////////////////////////////////////
// 파일에서 nLine위치의 문자열을 반환한다.
// 파일이 열려 있지 않거나, nLine이 잘못 지정되면 NULL을 리턴한다.
BOOL CDataFile::GetLineString(UINT nLine,CString& rtnStr)
{
	if(!m_bOpened) return FALSE;

	// 파일의 라인수를 읽는다.
	UINT nMaxLine = GetTotalLines();
	
	// 읽고자 하는 라인번호가 적당한지를 조사한다.
	if (nLine > nMaxLine || nLine < 1) return FALSE;

	int nStartIndex,nEndIndex;
	
	// 라인의 시작 위치를 찾는다.
	nStartIndex = GetLineStartPosition(nLine);

	// 다음 라인의 시작 위치를 찾는다.
	if (nLine < nMaxLine)
		nEndIndex = GetLineStartPosition(nLine+1);
	else
		nEndIndex = m_str.GetLength();

	// 라인의 스트링을 읽어 들인다.
	rtnStr = GetString(nStartIndex,nEndIndex);

	// 라인에서 라인피드 문자가 있을 경우 자동 삭제 된다.
	rtnStr.Remove('\r');
	// 라인에서 뉴라인 문자가 있을 경우 자동 삭제 된다.
	rtnStr.Remove('\n');
	

	return TRUE;
}
////////////////////////////////////////////////////////
// 파일에서 nStart라인부터 nEnd라인 까지의 문자열을 반환한다.
// 이때 반환되는 문자열은 NULL로 종료되는 문자열이다.
// 파일이 열려 있지 않거나, nLine이 잘못 지정되면 NULL을 리턴한다.
CString CDataFile::GetLineString(int nStart, int nEnd)
{
	if(!m_bOpened) return CString("");
	
	// 파일의 라인수를 읽는다.
	int nMaxLine = GetTotalLines();
	
	// 읽고자 하는 라인번호가 적당한지를 조사한다.
	if (nStart < 1 || nEnd > nMaxLine || nStart > nEnd)
	{
		AfxMessageBox(_T("Multi line string read error"));
		return CString("");
	}

	if(nStart == nEnd)
		return GetLineString(nStart);

	int nStartIndex,nEndIndex;
	
	// 라인의 시작 위치를 찾는다.
	nStartIndex = GetLineStartPosition(nStart);

	// 다음 라인의 시작 위치를 찾는다.
	if (nEnd < nMaxLine)
		nEndIndex = GetLineStartPosition(nEnd+1);
	else
		nEndIndex = m_str.GetLength();

	// 라인(Block)의 스트링을 읽어 들인다.
	CString strLine;
	strLine = GetString(nStartIndex,nEndIndex);


	// 스트링의 끝에 라인피드 문자가 있을 경우 자동 삭제 된다.
	int n = strLine.ReverseFind('\r');
	if (n > 0)
		strLine.Delete(n,1);

	// 스트링의 끝에 뉴라인 문자가 있을 경우 자동 삭제 된다..
	n = strLine.ReverseFind('\n');
	if (n > 0)
		strLine.Delete(n,1);
	
	return strLine;
}

////////////////////////////////////////////////////////
// 파일의 끝에 str로 지정된 한라인의 문자열을 추가 한다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않으면 -1을 리턴한다.
LONG CDataFile::AppendLine(LPCTSTR str)
{
	if(!m_bOpened) return -1;

	// 파일내용이 없다면.
	if(!m_str.GetLength()) 
		return Append(str);
	
	// 스트링 앞에 뉴라인 문자를 삽입하여 파일에 추가 시킨다.
	CString strTemp;
	strTemp = _T("\r\n") + (CString)str;
	Append(strTemp);

	return GetSize();
}

////////////////////////////////////////////////////////
// 파일의 지정된 라인앞에 문자열을 삽입한다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않으면 NULL을 리턴한다.
LONG CDataFile::InsertLine(UINT nLine, LPCTSTR strInsert)
{
	if(!m_bOpened) return -1;
	// 파일의 내용이 없다면,
	if(!m_str.GetLength())
		return Append(strInsert);

	// 파일의 라인수를 읽는다.
	UINT nMaxLine = GetTotalLines();
	
	// 삽입하고자 하는 라인번호가 적당한지를 조사한다.
	if (nLine < 1)
	{
		AfxMessageBox(_T("Line string insert error"));
		return -1;
	}
	CString str;
	if(nLine > nMaxLine)
	{
		str = _T("\r\n") + (CString)strInsert;
		Append(str);
	}
	else
	{
		// 삽입하고자하는 스트링에 라인피드 및 뉴라인 문자가 없을 경우 자동 추가 된다.	
		str = strInsert;
		int n = str.Find(_T("\r\n"),0);
		if (n < 0)
			str += _T("\r\n");
		
		int nIndex;
		// 라인의 시작위치를 찾는다.
		nIndex = GetLineStartPosition(nLine);

		Insert(nIndex,str);
	}
	return GetSize();
}
////////////////////////////////////////////////////////
// 파일에서 nLine의 문자열을 삭제 한다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않거나, nLine이 잘못 지정되면 -1을 리턴한다.
LONG CDataFile::DeleteLine(UINT nLine)
{
	if(!m_bOpened) return -1;
	
	if(!m_str.GetLength())
		return -1;

	// 파일의 라인수를 읽는다.
	int nMaxLine = GetTotalLines();
	
	// 지우고자 하는 라인번호가 적당한지를 조사한다.
	if (nLine < 1)
	{
		AfxMessageBox(_T("Line string Delete error"));
		return -1;
	}
	
	int nStartIndex,nEndIndex;
	// 라인의 시작위치를 찾는다.
	nStartIndex = GetLineStartPosition(nLine);
	// 라인의 종료 위치를 찾는다.
	nEndIndex = m_str.Find('\n',nStartIndex)+1;

	Delete(nStartIndex,nEndIndex-nStartIndex);
	return GetSize();
}
////////////////////////////////////////////////////////
// 파일에서 nLine의 문자열을 lpszNew의 문자열로 변경한다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않거나, nLine이 잘못 지정되면 -1을 리턴한다.
LONG CDataFile::ReplaceLine(UINT nLine,LPCTSTR lpszNew)
{
	if(!m_bOpened) return -1;
	// 파일의 라인수를 읽는다.
	UINT nMaxLine = GetTotalLines();
	// 변경 하고자 하는 라인번호가 적당한지를 조사한다.
	if (nLine < 1 || nLine > nMaxLine)
	{
		AfxMessageBox(_T("Line string Replace Error"));
		return -1;
	}
	
	DeleteLine(nLine);
	InsertLine(nLine,lpszNew);
	return GetSize();
}

////////////////////////////////////////////////////////
// 파일에서 nLine의 문자열중 lpszOld부분을 lpszNew의 문자열로 변경한다.
// 변경된 파일의 크기를 LONG단위로 리턴한다.
// 파일이 열려 있지 않거나, nLine이 잘못 지정되면 -1을 리턴한다.
LONG CDataFile::ReplaceLineString(UINT nLine,LPCTSTR lpszOld,LPCTSTR lpszNew)
{
	if(!m_bOpened) return -1;
	// 파일의 라인수를 읽는다.
	UINT nMaxLine = GetTotalLines();
	// 변경 하고자 하는 라인번호가 적당한지를 조사한다.
	if (nLine < 1 || nLine > nMaxLine)
	{
		AfxMessageBox(_T("Line string Replace Error"));
		return -1;
	}
	CString strLine = GetLineString(nLine);
	strLine.Replace(lpszOld,lpszNew);

	DeleteLine(nLine);
	InsertLine(nLine,strLine);
	return GetSize();
}

// nSrcLine의 라인을, int nDestLine으로 옮긴다.
BOOL CDataFile::MoveLine(int nSrcLine, int nDestLine)
{
	if(!m_bOpened) return FALSE;
	// 파일의 라인수를 읽는다.
	int nMaxLine = GetTotalLines();
	
	// 이동 하고자 하는 라인번호가 적당한지를 조사한다.
	if (nSrcLine < 1 || nDestLine < 1 || nSrcLine > nMaxLine || nDestLine > nMaxLine)
	{
		AfxMessageBox(_T("Line string move error"));
		return FALSE;
	}

	// 이동 할 라인의 시작 위치를 찾는다.
	int nStartIndex;
	nStartIndex = GetLineStartPosition(nSrcLine);
	// 이동 할 라인의 종료 위치를 찾는다.
	int nEndIndex;
	if(nSrcLine == nMaxLine)
		nEndIndex = m_str.GetLength();
	else
		nEndIndex = GetLineStartPosition(nSrcLine+1);

	// 이동 할 라인의 스트링("\r\n"포함)을 읽어 들인다.
	CString strLine;
	strLine = GetString(nStartIndex,nEndIndex);		
	// 이동 할 라인을 지운다.
	Delete(nStartIndex,nEndIndex-nStartIndex);
	// 이동 시킬 라인의 시작 위치를 찾는다.
	if(nSrcLine < nDestLine)
	{
		if(nDestLine==nMaxLine)
		{
			// 이동 시킬 라인이 마지막 라인 이라면 파일의 끝에 라인피드및 뉴라인 문자를 추가
			nStartIndex = m_str.GetLength();
			// 파일의 끝이 라인피드및 뉴라인 문자로 끝나지 않고 있으면,
			if(m_str.Find(_T("\r\n"),m_str.GetLength()-2) < 0)
			{
				strLine = _T("\r\n") + strLine;
				strLine.Delete(strLine.GetLength()-2,2);
			}

		}
		else
			nStartIndex = GetLineStartPosition(nDestLine-1);
	}
	else
	{
		nStartIndex = GetLineStartPosition(nDestLine);

		// 이동 라인이 마지막 라인이고 라인피드및 뉴라인 문자가 없다면
		// 이동라인의 마지막에 라인피드및 뉴라인 문자를 추가 시키고, 
		// 파일의 끝에는 라인피드및 뉴라인 문자를 지운다.
		if(nSrcLine == nMaxLine)
		{
			if(strLine.Find(_T("\r\n"),0) < 0)
			{
				strLine += _T("\r\n");
				Delete(m_str.GetLength()-2,2);
			}
		}
	}

	// 이동 시킬 위치에 라인을 삽입한다.
	Insert(nStartIndex,strLine);
	return TRUE;
}

BOOL CDataFile::ReadData(CString strDataGroup, CString strDescription,int nTh,double &fData)
{
	CString strValue, strLine;
	int nLine, nStartRow;

	// 파일에서 데이타 그룹의 시작 라인 번호를 읽어 온다.
	nStartRow = GetLineNumber(strDataGroup,1);
	if(nStartRow < 0)
	{
		strLine.Format(_T("None Exist Group Name in %s file"),m_FilePath);
		AfxMessageBox(strLine,MB_ICONSTOP|MB_OK);
		return FALSE;
	}
	nStartRow++;
	// 데이터 기술자(Decriptor)를 기준으로 데이터의 라인 번호를 얻는다.
	nLine = GetLineNumOfDescriptor(strDescription,nStartRow);
	if(nLine < 0)
	{
		strLine.Format(_T("None Exist Data Descriptor in %s file"),m_FilePath);
		AfxMessageBox(strLine,MB_ICONSTOP|MB_OK);
		return FALSE;
	}
	// 파일로 부터 데이타 문자열을 읽는다.
	strLine = GetLineString(nLine);
	// 데이터 라인으로부터 주석문을 제거 한다.
	strValue = ExtractFirstWordcom(&strLine, _T("//"),TRUE); // Row Index 의 String
	// 데이타 라인으로브터 기술자를 제거한다.
	ExtractFirstWordcom(&strValue, _T("=")); // Row Index 의 String

	if(nTh <= 0)
	{
		AfxMessageBox(_T("Data Index Muse be more than 0"),MB_ICONSTOP|MB_OK);
		return FALSE;
	}
	
	// 데이타 n번째 이전 데이터를 제거한다.
	for(int i = 0; i < nTh-1; i++)
		ExtractFirstWordcom(&strValue, _T(","));

	fData = _tcstod(strValue, NULL);

	return TRUE;
}
