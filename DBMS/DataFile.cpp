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
			// ���� ���¿� ���н� 
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
	if(!m_FilePath.Compare(lpszFileName)) return TRUE; // �̹� ���� �ִ� ��� TRUE�� �����Ѵ�.
	
	CFile file;
	CFileException pError;
	if(!file.Open(lpszFileName,CFile::modeRead,&pError))
	{
		// ���� ���¿� ���н� 
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

		// ���� ũ�� ��ŭ ���� �޸𸮸� Ȯ���Ѵ�.
		TCHAR *cpTemp;
		cpTemp = (TCHAR *)malloc(sizeof(TCHAR) * dwFileLength+1);
		
		//file�� ������ ���۷� �о� ���δ�.
		UINT nBytesRead = file.Read(cpTemp,dwFileLength);
		
		file.Close();
		
		// ������ ��ó���� �Ѵ�.
		*(cpTemp+dwFileLength) = NULL;
		// ������ ������ CString���� �ű��.
		
		LPTSTR p;
		if(nBytesRead >0)
		{
			//m_str.GetBuffer(nBytesRead+1);
			p = m_str.GetBuffer(nBytesRead+1);
			_tcscpy(p,cpTemp);
			m_str.ReleaseBuffer();  // Surplus memory released, p is now invalid.
			ASSERT( m_str.GetLength() == (int)nBytesRead ); // Length still nBytesRead
		}
		
		// ���� �޸𸮸� ������ �Ѵ�.
		free(cpTemp);

		m_FilePath = lpszFileName;

		m_bModified = FALSE;
		m_bOpened = TRUE;
	}
	return TRUE;
}

BOOL CDataFile::Close()
{
	if(m_bOpened && m_bModified) // ������ �����Ǿ��� ��� ������ ���븦 ���� ���Ͽ� �ݿ��Ѵ�.
	{
		CFile file;
		CFileException pError;
		if(!file.Open(m_FilePath,CFile::modeWrite,&pError))
		{
			// ���� ���¿� ���н� 
			#ifdef _DEBUG
			   afxDump << "File could not be opened " << pError.m_cause << "\n";
			#endif
			return FALSE;
		}
		else
		{	
			//������ ������ file�� �����Ѵ�.
			file.SeekToBegin();
			file.Write(m_str,m_str.GetLength());
			file.Close();
			m_bModified = FALSE;
		}
	}
	
	return TRUE;
}

////////////////////////////////////////////////////////
// ������ ũ��(Size)�� �����Ѵ�.
// ������ ���� ���� ������ 0�� �����Ѵ�.
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
// ������ ���� ���ڿ��� ���δ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� ������ 0�� �����Ѵ�.
LONG CDataFile::Append(LPCTSTR str)
{
	if(!m_bOpened) return 0;
	m_str.Insert(m_str.GetLength(),str);
	m_bModified = TRUE;
	return GetSize();
}

//////////////////////////////////////////////////////
// ������ ������ �������� nIndex�� ��ġ�� str�� ���ڿ��� �����Ѵ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� ������ 0�� �����Ѵ�.
LONG CDataFile::Insert(int nIndex,LPCTSTR str)
{
	if(!m_bOpened) return 0;
	m_str.Insert(nIndex,str);
	m_bModified = TRUE;
	return GetSize();
}
//////////////////////////////////////////////////////
// ������ ������ �������� nIndex�� ��ġ�� ���� nCount�� ���ڸ� �����Ѵ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� ������ 0�� �����Ѵ�.
LONG CDataFile::Delete(int nIndex, int nCount)
{
	if(!m_bOpened) return 0;
	m_str.Delete(nIndex,nCount);
	m_bModified = TRUE;
	return GetSize();
}

////////////////////////////////////////////////////////
// ���Ͽ��� ó������ lpszOld�� ������ ���ڿ��� ��� ã�� lpszNew�� �ٲ۴�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� ������ 0�� �����Ѵ�.
LONG CDataFile::Replace(LPCTSTR lpszOld, LPCTSTR lpszNew)
{
	if(!m_bOpened) return 0;
	int nCount = m_str.Replace(lpszOld,lpszNew);
	m_bModified = TRUE;
	return GetSize();
}
////////////////////////////////////////////////////////
// ������ ���μ��� �����Ѵ�.
// ������ ���� ���� ������ 0�� �����Ѵ�.
UINT CDataFile::GetTotalLines()
{
	if(!m_bOpened) return 0;

	UINT nLine = 0;
	int nCurPos,nPrevPos = 0;
	// '\n'�� �̿��Ͽ� ������ ���� �д´�.
	do{
		nCurPos = m_str.Find('\n',nPrevPos);
		if(nCurPos == -1) break;
		nLine++;
		nPrevPos = (nCurPos+1);
	}while(1);
	
	// ������ '\n'�� ��ġ�� ������ ���� �ٸ��ٸ� ���μ��� �ϳ� ������Ų��.  
	nCurPos = m_str.GetLength();
	if(nPrevPos != m_str.GetLength())
		nLine++;
	return nLine;
}

////////////////////////////////////////////////////////
// ���Ͽ��� nStart��ġ�� ���� str�� ã�� str�� ���� ��ġ�� �����Ѵ�.
// ������ ���� ���� �ʰų�, �ش� ��Ʈ���� ������ -1�� �����Ѵ�.
int CDataFile::GetLineNumber(LPCTSTR str, int nStart)
{
	if(!m_bOpened) return -1;

	CString strLine;
	UINT nMaxLine = GetTotalLines();
	for(UINT nLine = nStart; nLine <= nMaxLine; nLine++)
	{
		// ������ ��Ʈ���� �о� ���δ�.
		strLine = GetLineString(nLine);
		// str�� ���� �ϴ� ���� ���� �Ѵ�.
		if(strLine.Find(str,0) != -1)
			return nLine;
	}
	return -1;
}

////////////////////////////////////////////////////////
// ���Ͽ��� nStart��ġ�� ���� Descriptor�� ã�� ���� ��ġ�� �����Ѵ�.
// ������ ���� ���� �ʰų�, �ش� ��Ʈ���� ������ -1�� �����Ѵ�.
int CDataFile::GetLineNumOfDescriptor(LPCTSTR Descriptor, int nStart)
{
	if(!m_bOpened) return -1;

	CString strLine,strValue;
	UINT nMaxLine = GetTotalLines();
	for(UINT nLine = nStart; nLine <= nMaxLine; nLine++)
	{
		// ������ ��Ʈ���� �о� ���δ�.
		strLine = GetLineString(nLine);
		// Descriptor�κ��� �����Ѵ�.
		strValue = ExtractFirstWordcom(&strLine,_T("="),TRUE);
		// ã���� �ϴ� Descriptor�� ���� �ϴ� ���� ���� �Ѵ�.
		if(!strValue.Compare(Descriptor))
			return nLine;
	}
	return -1;
}

////////////////////////////////////////////////////////
// ���Ͽ��� nLine���� ������ ������ ���� ��ġ�� �����Ѵ�.
// ������ ���� ���� ������ -1�� �����Ѵ�.
int CDataFile::GetLineStartPosition(UINT nLine)
{
	if(!m_bOpened) return -1;

	if(nLine == 1) return 0;

	int nCurPos,nPrevPos = 0;
	UINT  nCurLine = 1;
	// '\n'�� �̿��Ͽ� ������ ���� �д´�.
	do{
		nCurPos = m_str.Find('\n',nPrevPos); // ������ ���� ã�´�.
		if(nCurPos == -1) break;
		nCurLine++;
		nPrevPos = (nCurPos+1);
		char ch = m_str.GetAt(nCurPos+1);
		CString str = m_str.Mid(nCurPos+1,5);
		if(nCurLine == nLine) break;
	}while(1);
	
	// ������ '\n'�� ��ġ�� ������ ���� �ٸ��ٸ� ���μ��� �ϳ� ������Ų��.  
	return nPrevPos;
}

////////////////////////////////////////////////////////
// ���Ͽ��� nStartIndex��ġ���� nEndIndex�� ��ġ������ ���ڿ��� ��ȯ�Ѵ�.
// �̶� ��ȯ�Ǵ� ���ڿ��� NULL�� ����Ǵ� ���ڿ��̴�.
// ������ ���� ���� ������ NULL�� �����Ѵ�.
CString CDataFile::GetString(int nStartIndex, int nEndIndex)
{
	if(!m_bOpened) return CString("");

	CString str;
	str = m_str.Mid(nStartIndex,nEndIndex-nStartIndex);
	return str;
}
////////////////////////////////////////////////////////
// ���Ͽ��� nLine��ġ�� ���ڿ��� ��ȯ�Ѵ�.
// ������ ���� ���� �ʰų�, nLine�� �߸� �����Ǹ� NULL�� �����Ѵ�.
CString CDataFile::GetLineString(UINT nLine)
{
	if(!m_bOpened) return CString("");

	// ������ ���μ��� �д´�.
	UINT nMaxLine = GetTotalLines();
	
	// �а��� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
	if (nLine > nMaxLine || nLine < 1)
	{
		AfxMessageBox(_T("Line string read error"));
		return CString("");
	}

	int nStartIndex,nEndIndex;
	
	// ������ ���� ��ġ�� ã�´�.
	nStartIndex = GetLineStartPosition(nLine);

	// ���� ������ ���� ��ġ�� ã�´�.
	if (nLine < nMaxLine)
		nEndIndex = GetLineStartPosition(nLine+1);
	else
		nEndIndex = m_str.GetLength();

	// ������ ��Ʈ���� �о� ���δ�.
	CString strLine;
	strLine = GetString(nStartIndex,nEndIndex);

	// ���ο��� �����ǵ� ���ڰ� ���� ��� �ڵ� ���� �ȴ�.
	strLine.Remove('\r');
	// ���ο��� ������ ���ڰ� ���� ��� �ڵ� ���� �ȴ�.
	strLine.Remove('\n');
	
	return strLine;
}

////////////////////////////////////////////////////////
// ���Ͽ��� nLine��ġ�� ���ڿ��� ��ȯ�Ѵ�.
// ������ ���� ���� �ʰų�, nLine�� �߸� �����Ǹ� NULL�� �����Ѵ�.
BOOL CDataFile::GetLineString(UINT nLine,CString& rtnStr)
{
	if(!m_bOpened) return FALSE;

	// ������ ���μ��� �д´�.
	UINT nMaxLine = GetTotalLines();
	
	// �а��� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
	if (nLine > nMaxLine || nLine < 1) return FALSE;

	int nStartIndex,nEndIndex;
	
	// ������ ���� ��ġ�� ã�´�.
	nStartIndex = GetLineStartPosition(nLine);

	// ���� ������ ���� ��ġ�� ã�´�.
	if (nLine < nMaxLine)
		nEndIndex = GetLineStartPosition(nLine+1);
	else
		nEndIndex = m_str.GetLength();

	// ������ ��Ʈ���� �о� ���δ�.
	rtnStr = GetString(nStartIndex,nEndIndex);

	// ���ο��� �����ǵ� ���ڰ� ���� ��� �ڵ� ���� �ȴ�.
	rtnStr.Remove('\r');
	// ���ο��� ������ ���ڰ� ���� ��� �ڵ� ���� �ȴ�.
	rtnStr.Remove('\n');
	

	return TRUE;
}
////////////////////////////////////////////////////////
// ���Ͽ��� nStart���κ��� nEnd���� ������ ���ڿ��� ��ȯ�Ѵ�.
// �̶� ��ȯ�Ǵ� ���ڿ��� NULL�� ����Ǵ� ���ڿ��̴�.
// ������ ���� ���� �ʰų�, nLine�� �߸� �����Ǹ� NULL�� �����Ѵ�.
CString CDataFile::GetLineString(int nStart, int nEnd)
{
	if(!m_bOpened) return CString("");
	
	// ������ ���μ��� �д´�.
	int nMaxLine = GetTotalLines();
	
	// �а��� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
	if (nStart < 1 || nEnd > nMaxLine || nStart > nEnd)
	{
		AfxMessageBox(_T("Multi line string read error"));
		return CString("");
	}

	if(nStart == nEnd)
		return GetLineString(nStart);

	int nStartIndex,nEndIndex;
	
	// ������ ���� ��ġ�� ã�´�.
	nStartIndex = GetLineStartPosition(nStart);

	// ���� ������ ���� ��ġ�� ã�´�.
	if (nEnd < nMaxLine)
		nEndIndex = GetLineStartPosition(nEnd+1);
	else
		nEndIndex = m_str.GetLength();

	// ����(Block)�� ��Ʈ���� �о� ���δ�.
	CString strLine;
	strLine = GetString(nStartIndex,nEndIndex);


	// ��Ʈ���� ���� �����ǵ� ���ڰ� ���� ��� �ڵ� ���� �ȴ�.
	int n = strLine.ReverseFind('\r');
	if (n > 0)
		strLine.Delete(n,1);

	// ��Ʈ���� ���� ������ ���ڰ� ���� ��� �ڵ� ���� �ȴ�..
	n = strLine.ReverseFind('\n');
	if (n > 0)
		strLine.Delete(n,1);
	
	return strLine;
}

////////////////////////////////////////////////////////
// ������ ���� str�� ������ �Ѷ����� ���ڿ��� �߰� �Ѵ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� ������ -1�� �����Ѵ�.
LONG CDataFile::AppendLine(LPCTSTR str)
{
	if(!m_bOpened) return -1;

	// ���ϳ����� ���ٸ�.
	if(!m_str.GetLength()) 
		return Append(str);
	
	// ��Ʈ�� �տ� ������ ���ڸ� �����Ͽ� ���Ͽ� �߰� ��Ų��.
	CString strTemp;
	strTemp = _T("\r\n") + (CString)str;
	Append(strTemp);

	return GetSize();
}

////////////////////////////////////////////////////////
// ������ ������ ���ξտ� ���ڿ��� �����Ѵ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� ������ NULL�� �����Ѵ�.
LONG CDataFile::InsertLine(UINT nLine, LPCTSTR strInsert)
{
	if(!m_bOpened) return -1;
	// ������ ������ ���ٸ�,
	if(!m_str.GetLength())
		return Append(strInsert);

	// ������ ���μ��� �д´�.
	UINT nMaxLine = GetTotalLines();
	
	// �����ϰ��� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
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
		// �����ϰ����ϴ� ��Ʈ���� �����ǵ� �� ������ ���ڰ� ���� ��� �ڵ� �߰� �ȴ�.	
		str = strInsert;
		int n = str.Find(_T("\r\n"),0);
		if (n < 0)
			str += _T("\r\n");
		
		int nIndex;
		// ������ ������ġ�� ã�´�.
		nIndex = GetLineStartPosition(nLine);

		Insert(nIndex,str);
	}
	return GetSize();
}
////////////////////////////////////////////////////////
// ���Ͽ��� nLine�� ���ڿ��� ���� �Ѵ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� �ʰų�, nLine�� �߸� �����Ǹ� -1�� �����Ѵ�.
LONG CDataFile::DeleteLine(UINT nLine)
{
	if(!m_bOpened) return -1;
	
	if(!m_str.GetLength())
		return -1;

	// ������ ���μ��� �д´�.
	int nMaxLine = GetTotalLines();
	
	// ������� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
	if (nLine < 1)
	{
		AfxMessageBox(_T("Line string Delete error"));
		return -1;
	}
	
	int nStartIndex,nEndIndex;
	// ������ ������ġ�� ã�´�.
	nStartIndex = GetLineStartPosition(nLine);
	// ������ ���� ��ġ�� ã�´�.
	nEndIndex = m_str.Find('\n',nStartIndex)+1;

	Delete(nStartIndex,nEndIndex-nStartIndex);
	return GetSize();
}
////////////////////////////////////////////////////////
// ���Ͽ��� nLine�� ���ڿ��� lpszNew�� ���ڿ��� �����Ѵ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� �ʰų�, nLine�� �߸� �����Ǹ� -1�� �����Ѵ�.
LONG CDataFile::ReplaceLine(UINT nLine,LPCTSTR lpszNew)
{
	if(!m_bOpened) return -1;
	// ������ ���μ��� �д´�.
	UINT nMaxLine = GetTotalLines();
	// ���� �ϰ��� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
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
// ���Ͽ��� nLine�� ���ڿ��� lpszOld�κ��� lpszNew�� ���ڿ��� �����Ѵ�.
// ����� ������ ũ�⸦ LONG������ �����Ѵ�.
// ������ ���� ���� �ʰų�, nLine�� �߸� �����Ǹ� -1�� �����Ѵ�.
LONG CDataFile::ReplaceLineString(UINT nLine,LPCTSTR lpszOld,LPCTSTR lpszNew)
{
	if(!m_bOpened) return -1;
	// ������ ���μ��� �д´�.
	UINT nMaxLine = GetTotalLines();
	// ���� �ϰ��� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
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

// nSrcLine�� ������, int nDestLine���� �ű��.
BOOL CDataFile::MoveLine(int nSrcLine, int nDestLine)
{
	if(!m_bOpened) return FALSE;
	// ������ ���μ��� �д´�.
	int nMaxLine = GetTotalLines();
	
	// �̵� �ϰ��� �ϴ� ���ι�ȣ�� ���������� �����Ѵ�.
	if (nSrcLine < 1 || nDestLine < 1 || nSrcLine > nMaxLine || nDestLine > nMaxLine)
	{
		AfxMessageBox(_T("Line string move error"));
		return FALSE;
	}

	// �̵� �� ������ ���� ��ġ�� ã�´�.
	int nStartIndex;
	nStartIndex = GetLineStartPosition(nSrcLine);
	// �̵� �� ������ ���� ��ġ�� ã�´�.
	int nEndIndex;
	if(nSrcLine == nMaxLine)
		nEndIndex = m_str.GetLength();
	else
		nEndIndex = GetLineStartPosition(nSrcLine+1);

	// �̵� �� ������ ��Ʈ��("\r\n"����)�� �о� ���δ�.
	CString strLine;
	strLine = GetString(nStartIndex,nEndIndex);		
	// �̵� �� ������ �����.
	Delete(nStartIndex,nEndIndex-nStartIndex);
	// �̵� ��ų ������ ���� ��ġ�� ã�´�.
	if(nSrcLine < nDestLine)
	{
		if(nDestLine==nMaxLine)
		{
			// �̵� ��ų ������ ������ ���� �̶�� ������ ���� �����ǵ�� ������ ���ڸ� �߰�
			nStartIndex = m_str.GetLength();
			// ������ ���� �����ǵ�� ������ ���ڷ� ������ �ʰ� ������,
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

		// �̵� ������ ������ �����̰� �����ǵ�� ������ ���ڰ� ���ٸ�
		// �̵������� �������� �����ǵ�� ������ ���ڸ� �߰� ��Ű��, 
		// ������ ������ �����ǵ�� ������ ���ڸ� �����.
		if(nSrcLine == nMaxLine)
		{
			if(strLine.Find(_T("\r\n"),0) < 0)
			{
				strLine += _T("\r\n");
				Delete(m_str.GetLength()-2,2);
			}
		}
	}

	// �̵� ��ų ��ġ�� ������ �����Ѵ�.
	Insert(nStartIndex,strLine);
	return TRUE;
}

BOOL CDataFile::ReadData(CString strDataGroup, CString strDescription,int nTh,double &fData)
{
	CString strValue, strLine;
	int nLine, nStartRow;

	// ���Ͽ��� ����Ÿ �׷��� ���� ���� ��ȣ�� �о� �´�.
	nStartRow = GetLineNumber(strDataGroup,1);
	if(nStartRow < 0)
	{
		strLine.Format(_T("None Exist Group Name in %s file"),m_FilePath);
		AfxMessageBox(strLine,MB_ICONSTOP|MB_OK);
		return FALSE;
	}
	nStartRow++;
	// ������ �����(Decriptor)�� �������� �������� ���� ��ȣ�� ��´�.
	nLine = GetLineNumOfDescriptor(strDescription,nStartRow);
	if(nLine < 0)
	{
		strLine.Format(_T("None Exist Data Descriptor in %s file"),m_FilePath);
		AfxMessageBox(strLine,MB_ICONSTOP|MB_OK);
		return FALSE;
	}
	// ���Ϸ� ���� ����Ÿ ���ڿ��� �д´�.
	strLine = GetLineString(nLine);
	// ������ �������κ��� �ּ����� ���� �Ѵ�.
	strValue = ExtractFirstWordcom(&strLine, _T("//"),TRUE); // Row Index �� String
	// ����Ÿ �������κ��� ����ڸ� �����Ѵ�.
	ExtractFirstWordcom(&strValue, _T("=")); // Row Index �� String

	if(nTh <= 0)
	{
		AfxMessageBox(_T("Data Index Muse be more than 0"),MB_ICONSTOP|MB_OK);
		return FALSE;
	}
	
	// ����Ÿ n��° ���� �����͸� �����Ѵ�.
	for(int i = 0; i < nTh-1; i++)
		ExtractFirstWordcom(&strValue, _T(","));

	fData = _tcstod(strValue, NULL);

	return TRUE;
}
