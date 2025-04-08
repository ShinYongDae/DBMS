#include "stdafx.h"
#include "TestQuery.h"


CTestQuery::CTestQuery()
{
}


CTestQuery::~CTestQuery()
{
}

BOOL CTestQuery::GetTables(CString strQuery, CStringArray & strTableList)
{
	long lMaxCols, lMaxRows, lRow;

	if (!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindCustomerCode()\r\n%s"), m_dataSource.GetLastError());
		return FALSE;
	}

	lMaxCols = m_dataSource.m_pRS->Fields->Count;
	lMaxRows = m_dataSource.m_pRS->RecordCount;

	if (lMaxRows > 0)
	{
		_variant_t vColName, vValue;
		CString strData;
		m_dataSource.m_pRS->MoveFirst();

		for (lRow = 0; lRow < lMaxRows; lRow++)
		{
			// fetch equipment name
			vValue = m_dataSource.m_pRS->Fields->Item[(long)0]->Value;
			if (vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal;
				strTableList.Add(strData);
			}
			m_dataSource.m_pRS->MoveNext();
		}
	}

	return TRUE;
}
BOOL CTestQuery::GetList(CString strQuery, CStringArray & strList1, CStringArray & strList2)
{
	long lMaxCols, lMaxRows, lRow;

	
	if (!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindCustomerCode()\r\n%s"), m_dataSource.GetLastError());
		return FALSE;
	}

	lMaxCols = m_dataSource.m_pRS->Fields->Count;
	lMaxRows = m_dataSource.m_pRS->RecordCount;
	BOOL bCheck;
	if (lMaxRows > 0)
	{
		_variant_t vColName, vValue;
		CString strData;
		m_dataSource.m_pRS->MoveFirst();

		for (lRow = 0; lRow < lMaxRows; lRow++)
		{
			// fetch equipment name
			bCheck = FALSE;
			vValue = m_dataSource.m_pRS->Fields->Item[(long)0]->Value;
			if (vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal;
				strList1.Add(strData);
				bCheck = TRUE;
			}
			if (bCheck)
			{
				vValue = m_dataSource.m_pRS->Fields->Item[(long)1]->Value;
				if (vValue.vt != VT_NULL)
				{
					vValue.ChangeType(VT_BSTR);
					strData = vValue.bstrVal;
					strList2.Add(strData);					
				}
				else
				{
					strData.IsEmpty();
					strList2.Add(strData);
				}
			}
			m_dataSource.m_pRS->MoveNext();
		}
	}

	return TRUE;
}

BOOL CTestQuery::RIGHT(CString strQuery)
{
	if (!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindCustomerCode()\r\n%s"), m_dataSource.GetLastError());
		return FALSE;
	}
	
	return TRUE;
}

void CTestQuery::InitDB(LPCTSTR szServerIP, LPCTSTR szCatalog, LPCTSTR szUserID, LPCTSTR szUserPassword)
{
	m_dataSource.InitDB(szServerIP, szCatalog, szUserID, szUserPassword);
}

void CTestQuery::InitDB(LPCTSTR szServerIP, LPCTSTR szCatalog, LPCTSTR szUserID, LPCTSTR szUserPassword, int nDbType)
{
	m_dataSource.InitDB(szServerIP, szCatalog, szUserID, szUserPassword, nDbType);
}