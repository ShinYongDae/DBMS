#pragma once
class CTestQuery
{
public:
	CTestQuery();
	~CTestQuery();

	CDataSource m_dataSource;

	BOOL RIGHT(CString strQuery);
	BOOL GetTables(CString strQuery, CStringArray &strTableList);
	BOOL GetList(CString strQuery, CStringArray &strList1, CStringArray &strList2);

	void InitDB(LPCTSTR szServerIP, LPCTSTR szCatalog, LPCTSTR szUserID, LPCTSTR szUserPassword);
	void InitDB(LPCTSTR szServerIP, LPCTSTR szCatalog, LPCTSTR szUserID, LPCTSTR szUserPassword, int nDbtype);
};

