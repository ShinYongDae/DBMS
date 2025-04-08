#if !defined(AFX_SPREADSHEETEX_H__022C83D2_92D7_4253_8E47_17ECFEE50E75__INCLUDED_)
#define AFX_SPREADSHEETEX_H__022C83D2_92D7_4253_8E47_17ECFEE50E75__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SpreadSheetEx.h : header file
//
#include "Spread8/Include/Ssdllmfc.h"
#include "DataSource.h"

/////////////////////////////////////////////////////////////////////////////
// CSpreadSheetEx window
#define MAX_FIELD 256

typedef struct DataField{
	BOOL bExtField;		// External referenced field
	BOOL bPrimary;		// Primary Key
	BOOL bCombo;		// Combo Style Field
	BOOL bColor;		// Color Field
	CString strFieldId;	// Field Name
	CString strAlias;	// Alias name for field display
	CString StrComboData; // 
}sttField;
typedef CArray<sttField, sttField> CArField;

class CSpreadSheetEx : public TSpreadEx
{
// Construction
public:
	CSpreadSheetEx();

// Attributes
public:
	int InitOverEES(CString strQuery,CString strTableName= _T(""),long lActiveRow=1,BOOL bSheetLock=FALSE);
	CString	m_strAccessID;
	CString	m_strPassword;
	CString	m_strCatalog;
	CString	m_strServerIp;
	
	CDataSource m_dataSource;
	
	long m_nMaxCols;
	long m_nMaxRows;
	
	CDWordArray m_arDeletedRecord;
	CDWordArray m_arModifiedRecord;
	CDWordArray m_arAppendedRecord;
	CDWordArray m_arInsertedRecord;

	CArField *pArField; // CArray pointer for Field 
	
	CStringArray m_OrgFieldIdList; // My data table field name
	CStringArray m_AliasFieldNameList; // My data table field name
	CStringArray m_FieldIDList;	// Full field name
	CStringArray m_PrimaryKeyList; //PrimaryKey style field name
	CStringArray m_ComboStyleFieldIdList; // Combo style field name
	CStringArray m_ColorStyleFieldIdList; // Color style field name
	CStringArray m_ComboDataList; // Combo Data List for Combo style field
	CStringArray m_ComboDataQuery;

	BOOL m_bSheetLock;
	CString m_strQuery;
	CString m_strTable;
	COLORREF m_rgbHeadBg, m_rgbHeadFont;
	double m_dHeadFontSize;
	BOOL m_bHeadFontBold, m_bCellFontBold;
	CString m_strFontName;
	double m_dHeadHeight, m_dCellWidth, m_dCellHeight;
	int m_nCellFontSize;
	COLORREF m_rgbCellBg, m_rgbCellFont;
	long m_nCellHAlign,m_nCellVAlign;
	long m_nOldSelectRow, m_nNewSelectRow;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSpreadSheetEx)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CSpreadSheetEx();
	int m_nDBType;
	CString SetTest(CString strQuery);
	void Init(CString strQuery, CString strTableName=_T(""),long lActiveRow=1,BOOL bSheetLock=FALSE);
	void InitDisplayTable(CString strQuery, CString strTableName= _T(""));
	int InitOver(CString strQuery, CString strTableName= _T(""),long lActiveRow=1,BOOL bSheetLock=FALSE);
	int InitOverEquipment(CString strQuery,float *ReturnData1,float *ReturnData2,float *ReturnData3);
	int GetReportMachineInfo_ALL(CString strQuery,CStringArray &EquipmentNames,CString strTableName= _T(""),long lActiveRow=1,BOOL bSheetLock=FALSE);
	int GetReportMachineInfo_Status(CString strQuery,int nEquipmentCnt,CString strEquipCode, CString strTableName= _T(""),long lActiveRow=1,BOOL bSheetLock=FALSE);
	//int GetReportInspectionResult(CString strQuery1,CString strQuery2);
	void FindFieldIdOnDb();	// Find Field name list
	BOOL GetAliasFieldNameOnDb(CString strFieldId,CString &strAliasname);	// Find Field name list
	void FindOrgFieldIdOnDb(); // Find Original Field name list 
	void FindPkOnDb();
	void SetPkOnSpread();
	void FindComboFieldIdOnDb();					// 데이타베이스에서 콤보스타일의 필드명을 찾는다.
	void FindComboDataQueryOnDb();					// 데이타베이스에서 콤보데이타을 얻기위한 쿼리문을 찾는다.
	void FindComboDataOnDb(CString strFieldName);	// 데이타베이스에서 콤보데이타를 찾는다.
	void FindColorFieldIdOnDb();
	void SetComboOnSpread();
	long  GetColorFieldIndex();
	BOOL CheckPrimaryKeyField(CString strFieldId);
	BOOL CheckComboStyleField(CString strFieldId);
	BOOL CheckExtReferedField(CString strFieldId);
	BOOL CheckColorStyleField(CString strFieldId);

	CString GetComboDataList(CString strFieldName);
	int  GetPrimaryKeyData(int nRecordNo,CStringArray &strArray);
	int  GetInvalidPrimaryKeyField(int nRecordNo);
	int  GetInvalidComboStyleField(int nRecordNo);
	void ModifiedRecord(long lIndex,long lField);
	void InsertRecord(long lIndex);
	void DeleteRecord(long lIndex);
	void AppendRecord();
	void SaveRecord(long lIndex=1);
	void SaveEquipment(long lIndex=1);
	void ComboCloseUp(long Col, long Row, short SelChange);
	void ComboCloseUp(long Col, long Row);

	void SetDBServer(CString strServerIp,CString strCatalog,CString strAccessID,CString strPassword);
	void SetSpreadInfo(double dHeadHeight=15.0, COLORREF rgbHeadBg=RGB(226,233,251), COLORREF rgbHeadFont=RGB(0,0,0), double dHeadFontSize=14, BOOL bHeadFontBold=FALSE, double dCellWidth=10.0, double dCellHeight=11.0, int nCellFontSize=9, BOOL bCellFontBold=FALSE, COLORREF rgbCellBg=RGB(255,255,255), COLORREF rgbCellFont=RGB(0,0,0), long nCellHAlign=ALIGN_LEFT, long nCellVAlign=ALIGN_LEFT, CString strFontName=_T("굴림"));
	void DispSpread(BOOL Field=0);
	void DispSpreadTable(BOOL Field=0);
	void SetCellWidth(long lCol, double dSize);
	void SetCellHeight(long lRow, double dSize);
	void SetCellText(long lCol, long lRow, CString strText);
	void SetCellColor(long lCol, long lRow, COLORREF color);
	void SetComboOnCell(long lCol, long lRow, CString strData, int nMaxDrop=10);
	void DispSelectRow(long lRow);
	void DispSelectRow(long lRow, long lCol);
	void DispSelectRowOverdefectEquipment(long lRow);
	void SetFieldEditEnable(long lCol, BOOL bEnable=TRUE);
	void SetFieldComboEnable(long lCol, CStringArray &strList, BOOL bEnable=FALSE);

	//hsm add 160905
	void UpdateStripOut(CString strstrQuery);

	// Generated message map functions
protected:
	//{{AFX_MSG(CSpreadSheetEx)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SPREADSHEETEX_H__022C83D2_92D7_4253_8E47_17ECFEE50E75__INCLUDED_)
