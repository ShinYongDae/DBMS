// SpreadSheetEx.cpp : implementation file
//

#include "stdafx.h"

#include "SpreadSheetEx.h"
#include "afxcoll.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#pragma warning(disable : 4996)
/////////////////////////////////////////////////////////////////////////////
// CSpreadSheetEx

CSpreadSheetEx::CSpreadSheetEx()
{
	m_strServerIp = _T("100.100.101.108");
	m_strCatalog = _T("GVISDB");
	m_strAccessID = _T("root");
	m_strPassword = _T("My3309");

	m_nMaxCols = 0;
	m_nMaxRows = 0;

	m_nOldSelectRow = 0;
	m_nNewSelectRow = 0;

	m_dHeadHeight = 16.0;
	m_rgbHeadBg = RGB(226,233,251);
	m_rgbHeadFont = RGB(0,0,0);
	m_dHeadFontSize = 13;
	m_bHeadFontBold = FALSE;
	m_strFontName = _T("±º∏≤");

	m_dCellWidth = 10.0;
	m_dCellHeight = 11.0;
	m_nCellFontSize = 9;
	m_bCellFontBold = FALSE;
	m_rgbCellBg = RGB(255,255,255);
	m_rgbCellFont = RGB(0,0,0);
	m_nCellHAlign = ALIGN_LEFT;
	m_nCellVAlign = ALIGN_LEFT;
	
	m_ColorStyleFieldIdList.RemoveAll();
	m_ComboStyleFieldIdList.RemoveAll();
	m_PrimaryKeyList.RemoveAll();
	m_FieldIDList.RemoveAll();
	m_OrgFieldIdList.RemoveAll();
	m_AliasFieldNameList.RemoveAll();
	m_ComboDataQuery.RemoveAll();
	m_strQuery = _T("");

	m_strTable = _T("");

	m_arDeletedRecord.RemoveAll();
	m_arModifiedRecord.RemoveAll();
	m_arAppendedRecord.RemoveAll();
	m_arInsertedRecord.RemoveAll();

	m_bSheetLock = FALSE;

	m_nDBType = 1;

	pArField = new CArField;
}

CSpreadSheetEx::~CSpreadSheetEx()
{
	int nCount;

	m_arDeletedRecord.RemoveAll();
	m_arModifiedRecord.RemoveAll();
	m_arAppendedRecord.RemoveAll();
	m_arInsertedRecord.RemoveAll();

	nCount = pArField->GetSize();
	if(nCount)
	{
		pArField->RemoveAll();
	}
	delete pArField;
	pArField = NULL;

	CString data ;

	m_AliasFieldNameList.RemoveAll();
	m_OrgFieldIdList.RemoveAll();
	m_FieldIDList.RemoveAll();
	m_PrimaryKeyList.RemoveAll();
	m_ColorStyleFieldIdList.RemoveAll();
	m_ComboStyleFieldIdList.RemoveAll();
	m_ComboDataQuery.RemoveAll();
	m_ComboDataList.RemoveAll();
}



/////////////////////////////////////////////////////////////////////////////
// CSpreadSheetEx message handlers
void CSpreadSheetEx::ComboCloseUp(long Col, long Row, short SelChange)
{
	CString strTemp, strFirst, strSecond;
	SetCol(Col);
	SetRow(Row);
	strTemp = GetText();
	int nPos = strTemp.Find(':', 0);
	if(nPos>0)
	{
		strFirst = strTemp.Left(nPos-1);
		strSecond = strTemp.Right(strTemp.GetLength()-nPos-1);

		strFirst.TrimRight(_T(" "));
		SetText(strFirst);
		SetForeColor(RGB_RED);

		SetCol(Col+1);
		strSecond.TrimLeft(_T(" "));
		SetText(strSecond);
		SetForeColor(RGB_RED);
	}	
}
void CSpreadSheetEx::ComboCloseUp(long Col, long Row)
{
	CString strTemp, strFirst, strSecond;
	SetCol(Col);
	SetRow(Row);
	strTemp = GetText();
	int nPos = strTemp.Find(':', 0);
	if(nPos>0)
	{
		strFirst = strTemp.Left(nPos-1);
		strSecond = strTemp.Right(strTemp.GetLength()-nPos-1);

		strFirst.TrimRight(_T(" "));
		SetText(strFirst);
		SetForeColor(RGB_RED);

		SetCol(Col+1);
		SetRow(Row);
		strSecond.TrimLeft(_T(" "));
		SetText(strSecond);
		SetForeColor(RGB_RED);
	}	
}

void CSpreadSheetEx::ModifiedRecord(long lIndex,long lField)
{
	SetReDraw(FALSE);
	
	long lRow=lIndex;
	sttField dbField;

	SetRow(lRow);
	SetCol(lField);
	SetForeColor(RGB_RED);

	int i = 0;
	int nInsertedRow,nAppendedRow,nModifiedRow;
	int nCount;
	nCount = m_arInsertedRecord.GetSize();
	for(i = 0; i < nCount; i++)
	{
		nInsertedRow = m_arInsertedRecord.GetAt(i);
		if(nInsertedRow==lRow)
			return;
	}

	nCount = m_arAppendedRecord.GetSize();
	for(i = 0; i < nCount; i++)
	{
		nAppendedRow = m_arAppendedRecord.GetAt(i);
		if(nAppendedRow==lRow)
			return;
	}

	nCount = m_arModifiedRecord.GetSize();
	for(i = 0; i < nCount; i++)
	{
		nModifiedRow = m_arModifiedRecord.GetAt(i);
		if(nModifiedRow==lRow)
			return;
	}

	if(nCount >0)
	{
		for(i = 0; i < nCount; i++)
		{
			nModifiedRow = m_arModifiedRecord.GetAt(i);
			if(nModifiedRow>=lRow)
			{
				m_arModifiedRecord.InsertAt(i,lRow);
				nCount = m_arModifiedRecord.GetSize();
				i++;
				for(; i < nCount; i++)
				{
					nModifiedRow = m_arModifiedRecord.GetAt(i);
					m_arModifiedRecord.SetAt(i,nModifiedRow+1);
				}
			}
			else
				m_arModifiedRecord.Add(lRow);
		}
		nCount = m_arModifiedRecord.GetSize();
		for(i = 0; i < nCount; i++)
		{
			nModifiedRow = m_arModifiedRecord.GetAt(i);
		}
	}
	else
		m_arModifiedRecord.Add(lRow);

	SetReDraw(TRUE);
}

void CSpreadSheetEx::InsertRecord(long lIndex)
{
	if(GetMaxRows() < 1)
	{
		CString strMsg;
		strMsg.Format(_T("[%s] "),m_strTable);
		AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	m_nMaxRows += 1;
	SetReDraw(FALSE);
	SetMaxRows(m_nMaxRows);
	
	long lRow=lIndex;
	sttField dbField;

	InsRowRange(lRow,1);
	
	SetRow(lRow);
	SetRowHeight(lRow, m_dCellHeight);

	for(long lCol=1; lCol<=m_nMaxCols; lCol++)
	{
		SetCol(lCol);
		SetCellType(TYPE_DEFAULT);

		dbField = pArField->GetAt(lCol-1);
		if(dbField.bCombo)
		{
			SetCellType(TYPE_COMBO);
			SetTypeComboBoxList (dbField.StrComboData);
			SetTypeComboBoxEditable(TRUE);
			SetTypeComboBoxMaxDrop(10);
		}

		SetFontSize((float)m_nCellFontSize);
		SetFontBold(m_bCellFontBold);
		SetForeColor(m_rgbCellFont);
		SetBackColor(m_rgbCellBg);
		SetTypeVAlign((long)m_nCellVAlign);
		SetTypeHAlign((long)m_nCellHAlign);
	}
	

	SetCol(0);
	SetRow(lRow);
	SetBackColor(m_rgbCellBg);

	m_nOldSelectRow++;
	DispSelectRow(lRow);
	SetActiveCell (1,lRow);
	SetReDraw(TRUE);
	this->SetFocus();

	int i = 0;
	int nInsertedRow;
	int nCount = m_arInsertedRecord.GetSize();
	if(nCount >0)
	{
		for(i = 0; i < nCount; i++)
		{
			nInsertedRow = m_arInsertedRecord.GetAt(i);
			if(nInsertedRow>=lRow)
			{
				m_arInsertedRecord.InsertAt(i,lRow);
				nCount = m_arInsertedRecord.GetSize();
				i++;
				for(; i < nCount; i++)
				{
					nInsertedRow = m_arInsertedRecord.GetAt(i);
					m_arInsertedRecord.SetAt(i,nInsertedRow+1);
				}
			}
		}
		nCount = m_arInsertedRecord.GetSize();
		for(i = 0; i < nCount; i++)
		{
			nInsertedRow = m_arInsertedRecord.GetAt(i);
		}
	}
	else
		m_arInsertedRecord.Add(lRow);
}

void CSpreadSheetEx::AppendRecord()
{
	m_nMaxRows += 1;
	SetReDraw(FALSE);
	SetMaxRows(m_nMaxRows);

	long lRow=m_nMaxRows;
	sttField dbField;
	{
		SetRow(lRow);
		SetRowHeight(lRow, m_dCellHeight);

		for(long lCol=1; lCol<=m_nMaxCols; lCol++)
		{
			SetCol(lCol);
			SetCellType(TYPE_DEFAULT);

			dbField = pArField->GetAt(lCol-1);
			if(dbField.bCombo)
			{
				SetCellType(TYPE_COMBO);
				SetTypeComboBoxList (dbField.StrComboData);
				SetTypeComboBoxEditable(TRUE);
				SetTypeComboBoxMaxDrop(10);
			}

			//if(dbField.bPrimary)
			//{
			//	SetCellType(TYPE_STATIC);
			//}

			SetFontSize((float)m_nCellFontSize);
			SetFontBold(m_bCellFontBold);
			SetForeColor(m_rgbCellFont);
			SetBackColor(m_rgbCellBg);
			SetTypeVAlign((long)m_nCellVAlign);
			SetTypeHAlign((long)m_nCellHAlign);
		}
	}	

	SetCol(0);
	SetRow(lRow);
	SetBackColor(m_rgbCellBg);

	DispSelectRow(lRow);
	SetActiveCell (1,lRow); 

	SetReDraw(TRUE);
	this->SetFocus();
	
	m_arAppendedRecord.Add(lRow);
}

void CSpreadSheetEx::DeleteRecord(long lIndex)
{
	if(GetMaxRows() < 1)
	{
		CString strMsg;
		strMsg.Format(_T("[%s] "),m_strTable);
		AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}
	
	int i,nCount;
	long lDeleteRow = lIndex; // GetRow();
	CString strMsg;
	if(lDeleteRow != 0)
	{
		// Check appended record
		nCount = m_arAppendedRecord.GetSize();
		for(i=0; i < nCount; i++)
		{
			long lAppendedRow = m_arAppendedRecord.GetAt(i); // get deleted record index number
			if(lDeleteRow == lAppendedRow)
			{
				strMsg=_T("Are you sure delete from sheet?\r\n\r\n Record info: Appended Record");
				if(AfxMessageBox(strMsg, MB_YESNOCANCEL | MB_ICONQUESTION | MB_DEFBUTTON2)!=IDYES)
					return;

				// delete appended record from spread
				m_arAppendedRecord.RemoveAt(i);
				for(;i < nCount-1; i++)
				{
					lAppendedRow = m_arAppendedRecord.GetAt(i);
					m_arAppendedRecord.SetAt(i,lAppendedRow-1);
				}


				SetReDraw(FALSE);
				long lNumRows = 1;
				// delete contents
				DelRowRange(lDeleteRow, lNumRows);
				
				m_nMaxRows -= 1;
				SetMaxRows(m_nMaxRows);
				DispSelectRow(lDeleteRow);
				SetReDraw(TRUE);
				return;
			}
		}		
	
		// Check inserted record
		nCount = m_arInsertedRecord.GetSize();
		for(i=0; i < nCount; i++)
		{
			long lInsertedRow = m_arInsertedRecord.GetAt(i); // get deleted record index number
			if(lDeleteRow == lInsertedRow)
			{
				strMsg=_T("Are you sure delete from sheet?\r\n\r\n Record info: Inserted Record");
				if(AfxMessageBox(strMsg, MB_YESNOCANCEL | MB_ICONQUESTION | MB_DEFBUTTON2)!=IDYES)
					return;
				
				// Delete inserted record from spread
				m_arInsertedRecord.RemoveAt(i);
				for(;i < nCount-1; i++)
				{
					lInsertedRow = m_arInsertedRecord.GetAt(i);
					m_arInsertedRecord.SetAt(i,lInsertedRow-1);
				}


				SetReDraw(FALSE);
				long lNumRows = 1;
				// delete contents
				DelRowRange(lDeleteRow, lNumRows);		
				m_nMaxRows -= 1;
				SetMaxRows(m_nMaxRows);
				DispSelectRow(lDeleteRow);
				SetReDraw(TRUE);
				return;
			}
		}		

		// delete query
		int nPrimaryKeyCount=0;
		sttField dbField;
		CString strQuery,strTemp;
		strMsg=_T("Are you sure Delete this record from DB?\r\n\r\n Record info: ");
		CString strPrimaryKeyData,strPrimaryKeyName;
	
		strQuery.Format(_T("DELETE FROM %s WHERE "),m_strTable);
		
		nCount= pArField->GetSize();
		SetRow(lDeleteRow);
		for(i = 0; i < nCount; i++)
		{
			dbField = pArField->GetAt(i);
			if(dbField.bPrimary)
			{
				SetCol(i+1);
				strPrimaryKeyData = GetText();
				strPrimaryKeyName = dbField.strFieldId;
				if(nPrimaryKeyCount>0)
				{
					strQuery += _T(" AND ");
					strMsg += _T(" and ");
				}
				strTemp.Format(_T("%s = '%s'"),strPrimaryKeyName,strPrimaryKeyData);
				strQuery += strTemp;
				strMsg += strTemp;
				nPrimaryKeyCount++;
			}
		}
		
		if(AfxMessageBox(strMsg, MB_YESNOCANCEL | MB_ICONQUESTION | MB_DEFBUTTON2)==IDYES)
		{
			if(!m_dataSource.DirectQuery((LPCTSTR)strQuery))
			{
		/*		CString strMsg;
				strMsg.Format("Error occur at m_dataSource.DirectQuery() at DeleteRecord()\r\n%s",m_dataSource.GetLastError());
				Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
				return;
		*/
			}

			m_arDeletedRecord.Add(lDeleteRow);
			SetReDraw(FALSE);
			long lNumRows = 1;
			// delete contents
			DelRowRange(lDeleteRow,lNumRows);
			//DeleteRows(lDeleteRow, lNumRows);
			m_nMaxRows -= 1;
			SetMaxRows(m_nMaxRows);
			DispSelectRow(lDeleteRow);
			SetReDraw(TRUE);
		}
		else
			return;
	}

	Init(m_strQuery,m_strTable,lDeleteRow);
	
}

void CSpreadSheetEx::UpdateStripOut(CString strstrQuery)
{
	if(!m_dataSource.ExecuteQuery((LPCTSTR)strstrQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at UpdateStripOut() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}
}

void CSpreadSheetEx::SaveRecord(long lIndex)
{
	int i,j;
	int nCount;
	CString strCellData;
	CString strQuery;
	sttField dbField;
	CString strTemp,strData, strTemp2;
	int nField = 0;
	CStringArray m_strArrayData;
	
	// check and update modified record
	nCount = m_arModifiedRecord.GetSize();

	for(i=0; i < nCount; i++)
	{
		long lRow = m_arModifiedRecord.GetAt(i); // get deleted record index number
		SetRow(lRow);
		for(j = 0; j < GetMaxCols(); j++)
		{
			SetCol(j+1);
			strCellData = GetText();
			m_strArrayData.Add(strCellData);
		}
		if(m_nDBType ==1)
		{
			strQuery.Format(_T("CALL USP_%s("),m_strTable);
		}
		else
		{
			strQuery.Format(_T("EXEC USP_%s"),m_strTable);
		}
		

		j = 0;
		for(nField=0; nField<GetMaxCols(); nField++)
		{
			dbField = pArField->GetAt(nField);
			if(!dbField.bExtField)
			{
				strData = m_strArrayData.GetAt(nField);
				if(strData.GetLength()==0)
					strData = _T("");
				if(j>0) 
					strQuery += _T(", ");
				if(m_nDBType ==1)
				{
					strTemp2 = m_FieldIDList.GetAt(nField);
					if(strTemp2.Find(_T("END_DATE")) >=0 || strTemp2.Find(_T("DT_REGIST")) >=0 || strTemp2.Find(_T("DT_UPDATE")) >=0)
					{
						strData = _T("1990-01-01 00:00:00");
					}
					strTemp.Format(_T("'%s'"),strData); 
				}
				else
				{
					strTemp.Format(_T(" @%s = '%s'"),dbField.strFieldId,strData); 
				}
				
				strQuery += strTemp;
				j++;
			}
					
		}
		if(m_nDBType ==1)
		{
			strQuery +=_T(")");
		}	
	//	TRACE("%s\n",strQuery);
		if(!m_dataSource.DirectQuery((LPCTSTR)strQuery))
		{
			CString strMsg;
			strMsg.Format(_T("Error occur at m_dataSource.DirectQuery() at SaveRecord()\r\n%s"),m_dataSource.GetLastError());
			Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
			return;
		}

		m_strArrayData.RemoveAll();
	}
	m_arModifiedRecord.RemoveAll();

	// check and update append record
	// append query
	nCount = m_arAppendedRecord.GetSize();
	for(i=0; i < nCount; i++)
	{

		long lRow = m_arAppendedRecord.GetAt(i); // get deleted record index number
		SetRow(lRow);
		for(j = 0; j < GetMaxCols(); j++)
		{
			SetCol(j+1);
			strCellData = GetText();
			m_strArrayData.Add(strCellData);
		}
		if(m_nDBType ==1)
		{
			strQuery.Format(_T("CALL USP_%s("),m_strTable);
		}
		else
		{
			strQuery.Format(_T("EXEC USP_%s"),m_strTable);
		}
		

		j = 0;
		for(nField=0; nField<GetMaxCols(); nField++)
		{
			dbField = pArField->GetAt(nField);
			if(!dbField.bExtField)
			{
				strData = m_strArrayData.GetAt(nField);
				if(strData.GetLength()==0)
					strData = "";
				if(j>0) 
					strQuery += ", ";
				if(m_nDBType ==1)
				{
					SetRow(0);
					SetCol(nField+1);
					strTemp2 = GetText();
					strTemp2 = m_FieldIDList.GetAt(nField);
					if(strTemp2.Find(_T("END_DATE")) >=0 || strTemp2.Find(_T("DT_REGIST")) >=0 || strTemp2.Find(_T("DT_UPDATE")) >=0)
					{
						strData = _T("1990-01-01 00:00:00");
					}

					strTemp.Format(_T("'%s'"),strData); 
				}
				else
				{
					strTemp.Format(_T(" @%s = '%s'"),dbField.strFieldId,strData); 
					
				}		
				strQuery += strTemp;
				j++;
			}
		}
		if(m_nDBType ==1)
		{
			strQuery +=_T(")");
		}	
		
		TRACE(_T("%s\n"),strQuery);
		if(!m_dataSource.DirectQuery((LPCTSTR)strQuery))
		{
			CString strMsg;
			strMsg.Format(_T("Error occur at m_dataSource.DirectQuery() at SaveRecord()\r\n%s"),m_dataSource.GetLastError());
			Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
			return;
		}

		m_strArrayData.RemoveAll();
	}
	m_arAppendedRecord.RemoveAll();

	// check and update inserted record
	nCount = m_arInsertedRecord.GetSize();
	for(i=0; i < nCount; i++)
	{

		long lRow = m_arInsertedRecord.GetAt(i); // get deleted record index number
		SetRow(lRow);
		for(j = 0; j < GetMaxCols(); j++)
		{
			SetCol(j+1);
			strCellData = GetText();
			m_strArrayData.Add(strCellData);
		}
		if(m_nDBType ==1)
		{
			strQuery.Format(_T("CALL USP_%s("),m_strTable);
		}
		else
		{
			strQuery.Format(_T("EXEC USP_%s"),m_strTable);
		}
		

		j = 0;
		for(nField=0; nField<GetMaxCols(); nField++)
		{
			dbField = pArField->GetAt(nField);
			if(!dbField.bExtField)
			{
				strData = m_strArrayData.GetAt(nField);
				if(strData.GetLength()==0)
				{
					strData = _T("");
				}
				if(m_nDBType ==1)
				{
					strTemp.Format(_T("'%s',"),strData); 
				}
				else
				{
					strTemp.Format(_T(" @%s = '%s', "),dbField.strFieldId,strData); 
				}
				
				strQuery += strTemp;
				j++;
			}
		}
		ExtractLastWordcom(&strQuery,',');
		if(m_nDBType ==1)
		{
			strQuery +=_T(")");
		}
		if(!m_dataSource.DirectQuery((LPCTSTR)strQuery))
		{
			CString strMsg;
			strMsg.Format(_T("Error occur at m_dataSource.DirectQuery() at SaveRecord()\r\n%s"),m_dataSource.GetLastError());
			Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
			return;
		}

		m_strArrayData.RemoveAll();
	}
	m_arInsertedRecord.RemoveAll();

	
	// check and update Deleteded record	
	m_arDeletedRecord.RemoveAll();
	
	Init(m_strQuery,m_strTable,lIndex);
}

void CSpreadSheetEx::SaveEquipment(long lIndex)
{
	/*
	int i,j;
	int nCount;
	CString strCellData;
	CString strQuery;
	sttField dbField;
	CString strTemp,strData, strTemp2;
	int nField = 0;
	CStringArray m_strArrayData;
	sttEquipInfo EquipInfo;
	// check and update modified record
	nCount = m_arModifiedRecord.GetSize();

	CString strBiz_code, strequipcode, strBiz_name, strprcname, strmng_code, strEquip_yn, strEquipModel_code, strEquipType_name;

	for(i=0; i < nCount; i++)
	{
		strQuery.Empty();
		long lRow = m_arModifiedRecord.GetAt(i); // get deleted record index number
		SetRow(lRow);
	
		strQuery.Format(_T("CALL USP_EQUIPMENT("));
		//get info
		j = 1;
		SetCol(j++);
		EquipInfo.Equip_Code = GetText();  //equip_code
		SetCol(j++);
		EquipInfo.Equip_Name = GetText();  //equip_name
		SetCol(j++);	
		EquipInfo.Equip_Nick = GetText();   //nick	
		SetCol(j++);
		strBiz_code = GetText();  //biz_code
		SetCol(j++);
		strBiz_name = GetText();  //biz_name ,5
		SetCol(j++);
		EquipInfo.Process = GetText();   //prc_code
		SetCol(j++);
		strprcname = GetText();   //prc_name
		SetCol(j++);
		EquipInfo.IP_Addr = GetText();   //IP		
		SetCol(j++);
		EquipInfo.MAC_Addr = GetText();   //MAC
		SetCol(j++);
		strequipcode = GetText();   //EQUIPTYPE_CODE  ,10
		SetCol(j++);
		strEquipType_name = GetText();   //EQUIPTYPE_name 
		SetCol(j++);
		EquipInfo.Maker = GetText();   //maker
		SetCol(j++);
		EquipInfo.Agent = GetText();   //agent
		SetCol(j++);
		EquipInfo.Install_Loc = GetText();   //install_local
		SetCol(j++);
		EquipInfo.Purchase_Date = GetText();   //pur_date
		SetCol(j++);
		EquipInfo.Purchase_Price = GetText();   //pur-amt  ,15
		SetCol(j++);
		strmng_code = GetText();   //mng_code
		SetCol(j++);
		EquipInfo.Manager_Name = GetText();   //mng_name
		SetCol(j++);		
		EquipInfo.Using_Status = GetText();   //use_yn
		SetCol(j++);		
		strEquip_yn= GetText();   //equip_yn  		
		SetCol(j++);	
		strEquipModel_code = GetText();   //equipmodel_code  ,20
		SetCol(j++);
		EquipInfo.Equip_Model_Name = GetText();   //equipmodel_name
		SetCol(j++);
		EquipInfo.Soft_Version = GetText();   //sw_version
		SetCol(j++);	
		EquipInfo.Register_Name = GetText();   //id_regist   
		SetCol(j++);
		SetCol(j++);		
		EquipInfo.Register_Date = GetText();   //dt_regist  //25
		
		

		//set
		strTemp.Format(_T("'%s',"),  EquipInfo.Equip_Code);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  strBiz_code);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Process);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Equip_Name);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Equip_Nick);  //5
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.IP_Addr);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.MAC_Addr);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"), strequipcode);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Maker);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Agent);//10
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Install_Loc);  
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Purchase_Date);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Purchase_Price);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  strmng_code);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Using_Status); //15
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"), strEquipModel_code);  
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Soft_Version);
		strQuery +=strTemp;
		strTemp.Format(_T("'',"),  EquipInfo.Register_Name);
		strQuery +=strTemp;
		strTemp.Format(_T("'',"),  EquipInfo.Register_Date);
		strQuery +=strTemp;
		strTemp.Format(_T("'',"));//20
		strQuery +=strTemp;
		strTemp.Format(_T("'')"));   
		strQuery +=strTemp;
		
		
		//	TRACE("%s\n",strQuery);
		if(!m_dataSource.DirectQuery((LPCTSTR)strQuery))
		{
			CString strMsg;
			strMsg.Format(_T("Error CCccur at m_dataSource.DirectQuery() at SaveRecord()\r\n%s"),m_dataSource.GetLastError());
			Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
			return;
		}

 		m_strArrayData.RemoveAll();
	}
	m_arModifiedRecord.RemoveAll();

	// check and update append record
	// append query
	nCount = m_arAppendedRecord.GetSize();
	for(i=0; i < nCount; i++)
	{

		long lRow = m_arAppendedRecord.GetAt(i); // get deleted record index number
		SetRow(lRow);
		
		strQuery.Format(_T("CALL USP_EQUIPMENT("));
		//get info
		j = 1;
		SetCol(j++);
		EquipInfo.Equip_Code = GetText();  //equip_code
		SetCol(j++);
		EquipInfo.Equip_Name = GetText();  //equip_name
		SetCol(j++);	
		EquipInfo.Equip_Nick = GetText();   //nick	
		SetCol(j++);
		strBiz_code = GetText();  //biz_code
		SetCol(j++);
		strBiz_name = GetText();  //biz_name ,5
		SetCol(j++);
		EquipInfo.Process = GetText();   //prc_code
		SetCol(j++);
		strprcname = GetText();   //prc_name
		SetCol(j++);
		EquipInfo.IP_Addr = GetText();   //IP		
		SetCol(j++);
		EquipInfo.MAC_Addr = GetText();   //MAC
		SetCol(j++);
		strequipcode = GetText();   //EQUIPTYPE_CODE  ,10
		SetCol(j++);
		strEquipType_name = GetText();   //EQUIPTYPE_name 
		SetCol(j++);
		EquipInfo.Maker = GetText();   //maker
		SetCol(j++);
		EquipInfo.Agent = GetText();   //agent
		SetCol(j++);
		EquipInfo.Install_Loc = GetText();   //install_local
		SetCol(j++);
		EquipInfo.Purchase_Date = GetText();   //pur_date
		SetCol(j++);
		EquipInfo.Purchase_Price = GetText();   //pur-amt  ,15
		SetCol(j++);
		strmng_code = GetText();   //mng_code
		SetCol(j++);
		EquipInfo.Manager_Name = GetText();   //mng_name
		SetCol(j++);		
		EquipInfo.Using_Status = GetText();   //use_yn
		SetCol(j++);		
		strEquip_yn= GetText();   //equip_yn  		
		SetCol(j++);	
		strEquipModel_code = GetText();   //equipmodel_code  ,20
		SetCol(j++);
		EquipInfo.Equip_Model_Name = GetText();   //equipmodel_name
		SetCol(j++);
		EquipInfo.Soft_Version = GetText();   //sw_version
		SetCol(j++);	
		EquipInfo.Register_Name = GetText();   //id_regist   
		SetCol(j++);
		SetCol(j++);		
		EquipInfo.Register_Date = GetText();   //dt_regist  //25



		//set
		strTemp.Format(_T("'%s',"),  EquipInfo.Equip_Code);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  strBiz_code);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Process);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Equip_Name);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Equip_Nick);  //5
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.IP_Addr);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.MAC_Addr);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"), strequipcode);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Maker);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Agent);//10
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Install_Loc);  
		strQuery +=strTemp;
		if(EquipInfo.Purchase_Date.GetLength()==0)
		{
			strTemp.Format(_T("'1990-01-01 00:00:00',"));
		}
		else
		{
			strTemp.Format(_T("'%s',"),  EquipInfo.Purchase_Date);
		}		
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Purchase_Price);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  strmng_code);
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Using_Status); //15
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"), strEquipModel_code);  
		strQuery +=strTemp;
		strTemp.Format(_T("'%s',"),  EquipInfo.Soft_Version);
		strQuery +=strTemp;
		strTemp.Format(_T("'',"),  EquipInfo.Register_Name);
		strQuery +=strTemp;
		strTemp.Format(_T("'',"),  EquipInfo.Register_Date);
		strQuery +=strTemp;
		strTemp.Format(_T("'',"));//20
		strQuery +=strTemp;
		strTemp.Format(_T("'')"));   
		strQuery +=strTemp;
		

		TRACE(_T("%s\n"),strQuery);
		if(!m_dataSource.DirectQuery((LPCTSTR)strQuery))
		{
			CString strMsg;
			strMsg.Format(_T("Error occur at m_dataSource.DirectQuery() at SaveRecord()\r\n%s"),m_dataSource.GetLastError());
			Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
			return;
		}

		m_strArrayData.RemoveAll();
	}
	m_arAppendedRecord.RemoveAll();



	// check and update Deleteded record	
	m_arDeletedRecord.RemoveAll();

	Init(m_strQuery,m_strTable,lIndex);
	*/
}


void CSpreadSheetEx::SetDBServer(CString strServerIp,CString strCatalog,CString strAccessID,CString strPassword)
{
	m_strServerIp = strServerIp;
	m_strCatalog = strCatalog;
	m_strAccessID = strAccessID;
	m_strPassword = strPassword;
}

void CSpreadSheetEx::SetComboOnSpread() // Pk : Primary Key...
{
	CString strData;
	int nMaxRows = m_dataSource.m_pRS->RecordCount;
	int nMaxCols = m_dataSource.m_pRS->Fields->GetCount();
	
	long lCol;
	if(nMaxCols > 0)
	{
		for(lCol=1; lCol<=m_nMaxCols; lCol++)
		{
			SetRow(0);
			SetCol(lCol);
			strData = GetText();

			CString data;
			int j = m_ComboStyleFieldIdList.GetSize();
			for(long i = 0 ; i < m_ComboStyleFieldIdList.GetSize() ; i++)
			{
				data = m_ComboStyleFieldIdList.GetAt(i);
				if(data == strData)
				{
					SetColWidth(lCol, 20.0);
					SetFieldComboEnable(lCol, m_ComboDataList, TRUE);
				}
			}
		}
	}
}

void CSpreadSheetEx::SetFieldComboEnable(long lCol, CStringArray &strList, BOOL bEnable)
{
	if(lCol < 0)
		return;

	CString data, strComboData=_T("");
	for(long i = 0 ; i < strList.GetSize() ; i++)
	{
		data = strList.GetAt(i);
		if(i == strList.GetSize() - 1)
			strComboData += data;
		else
			strComboData += (data+_T("\t"));
	}

	SetCol(lCol);
	for(long lRow=1; lRow<=m_nMaxRows; lRow++)
	{
		SetRow(lRow);
		if(bEnable)
		{
			SetComboOnCell(lCol, lRow, strComboData);
		}
		else
			SetCellType(TYPE_DEFAULT);
	}
}

void CSpreadSheetEx::SetComboOnCell(long lCol, long lRow, CString strData, int nMaxDrop)
{
	// ComboBox in Cell
	SetCol(lCol);
	SetRow(lRow);	
	//SetLock(FALSE);
	//SetCellType(TYPE_DEFAULT);
	SetCellType(TYPE_COMBO);
	SetTypeComboBoxList (strData);
	SetTypeComboBoxEditable(TRUE);
	SetTypeComboBoxMaxDrop(nMaxDrop);

	SetFontSize((float)m_nCellFontSize);
	SetFontBold(m_bCellFontBold);
	SetForeColor(m_rgbCellFont);
	SetBackColor(m_rgbCellBg);
	SetTypeVAlign((long)m_nCellVAlign);
	SetTypeHAlign((long)m_nCellHAlign);
}

CString CSpreadSheetEx::GetComboDataList(CString strFieldName)
{
	FindComboDataOnDb(strFieldName);
	
	CString data, strComboData= _T("");
	for(long i = 0 ; i < m_ComboDataList.GetSize() ; i++)
	{
		data = m_ComboDataList.GetAt(i);
		if(i == m_ComboDataList.GetSize() - 1)
			strComboData += data;
		else
			strComboData += (data+_T("\t"));
	}
	return strComboData;
}

void CSpreadSheetEx::FindComboDataOnDb(CString strFieldName) 
{
	int i=0;
	int nCount = 0;
	int nIndex = 0;
	CString strQuery;

	nCount = m_ComboStyleFieldIdList.GetSize();
	CString strComboFieldName;
	for(nIndex = 0; nIndex < nCount; nIndex++)
	{
		strComboFieldName = m_ComboStyleFieldIdList.GetAt(nIndex);
		if(strFieldName == strComboFieldName)
			break;
	}
	
	if(!(nCount = m_ComboDataQuery.GetSize()))
		return;

	strQuery = m_ComboDataQuery.GetAt(nIndex);
	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindComboDataOnDb()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	int nMaxRows = m_dataSource.m_pRS->RecordCount;
	m_ComboDataList.RemoveAll();

	if(nMaxRows > 0)
	{
		CString strData;
		_variant_t vValue; 

		m_dataSource.m_pRS->MoveFirst();
		for(long lRow=0; lRow<nMaxRows; lRow++)
		{
			vValue = m_dataSource.m_pRS->Fields->Item[(long)0]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
			}
			else
				strData.Empty();

			m_ComboDataList.Add(strData);

			m_dataSource.m_pRS->MoveNext();
		}
	}

}

void CSpreadSheetEx::FindComboDataQueryOnDb()
{
	CString strQuery;
	strQuery.Format(_T("SELECT SQLQUERY FROM COLUMNINFO WHERE TABLE_NAME = '%s' AND COLUMN_TYPE = '1'"),m_strTable);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindComboDataQueryOnDb()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	m_ComboDataQuery.RemoveAll();

	int nMaxRows = m_dataSource.m_pRS->RecordCount;

	if(nMaxRows > 0)
	{
		CString strData;
		_variant_t vValue; 
		m_dataSource.m_pRS->MoveFirst();

		for(long lRow=0; lRow<nMaxRows; lRow++)
		{
			vValue = m_dataSource.m_pRS->Fields->Item[(long)0]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
			}
			else
				strData.Empty();

			m_ComboDataQuery.Add(strData);
			m_dataSource.m_pRS->MoveNext();
		}			
	}
}

void CSpreadSheetEx::FindComboFieldIdOnDb()
{
	CString strQuery;
	strQuery.Format(_T("SELECT COLUMN_NAME FROM COLUMNINFO WHERE TABLE_NAME = '%s' AND COLUMN_TYPE = '1'"),m_strTable);
	
	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindComboFieldIdOnDb()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	int nMaxRows = m_dataSource.m_pRS->RecordCount;

	m_ComboStyleFieldIdList.RemoveAll();
	
	if(nMaxRows > 0)
	{
		CString strData;
		_variant_t vValue; 

		m_dataSource.m_pRS->MoveFirst();
		for(long lRow=1; lRow<=nMaxRows; lRow++)
		{
			vValue = m_dataSource.m_pRS->Fields->Item[(long)0]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
			}
			else
				strData.Empty();
			m_ComboStyleFieldIdList.Add(strData);
			if(lRow < nMaxRows)
				m_dataSource.m_pRS->MoveNext();
		}
	}
}

void CSpreadSheetEx::SetPkOnSpread() // Pk : Primary Key...
{
	CString strData;
	int nMaxRows = m_dataSource.m_pRS->RecordCount;
	int nMaxCols = m_dataSource.m_pRS->Fields->GetCount();
	
	long lCol;
	if(nMaxCols > 0)
	{
		for(lCol=1; lCol<=m_nMaxCols; lCol++)
		{
			SetRow(0);
			SetCol(lCol);
			strData = GetText();

			CString data;
			for(long i = 0 ; i < m_PrimaryKeyList.GetSize() ; i++)
			{
				data = m_PrimaryKeyList.GetAt(i);
				if(data == strData)
				{
					SetFieldEditEnable(lCol, FALSE);
				}
			}
		}
	}
}

void CSpreadSheetEx::FindPkOnDb() // Pk : Primary Key...
{
	CString strQuery;

	if(m_nDBType ==1)
	{
		strQuery.Format(_T("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE CONSTRAINT_NAME LIKE 'P%%' AND TABLE_NAME = '%s' "),m_strTable);
	}
	else
	{
		strQuery.Format(_T("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE WHERE CONSTRAINT_NAME LIKE 'P%%' AND TABLE_NAME = '%s' "),m_strTable);
	}

	

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindPkOnDb()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	m_PrimaryKeyList.RemoveAll();

	int nMaxRows = m_dataSource.m_pRS->RecordCount;
	if(nMaxRows > 0)
	{
		CString strData;
		_variant_t vValue; 

		m_dataSource.m_pRS->MoveFirst();
		for(long lRow=1; lRow<=nMaxRows; lRow++)
		{
			vValue = m_dataSource.m_pRS->Fields->Item[(long)0]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
			}
			else
				strData.Empty();
			m_PrimaryKeyList.Add(strData);

			if(lRow < nMaxRows)
				m_dataSource.m_pRS->MoveNext();
		}
	}
}

void CSpreadSheetEx::FindFieldIdOnDb()
{
	CString strName;
	_variant_t vColName;
	int nField;

	m_FieldIDList.RemoveAll();

	if(!m_dataSource.ExecuteQuery(m_strQuery))
	{
		AfxMessageBox(_T("Error occur at m_dataSource.ExecuteQuery()"));
		return;
	}
		
	for(nField=0; nField< m_nMaxCols ; nField++)
	{
		vColName = m_dataSource.m_pRS->Fields->Item[(long)nField]->GetName();
		strName =  vColName.bstrVal;
		m_FieldIDList.Add(strName);
	}
}

BOOL CSpreadSheetEx::GetAliasFieldNameOnDb(CString strFieldId,CString &strAliasname)
{
	CString strName;
	_variant_t vColName;
	int nMaxCols,nMaxRows;
	CString strQuery;

	strQuery.Format(_T("SELECT ENG_NAME FROM DISPLAY WHERE ALIAS_NAME = '%s'"),strFieldId);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at GetAliasFieldNameOnDb()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return FALSE;
	}

	nMaxCols = m_dataSource.m_pRS->Fields->Count ;
	nMaxRows = m_dataSource.m_pRS->RecordCount;
	if(nMaxRows>0)
	{
		m_dataSource.m_pRS->MoveFirst();
		for(int nRow=0; nRow< nMaxRows ; nRow++)
		{
			for(int nCol=0; nCol< nMaxCols ; nCol++)
			{
				vColName = m_dataSource.m_pRS->Fields->Item[(long)nCol]->Value;
				strName =  vColName.bstrVal;
				strAliasname = strName;
				//m_FieldIDList.Add(strName);
			}
			m_dataSource.m_pRS->MoveNext();		
		}
	}
	else
	{
		strAliasname = "Not Defined";
		return FALSE;
	}
	return TRUE;
}

void CSpreadSheetEx::FindOrgFieldIdOnDb()
{
	CString strQuery;
	strQuery.Format(_T("SELECT TABLE_NAME, COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME ='%s' "),m_strTable);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindOrgFieldIdOnDb()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	CString strName;
	_variant_t vColName;		
	int nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	int nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();

	if(nMaxRows > 0)
	{
		CString strData;
		_variant_t vValue; 

		m_dataSource.m_pRS->MoveFirst();
		for(long lRow=0; lRow< nMaxRows ; lRow++)
		{
			vValue = m_dataSource.m_pRS->Fields->Item[(long)1]->Value;
			vValue.ChangeType(VT_BSTR);
			strData = vValue.bstrVal ;
			m_OrgFieldIdList.Add(strData);
			m_dataSource.m_pRS->MoveNext();
		}
	}
}
CString CSpreadSheetEx::SetTest(CString strQuery)
{

	time_t st, stquery, stspread;

	st = clock();

	CString strMsg;
	if (!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"), m_dataSource.GetLastError());
		Log(strMsg); 
		//AfxMessageBox(strMsg, MB_ICONSTOP);
		return strMsg;
	}

	stquery = (clock() - st);

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();

	if (m_nMaxCols < 1)
	{
		//AfxMessageBox(_T("No Field."));
		//AfxMessageBox("Field∞° æ¯Ω¿¥œ¥Ÿ.");
		strMsg = _T("No Field.");
		return strMsg;
	}
	

	SetReDraw(FALSE);
	SetShadowColor(m_rgbHeadBg, m_rgbHeadFont, COLOR_3DSHADOW, COLOR_3DHILIGHT);
	//SetShadowColor((unsigned long)m_rgbHeadBg);			
	//SetShadowText((unsigned long)m_rgbHeadFont);

	SS_CELLTYPE stype;
	SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER, 128, SS_CHRSET_CHR, SS_CASE_NOCASE);

	HFONT hfont;
	hfont = CreateFont(m_dHeadFontSize, 0, 0, 0, m_bHeadFontBold, 0, 0, 0, 0, 0, 0, 0, 0, _T("πŸ≈¡"));


	//SetFontBold((BOOL)m_bHeadFontBold);
	//SetFontSize((float)m_dHeadFontSize);
	//SetFontName((LPCTSTR)m_strFontName);

	SetMaxCols((long)m_nMaxCols);
	SetMaxRows((long)m_nMaxRows);
	SetTypeHAlign(SSS_ALIGN_CENTER);

	//Column Caption
	SetRow(0L);

	SetColWidth(0, 10);

	SetRowHeight(-1, m_dHeadHeight); // COL(0)'S HEIGHT
	

	SetRow(-1);
	SetCol(-1);
	SetFontSize(9);
	//hfont = CreateFont(13,0,0,0,m_bHeadFontBold,0,0,0,0,SHIFTJIS_CHARSET,0,0,0,_T("πŸ≈¡"));
	hfont = CreateFont(13, 0, 0, 0, m_bHeadFontBold, 0, 0, 0, 0, SHIFTJIS_CHARSET, 0, 0, 0, _T("MS PMincho"));
	SetFont(SS_ALLCOLS, SS_ALLROWS, hfont, TRUE);

	SetCellType(SS_ALLCOLS, SS_ALLROWS, &stype);

	CString strData;

	_variant_t vValue;
	CString strName;
	SetRow(0);
	
	for (long lCol = 0; lCol < m_nMaxCols; lCol++)
	{
		SetCol(lCol+1);
		vValue = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();
		strName = vValue.bstrVal;
		SetText(strName);
	}

	m_dataSource.m_pRS->MoveFirst();
	for (long lRow = 1; lRow <= m_nMaxRows; lRow++)
	{		
		SetRow(lRow);

		for (long lCol = 1; lCol <= m_nMaxCols; lCol++)
		{
			SetCol(lCol);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol - 1)]->Value;

			if (vValue.vt != VT_NULL)
			{

				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal;

				SetText(strData);			
			}

		}
		m_dataSource.m_pRS->MoveNext();
	}
	SetReDraw(TRUE);
	stspread = (clock() - st);
	
	strMsg.Format(_T("Query Time#1 : %d\nTime#2 : %d\n"), stquery, stspread);
	//AfxMessageBox(strMsg);
	return strMsg;
}

void CSpreadSheetEx::Init(CString strQuery,CString strTableName,long lActiveRow,BOOL bSheetLock)
{
	m_bSheetLock = bSheetLock;
	m_strQuery = strQuery;
	m_strTable = strTableName;
	
	m_dataSource.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strPassword);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();
	if(m_nMaxRows== -1)
	{
		AfxMessageBox(_T("No Record."));
		//AfxMessageBox("«ÿ¥Á«œ¥¬ ∑πƒ⁄µÂ∞° æ¯Ω¿¥œ¥Ÿ.");
		return;
	}
	if(lActiveRow>m_nMaxRows)
		lActiveRow = m_nMaxRows;
	if(lActiveRow==0)
		lActiveRow = 1;

	m_arModifiedRecord.RemoveAll();
	m_arAppendedRecord.RemoveAll();
	m_arInsertedRecord.RemoveAll();
	if(pArField->GetSize())
		pArField->RemoveAll();

	FindFieldIdOnDb();
	FindOrgFieldIdOnDb();
	
	FindPkOnDb();
	
	FindComboFieldIdOnDb();
	FindComboDataQueryOnDb();

	FindColorFieldIdOnDb();
	SetSpreadInfo();
	DispSpread();

	DispSelectRow(lActiveRow);
	SetActiveCell (1,lActiveRow);
	
	SetUserColAction(1);
}

int CSpreadSheetEx::InitOver(CString strQuery,CString strTableName,long lActiveRow,BOOL bSheetLock)
{	
	_variant_t vValue;
	_variant_t vColName;
	CString strData;
	CString strTitle;
	int nStringLenth=10;



	Reset();	
	m_dataSource.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strPassword);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return 0;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();
	
	SetFontSize(13);
	SetRowHeight(-1,16);
	SetMaxCols(m_nMaxCols);
	SetMaxRows(m_nMaxRows);
	SetTypeHAlign(SSS_ALIGN_CENTER);

	SS_CELLTYPE stype;
	SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER ,128,SS_CHRSET_CHR,SS_CASE_NOCASE);

	if(m_nMaxRows>0)
	{	
		m_dataSource.m_pRS->MoveFirst();
	}
	SetRow(0);
	for(long lCol=0; lCol<m_nMaxCols; lCol++)
	{
		SetCol(lCol+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();
		if(vColName.vt != VT_NULL)
		{
			vColName.ChangeType(VT_BSTR);
			strTitle = vColName.bstrVal ;
			if(lCol==5)
			{
				strTitle="DEF_CNT";
			}
			else if(lCol==6)
			{
				strTitle="INDEX";
			}
			else if(lCol==7)
			{
				strTitle="STATUS";
			}
			else if(lCol==8)
			{
				strTitle="START_DATE";
			}	
			else if(lCol==9)
			{
				strTitle="END_DATE";
			}
			SetCellType(lCol+1, 0, &stype);
			SetText(strTitle);
		}
	}

	if(m_nMaxRows>0)
	{
		m_dataSource.m_pRS->MoveFirst();
	}
	for(long lRow=0; lRow<m_nMaxRows; lRow++)
	{
		SetRow(lRow+1);
		
		for(long lCol=0; lCol<m_nMaxCols; lCol++)
		{
			SetCol(lCol+1);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
				SetCellType(lCol+1, lRow+1, &stype);
				SetText(strData);
				if(lRow==0)
				{
					nStringLenth=strData.GetLength();
					SetColWidth(lCol+1,nStringLenth+10);				
				}
			}

		}
		m_dataSource.m_pRS->MoveNext();
	}

	SetUserColAction(1);
	return m_nMaxRows;
}

void GetLanguageString(CString& strText,int nType,CString nParamcode)
{
	TCHAR szdata[100];
	if(nType==1)
	{
		if(0<GetPrivateProfileString(_T("AOI_PARAMETER"),nParamcode,NULL,szdata,sizeof(szdata),_T("C:\\TagLanguage_KOR.ini")))
		{
			strText.Format(_T("%s:: %s"),nParamcode,szdata);
		}
	}
	else if(nType==2)
	{
		if(0<GetPrivateProfileString(_T("VRS_PARAMETER"),nParamcode,NULL,szdata,sizeof(szdata),_T("C:\\TagLanguage_KOR.ini")))
		{
			strText.Format(_T("%s:: %s"),nParamcode,szdata);
		}
	}
	else if(nType==3)
	{
		if(nParamcode.Find(_T("1"),0)>=0)
			strText=_T("Active");
		else
			strText=_T("Deactive");

		/*
		if(0<GetPrivateProfileString("AOI_ALARAM",nParamcode,NULL,szdata,sizeof(szdata),"C:\\TagLanguage_KOR.ini"))
		{
			strText.Format("%s:: %s",nParamcode,szdata);
		}
		*/
	}
	else if(nType==4)
	{
		/*
		if(0<GetPrivateProfileString("VRS_ALARM",nParamcode,NULL,szdata,sizeof(szdata),"C:\\TagLanguage_KOR.ini"))
		{
			strText.Format("%s:: %s",nParamcode,szdata);
		}
*/
		if(nParamcode.Find(_T("1"),0)>=0)
			strText=_T("Active");
		else
			strText=_T("Deactive");
	}
	else if(nType==5)
	{
		
		CString strcode;

		if(nParamcode==_T("01"))
		{
			strcode=_T("1");
			strText=_T("Lot First");
		}
		else if(nParamcode==_T("02"))
		{
			strcode=_T("2");
			strText=_T("Lot End");
		}
		else if(nParamcode==_T("10"))
		{
			strcode=_T("3");
			strText=_T("Insert");
		}
		else if(nParamcode==_T("20"))
		{
			strcode=_T("4");
			strText=_T("Eject");
		}

/*
		if(0<GetPrivateProfileString("EVENT",strcode,NULL,szdata,sizeof(szdata),"C:\\TagLanguage_KOR.ini"))
		{
			strText.Format("%s:: %s",nParamcode,szdata);
		}
*/
	}
}


struct EES_PARAM_CODE
{
	CString strID;
	CString strType;
	CString strCode;
	CString strLayer;
	CString strLot;
	CString strSerial;
	int nCount;
	CString strParamData;
	CString strDTSignal;
	CString strTrans;
	CString strParamText[168];

	void Decode(CSpreadSheetEx* pSpread,int& nRow)
	{
		CString strData;
		TCHAR* pData;
		BOOL bFirst=TRUE;
		int nID=1;
		CString strIndex;
		while(nCount>0)
		{
			nCount--;

			if(bFirst==TRUE)
			{
				bFirst=FALSE;
				pData = _tcstok(strParamData.GetBuffer(strParamData.GetLength()),_T(","));
			}
			else
				pData = _tcstok(NULL,_T(","));
			
			strData.Format(_T("%s"),pData);

			pSpread->SetRow(nRow++);

			pSpread->SetCol(1);
			pSpread->SetText(strID);

			pSpread->SetCol(2);
			pSpread->SetText(strType);

			pSpread->SetCol(3);
			pSpread->SetText(strCode);

			pSpread->SetCol(4);
			pSpread->SetText(strLayer);
			
			pSpread->SetCol(5);
			pSpread->SetText(strLot);

			pSpread->SetCol(6);
			pSpread->SetText(strSerial);

			pSpread->SetCol(7);
			strIndex.Format(_T("%d"),nID);
			pSpread->SetText(strIndex);

			pSpread->SetCol(8);
			pSpread->SetText(strParamText[nID-1]);

			pSpread->SetCol(9);
			pSpread->SetText(strData);
	
			pSpread->SetCol(10);
			pSpread->SetText(strDTSignal);

			pSpread->SetCol(11);
			pSpread->SetText(strTrans);
			
			nID++;
		}
		strParamData.ReleaseBuffer();
	}
};


EES_PARAM_CODE g_ParamDecode;

int CSpreadSheetEx::InitOverEES(CString strQuery,CString strTableName,long lActiveRow,BOOL bSheetLock)
{	
	_variant_t vValue;
	_variant_t vColName;
	CString strData;
	CString strTitle;
	int nStringLenth=10;
	
	Reset();	
	m_dataSource.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strPassword);

	SetFontSize(13);
	SetRowHeight(-1,16);
	
	int nMode=1;
	if(strTableName=="PARAM")
	{
		nMode=2;
		
		CString strLangQ;

		strLangQ.Format(_T("select LANG_ENG FROM PARAM_LANGUAGE ORDER BY ID_CODE"));

		if(!m_dataSource.ExecuteQuery((LPCTSTR)strLangQ))
		{
			CString strMsg;
			strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
			Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
			return 0;
		}

		m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
		m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();
		for(long lRow=0; lRow<m_nMaxRows; lRow++)
		{			
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(0)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strParamText[lRow]=vValue.bstrVal;
			}
			m_dataSource.m_pRS->MoveNext();
		}
	}
	else if(strTableName=="ALARM")
	{ 
		nMode=3;	
	}
	else if(strTableName=="EVENT")
	{
		nMode=4;
	}
	else if(strTableName=="STAT")
	{
		nMode=5;
	}
	else if(strTableName=="LOSS")
	{
		nMode=6;
	}

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return 0;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();

	SetMaxCols(m_nMaxCols);
	SetMaxRows(m_nMaxRows);
	SetTypeHAlign(SSS_ALIGN_CENTER);
//	SetAutoCalc(TRUE);
	if(m_nMaxRows>0)
	{	
		m_dataSource.m_pRS->MoveFirst();
	}
	SetRow(0);
	int nType=0;
	for(long lCol=0; lCol<m_nMaxCols; lCol++)
	{
		SetCol(lCol+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();
		if(vColName.vt != VT_NULL)
		{
			vColName.ChangeType(VT_BSTR);
			strTitle = vColName.bstrVal ;
			SetText(strTitle);
			
			nStringLenth=strTitle.GetLength();
		}
		SetColWidth(lCol+1,nStringLenth+2);
	}

	if(m_nMaxRows>0)
	{
		m_dataSource.m_pRS->MoveFirst();
	}
	
	if(nMode==5)
	{
		SetCol(1);
		SetRow(0);
		SetText(_T("Index"));
		
		SetCol(2);
		SetText(_T("Type"));
		
		SetCol(3);
		SetText(_T("EQ Code"));
		
		SetCol(4);
		SetText(_T("Status Sign"));
		
		SetCol(5);
		SetText(_T("Run(1)/Idle(0)"));
		
		SetCol(6);
		SetText(_T("Time"));
		
		SetCol(7);
		SetText(_T("Send"));

		SetColWidth(6,25);
	}
	else if(nMode==6)
	{
		SetCol(1);
		SetRow(0);
		SetText(_T("Index"));
		
		SetCol(2);
		SetText(_T("Type"));
		
		SetCol(3);
		SetText(_T("EQ Code"));
		
		SetCol(4);
		SetText(_T("Name"));
		
		SetCol(5);
		SetText(_T("MDM"));
		
		SetCol(6);
		SetText(_T("Lot"));
		
		SetCol(7);
		SetText(_T("User"));
		
		SetCol(8);
		SetText(_T("Time"));
		
		SetCol(9);
		SetText(_T("Send"));

		SetColWidth(4,30);
		SetColWidth(5,20);
		SetColWidth(6,25);
		SetColWidth(8,25);
	}
	else if(nMode==4)
	{
		SetCol(1);
		SetRow(0);
		SetText(_T("Index"));
		
		SetCol(2);
		SetText(_T("Type"));
		
		SetCol(3);
		SetText(_T("EQ Code"));
		
		SetCol(4);
		SetText(_T("Code"));

		SetCol(5);
		SetText(_T("Model"));
		
		SetCol(6);
		SetText(_T("Layer"));

		SetCol(7);
		SetText(_T("Lot"));
		
		SetCol(8);
		SetText(_T("Serial"));
		
		SetCol(9);
		SetText(_T("Time"));
		
		SetCol(10);
		SetText(_T("Send"));

		SetCol(11);
		SetText(_T("Reserved"));
		
		SetColWidth(5,25);
		SetColWidth(6,25);
		SetColWidth(7,25);
		SetColWidth(9,25);
	}
	else if(nMode==3)
	{
		SetCol(1);
		SetRow(0);
		SetText(_T("Index"));
		
		SetCol(2);
		SetText(_T("Type"));
		
		SetCol(3);
		SetText(_T("EQ Code"));
		
		SetCol(4);
		SetText(_T("Name"));
		
		SetCol(5);
		SetText(_T("Status"));
		
		SetCol(6);
		SetText(_T("Lot"));
		
		SetCol(7);
		SetText(_T("Serial"));
		
		SetCol(8);
		SetText(_T("Time"));
		
		SetCol(11);
		SetText(_T("Send"));
		
		SetColWidth(4,25);
		SetColWidth(6,20);
	}
	else if(nMode==2)
	{
		SetMaxRows(m_nMaxRows*158);
		SetMaxCols(11);
		int nSheetRow=1;

		SetCol(1);
		SetRow(0);
		SetText(_T("Index"));

		SetCol(2);
		SetText(_T("Type"));

		SetCol(3);
		SetText(_T("EQ Code"));

		SetCol(4);
		SetText(_T("Layer Name"));

		SetCol(5);
		SetText(_T("Lot Name"));
		
		SetCol(6);
		SetText(_T("Serial"));

		SetCol(7);
		SetText(_T("Param ID"));

		SetCol(8);
		SetText(_T("Name"));

		SetCol(9);
		SetText(_T("Data"));

		SetCol(10);
		SetText(_T("Time"));

		SetCol(11);
		SetText(_T("Send"));


		SetColWidth(4,20);
		SetColWidth(5,20);
		SetColWidth(8,25);
		SetColWidth(9,20);
		SetColWidth(10,25);

		for(long lRow=0; lRow<m_nMaxRows; lRow++)
		{			
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(0)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strID=vValue.bstrVal;
			}

			vValue = m_dataSource.m_pRS->Fields->Item[(long)(1)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strType=vValue.bstrVal;
			}	
			
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(2)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strCode=vValue.bstrVal;
			}

			vValue = m_dataSource.m_pRS->Fields->Item[(long)(3)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strLot=vValue.bstrVal;
			}

			
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(4)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strLayer=vValue.bstrVal;
			}

			vValue = m_dataSource.m_pRS->Fields->Item[(long)(5)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strSerial=vValue.bstrVal;
			}

			vValue = m_dataSource.m_pRS->Fields->Item[(long)(6)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_INT);
				g_ParamDecode.nCount=vValue.intVal;
			}
			else
				g_ParamDecode.nCount=0;

			vValue = m_dataSource.m_pRS->Fields->Item[(long)(7)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strParamData=vValue.bstrVal;
			}

			vValue = m_dataSource.m_pRS->Fields->Item[(long)(8)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strDTSignal=vValue.bstrVal;
			}

			vValue = m_dataSource.m_pRS->Fields->Item[(long)(9)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				g_ParamDecode.strTrans=vValue.bstrVal;
			}

			g_ParamDecode.Decode(this,nSheetRow);

			m_dataSource.m_pRS->MoveNext();
		}
		
		return nSheetRow;
	}

	

	for(long lRow=0; lRow<m_nMaxRows; lRow++)
	{
		SetRow(lRow+1);
		
		for(long lCol=0; lCol<m_nMaxCols; lCol++)
		{
			SetCol(lCol+1);				
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
				if(nMode==2)
				{
					if(lCol==1 && strData=="AOI")
					{
						nType=1;
						SetText(strData);
					}
					else if(lCol==1 && strData=="VRS")
					{
						nType=2;
						SetText(strData);
					}
					else
					{
						SetText(strData);
					}
/*
					if(lCol==4)
					{
						CString strConv;
						//SetTypeHAlign(ALIGN_LEFT);
						//SetColWidth(5,26);
						//SetColWidth(6,25);
						GetLanguageString(strConv,nType,strData);
						SetText(strConv);
					}
*/
				}
				else if(nMode==3)
				{
					if(lCol==1 && strData=="AOI")
					{
						nType=3;
						SetText(strData);

						if(lRow==0)
						{
							nStringLenth=strData.GetLength();
							SetColWidth(lCol+1,nStringLenth+10);
						}
					}
					else if(lCol==1 && strData=="VRS")
					{
						nType=4;
						SetText(strData);

						if(lRow==0)
						{
							nStringLenth=strData.GetLength();
							SetColWidth(lCol+1,nStringLenth+10);
						}
					}
					else
					{
						SetText(strData);

						if(lRow==0)
						{
							nStringLenth=strData.GetLength();
							SetColWidth(lCol+1,nStringLenth+10);
						}
					}

					if(lCol==4)
					{
						CString strConv;
						//SetTypeHAlign(ALIGN_LEFT);
						//SetColWidth(5,26);
						//SetColWidth(6,25);
						GetLanguageString(strConv,nType,strData);
						SetText(strConv);	

						if(lRow==0)
						{
							nStringLenth=strData.GetLength();
							SetColWidth(lCol+1,nStringLenth+10);
						}
					}

				}
				else if(nMode==4)
				{
					if(lCol==3)
					{
						CString strConv;
						//SetTypeHAlign(ALIGN_LEFT);
					//	SetColWidth(5,26);
						//SetColWidth(6,25);
						GetLanguageString(strConv,5,strData);
						SetText(strConv);
						
					//	if(lRow==0)
					//	{
					//		nStringLenth=strData.GetLength();
					//		SetColWidth(lCol+1,nStringLenth+5);
					//	}
					}
					else
					{
						SetText(strData);
						if(lRow==0)
						{
							nStringLenth=strData.GetLength();
							SetColWidth(lCol+1,nStringLenth+10);
						}
					}
				}
				else
				{
					SetText(strData);	
					if(lRow==0)
					{
						nStringLenth=strData.GetLength();
						SetColWidth(lCol+1,nStringLenth+10);
					}
				}
				

				

			}
		}
		m_dataSource.m_pRS->MoveNext();
	}
	return m_nMaxRows;
}

int CSpreadSheetEx::InitOverEquipment(CString strQuery,float *ReturnData1,float *ReturnData2,float *ReturnData3)
{
	_variant_t vValue;
	_variant_t vColName;
	CString strData;
	CString strTitle;
	int nStringLenth=10;

	Reset();
	
	SetFontSize(13);
	SetRowHeight(-1,16);

	m_dataSource.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strPassword);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return 0;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();

	SetMaxCols(m_nMaxCols);
	SetMaxRows(m_nMaxRows);
	SetTypeHAlign(SSS_ALIGN_CENTER);

	if(m_nMaxRows>0)
	{	
		m_dataSource.m_pRS->MoveFirst();
	}
	SetRow(0);
	for(long lCol=0; lCol<m_nMaxCols; lCol++)
	{
		SetCol(lCol+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();
		if(vColName.vt != VT_NULL)
		{
			vColName.ChangeType(VT_BSTR);
			strTitle = vColName.bstrVal ;
			SetText(strTitle);

			nStringLenth=strTitle.GetLength();
		}
		else
		{
			//MessageBox(_T("NOTVAL"),_T("CAUTION"),MB_OK);
		}
		SetColWidth(lCol+1,nStringLenth+5);
	}

	if(m_nMaxRows>0)
	{
		m_dataSource.m_pRS->MoveFirst();
	}
	for(long lRow=0; lRow<m_nMaxRows; lRow++)
	{
		SetRow(lRow+1);
		
		for(long lCol=0; lCol<m_nMaxCols; lCol++)
		{
			SetCol(lCol+1);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
				SetText(strData);
				
				if(lCol==9)
				{
					ReturnData1[lRow]=_tcstod(strData, NULL);
					SetBackColor(RGB(250,234,122));
					
				}
				else if(lCol==10)
				{
					ReturnData2[lRow]=_tcstod(strData, NULL);
					SetBackColor(RGB(120,252,120));
				}
				else if(lCol==11)
				{
					ReturnData3[lRow]=_tcstod(strData, NULL);
					SetBackColor(RGB(252,206,120));
				}
			}
			else
			{
				if(lCol==9)
				{
					ReturnData1[lRow]=0;
					SetBackColor(RGB(250,234,122));
					
				}
				else if(lCol==10)
				{
					ReturnData2[lRow]=0;
					SetBackColor(RGB(120,252,120));
				}
				else if(lCol==11)
				{
					ReturnData3[lRow]=0;
					SetBackColor(RGB(252,206,120));
				}
			}
		}
		m_dataSource.m_pRS->MoveNext();
	}
	return m_nMaxRows;
}

int CSpreadSheetEx::GetReportMachineInfo_ALL(CString strQuery,CStringArray &EquipmentNames, CString strTableName,long lActiveRow,BOOL bSheetLock)
{
	_variant_t vValue;
	_variant_t vColName;
	CString strData;
	CString strTitle;
	CString strGetcurrentT;
	CString txt;
	int nStringLenth=10;

	EquipmentNames.RemoveAll();


	Reset();	
	m_dataSource.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strPassword);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return 0;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();

	SetMaxCols(m_nMaxCols);
	SetMaxRows(m_nMaxRows+10);
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetRowHeight(-1,15);
	SetGridType(SS_GRID_SOLID );

	SS_CELLTYPE stype;
	SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER ,128,SS_CHRSET_CHR,SS_CASE_NOCASE);
	HFONT hfont=CreateFont(12,0,0,0,FW_NORMAL,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(SS_ALLCOLS,SS_ALLROWS,hfont,TRUE);

	////////////////////////////////////////////0.Title
	SetRow(1);
	SetCol(1);
	SetRowHeight(1,50);
	AddCellSpan(1,1,m_nMaxCols,1);	

	hfont=CreateFont(25,0,0,0,FW_BOLD,0,TRUE,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(1,1,hfont,TRUE);
	SetCellType(1, 1, &stype);
	SetText(_T("Machine Infomation"));
	///////////////////////////////////////////////1.Name & Time
	SetRow(2);
	SetCol(1);
	CTime time=CTime::GetCurrentTime();
	strGetcurrentT.Format(_T("%d/%02d/%02d"),time.GetYear(),time.GetMonth(),time.GetDay());
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	SetText(strGetcurrentT);

	SetCol(m_nMaxCols-1);
//	SetFontUnderline(1);
//	SetTypeHAlign(ALIGN_LEFT);
//	SetTypeVAlign(SSS_ALIGN_VCENTER);
	SetText(_T("Manager:"));

	SetCol(m_nMaxCols);

	hfont=CreateFont(16,0,0,0,FW_NORMAL,0,TRUE,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(m_nMaxCols,2,hfont,TRUE);

	//SetFontUnderline(1);
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	SetText(_T("        "));


	//////////////////////////////////////////////2.subTitle
	SetCol(1);
	SetRow(4);
	SetRowHeight(4,30);
	AddCellSpan(1,4,3,1);
	SetFontSize(16);
	SetFontBold(1);
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetTypeVAlign(SSS_ALIGN_VCENTER);

	hfont=CreateFont(20,0,0,0,FW_BOLD,0,TRUE,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(1,4,hfont,TRUE);
	SetText(_T("1.Equipment Condition"));


	//////////////////////////////////////////3.Title
	if(m_nMaxRows>0)
	{	
		m_dataSource.m_pRS->MoveFirst();
	}
	SetRow(5);
	hfont=CreateFont(12,0,0,0,FW_BOLD,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(SS_ALLCOLS,5,hfont,TRUE);
	for(long lCol=0; lCol<m_nMaxCols; lCol++)
	{
		SetCol(lCol+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();
		if(vColName.vt != VT_NULL)
		{
			SetGridType(SS_GRID_SOLID );
			//SetGridSolid(1);
			vColName.ChangeType(VT_BSTR);
			strTitle = vColName.bstrVal ;		
			SetBackColor(RGB(200,200,200));	
			SetCellType(lCol+1, 5, &stype);
			SetText(strTitle);

			nStringLenth=strTitle.GetLength();
		}
		if(lCol==2 || lCol==3 ||lCol==4  || lCol==6 ||  lCol==8)
		{
			SetColWidth(lCol+1,nStringLenth+1);
		}
		else
		{
			SetColWidth(lCol+1,nStringLenth+9);
		}
	}

	////////////////////////////////////////4.value
	if(m_nMaxRows>0)
	{
		m_dataSource.m_pRS->MoveFirst();
	}
	for(long lRow=0; lRow<m_nMaxRows; lRow++)
	{
		SetRow(lRow+6);
		
		for(long lCol=0; lCol<m_nMaxCols; lCol++)
		{
			SetCol(lCol+1);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;				
				SetCellType(lCol+1, lRow+6, &stype);
				SetText(strData);
				if(lCol==0)
				{
					EquipmentNames.Add(strData);
				}
			}
		}
		m_dataSource.m_pRS->MoveNext();
	}

	return m_nMaxRows;
}

int CSpreadSheetEx::GetReportMachineInfo_Status(CString strQuery,int nEquipmentCnt,CString strEquipCode, CString strTableName,long lActiveRow,BOOL bSheetLock)
{
	_variant_t vValue;
	_variant_t vColName;
	CString strData;
	CString strTitle;
	CString strGetcurrentT;
	CString txt;
	int nStringLenth=10;
	CStringArray m_StrarryMachine;
	int nRowNum=nEquipmentCnt+10;
	int nColrecnt=0;

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);

		return 0;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();
	SS_CELLTYPE stype;
	SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER ,128,SS_CHRSET_CHR,SS_CASE_NOCASE);
	

	SetMaxRows(m_nMaxRows+nRowNum+20);


	SetCol(1);
	SetRow(nRowNum);
	SetRowHeight(nRowNum,30);
	AddCellSpan(1,nRowNum,3,1);
	SetFontSize(16);
	SetFontBold(1);
	SetTypeHAlign(ALIGN_LEFT);
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	HFONT hfont=CreateFont(20,0,0,0,FW_BOLD,0,TRUE,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(1,nRowNum,hfont,TRUE);
	txt.Format(_T("-%s Status"),strEquipCode);
	SetText(txt);
	/////////////////////////
	//////////////////////////////////////////3.Title
	if(m_nMaxRows>0)
	{	
		m_dataSource.m_pRS->MoveFirst();
	}
	SetRow(nRowNum+1);
	hfont=CreateFont(12,0,0,0,FW_BOLD,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(SS_ALLCOLS,nRowNum+1,hfont,TRUE);
	SetCellType(SS_ALLCOLS,nRowNum+1, &stype);
	for(long lCol=0; lCol<m_nMaxCols; lCol++)
	{
		SetCol(nColrecnt+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();
		if(vColName.vt != VT_NULL && lCol!=0 && lCol!=6 && lCol!=7)
		{
			if(lCol==9)
			{
				SetCol(nColrecnt+2);
			}
			SetGridType(SS_GRID_SOLID );
			//SetGridSolid(1);
			vColName.ChangeType(VT_BSTR);
			strTitle = vColName.bstrVal ;

			if(lCol==5)
				strTitle="Defect";

			if(lCol==6 ||lCol==7 || lCol==8)
				strTitle="Start";

		//	if(lCol==7)
			//	strTitle="Start";

			if(lCol==9)
				strTitle="End";
		
			SetBackColor(RGB(200,200,200));	
			
			SetText(strTitle);

//			nStringLenth=strTitle.GetLength();
			nColrecnt++;
		}
//		SetColWidth(lCol+1,nStringLenth+9);
	}
	AddCellSpan(6,nRowNum+1,2,1);
	AddCellSpan(8,nRowNum+1,2,1);

	////////////////////////////////////////4.value

	if(m_nMaxRows>0)
	{
		m_dataSource.m_pRS->MoveFirst();
	}
	for(long lRow=0; lRow<m_nMaxRows; lRow++)
	{
		SetRow(nRowNum+2+lRow);
		
		nColrecnt=0;
		for(long lCol=0; lCol<m_nMaxCols; lCol++)
		{
			SetCol(nColrecnt+1);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol)]->Value;
			if(vValue.vt != VT_NULL && lCol!=0 &&  lCol!=6 && lCol!=7)
			{
				if(lCol==9)
				{
					SetCol(nColrecnt+2);
					SetCellType(nColrecnt+2, nRowNum+2+lRow, &stype);
				}
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
				SetTypeVAlign(SSS_ALIGN_VCENTER);
				SetCellType(nColrecnt+1, nRowNum+2+lRow, &stype);
				SetText(strData);
				nColrecnt++;
			}
		}
		m_dataSource.m_pRS->MoveNext();
	}
//	AddCellSpan(7,nRowNum+2,2,1);
//	AddCellSpan(9,nRowNum+2,2,1);

	return m_nMaxRows;
}
/*
int CSpreadSheetEx::GetReportInspectionResult(CString strQuery1,CString strQuery2)
{
	CUploadDefect Upload;
	Upload.SetDefLangTable();
	Upload.SetBadLangTable();
	stDefectLang* stLang;

	_variant_t vValue;
	_variant_t vColName;
	CString strData;
	CString strTitle;
	CString strGetcurrentT;
	CString txt;
	int nStringLenth=10;
	CStringArray m_StrarryMachine;
	int nColrecnt=0;
	int nQuery2Row,nQuery2Col;

	Reset();	
	m_dataSource.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strPassword);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery1))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);

		return 0;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();

	SS_CELLTYPE stype;
	SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER ,128,SS_CHRSET_CHR,SS_CASE_NOCASE);
	
	
	SetMaxCols(m_nMaxCols);
	SetMaxRows(m_nMaxRows+10);
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetRowHeight(-1,15);
	SetGridType(SS_GRID_SOLID );
	//SetGridShowHoriz(FALSE);
	//SetGridShowVert(FALSE);
	HFONT hfont;	
	hfont=CreateFont(16,0,0,0,FW_NORMAL,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(SS_ALLCOLS,SS_ALLROWS,hfont,TRUE);
	
	////////////////////////////////////////////0.Title
	SetRow(1);
	SetCol(1);
	SetRowHeight(1,50);
	AddCellSpan(1,1,m_nMaxCols,1);
	hfont=CreateFont(25,0,0,0,FW_BOLD,0,TRUE,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(1,1,hfont,TRUE);	
	SetCellType(1, 1, &stype);
	SetText(_T("Inspection Result"));
	///////////////////////////////////////////////1.Name & Time
	
	SetRow(2);
	SetCol(1);
	CTime time=CTime::GetCurrentTime();
	strGetcurrentT.Format(_T("%d/%02d/%02d"),time.GetYear(),time.GetMonth(),time.GetDay());
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	SetText(strGetcurrentT);
	SetCol(m_nMaxCols-1);
	SetText(_T("Manager:"));
	SetCol(m_nMaxCols);
	//SetFontUnderline(1);
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	hfont=CreateFont(16,0,0,0,FW_BOLD,0,TRUE,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(m_nMaxCols,2,hfont,TRUE);	
	SetText(_T("              "));

	//////////////////////////////////////////////2.subTitle
	SetCol(1);
	SetRow(4);
	SetRowHeight(4,30);
	AddCellSpan(1,4,3,1);
	hfont=CreateFont(20,0,0,0,FW_BOLD,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(1,4,hfont,TRUE);	
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	SetText(_T("1.Defect Info"));


	if(m_nMaxRows>0)
	{	
		m_dataSource.m_pRS->MoveFirst();
	}
	hfont=CreateFont(16,0,0,0,FW_BOLD,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(SS_ALLCOLS,5,hfont,TRUE);
	SetRow(5);
	long lCol=0;
	for(lCol=0; lCol<m_nMaxCols; lCol++)
	{
		SetCol(lCol+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();
		if(vColName.vt != VT_NULL)
		{

			SetGridType(SS_GRID_SOLID);
			vColName.ChangeType(VT_BSTR);
			strTitle = vColName.bstrVal ;
			if(lCol==6)
			{
				switch(LANGUAGE_KOREAN)
				{
				case LANGUAGE_KOREAN:
					strTitle="KOR_NAME";
					break;

				case LANGUAGE_ENGLISH:
					strTitle="ENG_NAME";
					break;

				case LANGUAGE_JAPANESE:
					strTitle="JPN_NAME";
					break;

				case LANGUAGE_CHINESE:
					strTitle="CHINESE_NAME";
					break;
				}				
			}
			SetFontBold(1);
			SetBackColor(RGB(200,200,200));			
			SetCellType(lCol+1,5, &stype);
			SetText(strTitle);

			nStringLenth=strTitle.GetLength();
		}

		if(lCol==0 || lCol==3 || lCol==4 ||lCol==7)
		{
			SetColWidth(lCol+1,nStringLenth+5);
		}
		else
		{
			SetColWidth(lCol+1,nStringLenth+15);
		}
	}
	/////////////////////////////////////////////////title

	CString strbadlang;
	////////////////////////////////////////4.value
	if(m_nMaxRows>0)
	{
		m_dataSource.m_pRS->MoveFirst();
	}
	long lRow=0; 
	lCol=0;


	for( lRow=0; lRow<m_nMaxRows; lRow++)
	{
		SetRow(lRow+6);
		strbadlang="";
		for( lCol=0; lCol<m_nMaxCols; lCol++)
		{
			SetCol(lCol+1);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;
				if(lCol==5)
				{
					if(Upload.m_LanguageTable.Lookup(_ttoi(strData),stLang))
					{
						strbadlang=stLang->GetLang(pGlobalFrame->m_bLanguage);
					}
				}
				else if(lCol==6)
				{
					if(strbadlang=="")
						strData = vValue.bstrVal ;
					else
						strData=strbadlang;
				}
				SetTypeVAlign(SSS_ALIGN_VCENTER);
				SetCellType(lCol+1,lRow+6, &stype);
				SetText(strData);
			}
		}
		m_dataSource.m_pRS->MoveNext();
	}

	/////////////////////////////////////////////////////////////////////////////////

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery2))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);

		return 0;
	}
	nQuery2Col = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	nQuery2Row = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();

	//	SetMaxCols(m_nMaxCols);
	SetMaxRows(m_nMaxRows+nQuery2Row+10);


	//////////////////////////////////////////////2.subTitle
	SetCol(1);
	SetRow(m_nMaxRows+7);
	SetRowHeight(m_nMaxRows+7,30);
	AddCellSpan(1,m_nMaxRows+7,3,1);
	SetFontSize(16);
	SetFontBold(1);
	SetTypeHAlign(ALIGN_LEFT);
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	hfont=CreateFont(16,0,0,0,FW_BOLD,0,TRUE,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(1, m_nMaxRows+7,hfont,TRUE);	
	SetText(_T("2.Bad Info"));

	if(nQuery2Row>0)
	{	
		m_dataSource.m_pRS->MoveFirst();
	}
	SetRow(m_nMaxRows+8);
	hfont=CreateFont(16,0,0,0,FW_BOLD,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	SetFont(SS_ALLCOLS,m_nMaxRows+8,hfont,TRUE);
	for(lCol=0; lCol<m_nMaxCols; lCol++)
	{
		SetCol(lCol+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)lCol]->GetName();

		if(vColName.vt != VT_NULL)
		{
			SetGridType(SS_GRID_SOLID);
			//SetGridSolid(1);
			vColName.ChangeType(VT_BSTR);
			strTitle = vColName.bstrVal;

			if(lCol==6)
			{
				switch(pGlobalFrame->m_bLanguage)
				{
				case LANGUAGE_KOREAN:
					strTitle="KOR_NAME";
					break;

				case LANGUAGE_ENGLISH:
					strTitle="ENG_NAME";
					break;

				case LANGUAGE_JAPANESE:
					strTitle="JPN_NAME";
					break;

				case LANGUAGE_CHINESE:
					strTitle="CHINESE_NAME";
					break;
				}				
			}

			SetFontBold(1);
			SetBackColor(RGB(200,200,200));		
			SetCellType(lCol+1,m_nMaxRows+8, &stype);
			SetText(strTitle);

			//	nStringLenth=strTitle.GetLength();
		}
		//	SetColWidth(lCol+1,nStringLenth+9);
	}

	////////////////////////////////////////4.value
	if(nQuery2Row>0)
	{
		m_dataSource.m_pRS->MoveFirst();
	}


	for(lRow=0; lRow<nQuery2Row; lRow++)
	{
		SetRow(lRow+m_nMaxRows+9);

		for( lCol=0; lCol<m_nMaxCols; lCol++)
		{
			SetCol(lCol+1);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol)]->Value;
			if(vValue.vt != VT_NULL)
			{
				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal ;

				if(lCol==5)
				{
					if(Upload.m_LanguageTableBad.Lookup(_ttoi(strData),stLang))
					{
						strbadlang=stLang->GetLang(pGlobalFrame->m_bLanguage);
					}
				}
				else if(lCol==6)
				{
					if(strbadlang=="")
						strData = vValue.bstrVal ;
					else
						strData=strbadlang;
				}
				SetTypeVAlign(SSS_ALIGN_VCENTER);
				SetCellType(lCol+1,lRow+m_nMaxRows+9, &stype);
				SetText(strData);
			}
		}
		m_dataSource.m_pRS->MoveNext();
	}

	return m_nMaxRows;
}
*/
void CSpreadSheetEx::SetSpreadInfo(double dHeadHeight, COLORREF rgbHeadBg, COLORREF rgbHeadFont, double dHeadFontSize, BOOL bHeadFontBold, double dCellWidth, double dCellHeight, int nCellFontSize, BOOL bCellFontBold, COLORREF rgbCellBg, COLORREF rgbCellFont, long nCellHAlign, long nCellVAlign, CString strFontName)
{
	m_dHeadHeight = dHeadHeight;
	m_rgbHeadBg = rgbHeadBg;
	m_rgbHeadFont = rgbHeadFont;
	m_dHeadFontSize = dHeadFontSize;
	m_bHeadFontBold = bHeadFontBold;
	m_strFontName = strFontName;

	m_dCellWidth = dCellWidth;
	m_dCellHeight = dCellHeight;
	m_nCellFontSize = nCellFontSize;
	m_bCellFontBold = bCellFontBold;
	m_rgbCellBg = rgbCellBg;
	m_rgbCellFont = rgbCellFont;
	m_nCellHAlign = nCellHAlign;
	m_nCellVAlign = nCellVAlign;
	sttField dbField;
	CString strName;
	_variant_t vColName;
	
	// Get Field name

	CString strFieldId;
	int nIndex = 0;
	for(int nField=0; nField<m_nMaxCols; nField++)
	{
		strFieldId = m_FieldIDList.GetAt(nField);
		dbField.strFieldId =  strFieldId;
		dbField.bExtField = CheckExtReferedField(strFieldId);
		dbField.bPrimary = CheckPrimaryKeyField(strFieldId);
		dbField.bCombo = CheckComboStyleField(strFieldId);
		dbField.bColor = CheckColorStyleField(strFieldId);
		dbField.strAlias=strFieldId;
	//	GetAliasFieldNameOnDb(strFieldId,dbField.strAlias);
		if(dbField.bCombo)
		{
			dbField.StrComboData = GetComboDataList(strFieldId);
		}
		pArField->Add(dbField);
	}

	m_arDeletedRecord.RemoveAll();
	m_arModifiedRecord.RemoveAll();
	m_arAppendedRecord.RemoveAll();
}

BOOL CSpreadSheetEx::CheckExtReferedField(CString strFieldId)
{
	CString strOrgFieldId;
	for(long i = 0 ; i < m_OrgFieldIdList.GetSize() ; i++)
	{
		strOrgFieldId = m_OrgFieldIdList.GetAt(i);
		if(strOrgFieldId == strFieldId)
		{
			return FALSE;
		}
	}
	return TRUE;
}

long CSpreadSheetEx::GetColorFieldIndex()
{
	long nField=0;
	CString strFieldId,strColorStyleId;

	for(nField=0; nField<m_nMaxCols; nField++)
	{
		strFieldId = m_FieldIDList.GetAt(nField);
		if(CheckColorStyleField(strFieldId))
			break;
	}

	return nField;
}

BOOL CSpreadSheetEx::CheckPrimaryKeyField(CString strFieldId)
{
	CString strPrimaryFieldId;
	for(long i = 0 ; i < m_PrimaryKeyList.GetSize() ; i++)
	{
		strPrimaryFieldId = m_PrimaryKeyList.GetAt(i);
		if(strPrimaryFieldId == strFieldId)
		{
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CSpreadSheetEx::CheckComboStyleField(CString strFieldId)
{
	CString strComboStyleId;
	for(long i = 0 ; i < m_ComboStyleFieldIdList.GetSize() ; i++)
	{
		strComboStyleId = m_ComboStyleFieldIdList.GetAt(i);
		if(strComboStyleId == strFieldId)
		{
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CSpreadSheetEx::CheckColorStyleField(CString strFieldId)
{
	CString strColorStyleId;
	for(long i = 0 ; i < m_ColorStyleFieldIdList.GetSize() ; i++)
	{
		strColorStyleId = m_ColorStyleFieldIdList.GetAt(i);
		if(strColorStyleId == strFieldId)
		{
			return TRUE;
		}
	}
	return FALSE;
}

void CSpreadSheetEx::FindColorFieldIdOnDb()
{
	CString strQuery;
	strQuery.Format(_T("SELECT COLUMN_NAME FROM COLUMNINFO WHERE TABLE_NAME = '%s' AND COLUMN_TYPE = '2'"),m_strTable);
	
	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at FindColorFieldIdOnDb()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	int nMaxRows = m_dataSource.m_pRS->RecordCount;

	m_ColorStyleFieldIdList.RemoveAll();
	
	if(nMaxRows > 0)
	{
		CString strData;
		_variant_t vValue; 

		m_dataSource.m_pRS->MoveFirst();
		for(long lRow=1; lRow<=nMaxRows; lRow++)
		{
			vValue = m_dataSource.m_pRS->Fields->Item[(long)0]->Value;
			vValue.ChangeType(VT_BSTR);
			strData = vValue.bstrVal ;
			m_ColorStyleFieldIdList.Add(strData);
			if(lRow < nMaxRows)
				m_dataSource.m_pRS->MoveNext();
		}
	}
}

// Check and return primary key field is null?
int CSpreadSheetEx::GetInvalidPrimaryKeyField(int nRecordNo)
{
	int i,nInvalidField=0;
	sttField dbField;
	CString strPrimaryKeyData;
	for(i = 0; i < GetMaxCols(); i++)
	{
		dbField = pArField->GetAt(i);
		if(dbField.bPrimary)
		{
			SetRow(nRecordNo);
			SetCol(i+1);
			strPrimaryKeyData = GetText();
			if(strPrimaryKeyData.GetLength()==0)
			{
				nInvalidField = i;
				return nInvalidField;
			}
		}
	}
	return nInvalidField;
}

int  CSpreadSheetEx::GetInvalidComboStyleField(int nRecordNo)
{
	int i,nInvalidField=0;
	sttField dbField;
	CString strComboData;
	for(i = 0; i < GetMaxCols(); i++)
	{
		dbField = pArField->GetAt(i);
		if(dbField.bCombo)
		{
			SetRow(nRecordNo);
			SetCol(i+1);
			strComboData = GetText();
			if(strComboData.GetLength()==0)
			{
				nInvalidField = i;
				return nInvalidField;
			}
		}
	}
	return nInvalidField;
}

int CSpreadSheetEx::GetPrimaryKeyData(int nRecordNo,CStringArray &strArray)
{
	int i,nPrimarykeyCount=0;
	int nCount = GetMaxCols();
	sttField dbField;
	CString strPrimaryKeyName= _T("");
	for(i = 0; i < nCount; i++)
	{
		dbField = pArField->GetAt(i);
		if(dbField.bPrimary)
		{
			SetRow(nRecordNo);
			SetCol(i+1);
			strPrimaryKeyName = GetText();
			strArray.Add(strPrimaryKeyName);
			nPrimarykeyCount++;
		}
	}
	return nPrimarykeyCount;
}

void CSpreadSheetEx::SetCellColor(long lCol, long lRow, COLORREF color)
{
	SetCol(lCol);	
	SetRow(lRow);
	SetBackColor(color);
}

void CSpreadSheetEx::SetCellText(long lCol, long lRow, CString strText)
{
	SetCol(lCol);	
	SetRow(lRow);	
	SetTypeVAlign(SSS_ALIGN_VCENTER);
	SetTypeHAlign(SSS_ALIGN_CENTER);
	SetText((LPCTSTR)strText);	
}
	
void CSpreadSheetEx::SetCellWidth(long lCol, double dSize)
{
	SetColWidth(lCol, dSize);
}

void CSpreadSheetEx::SetCellHeight(long lRow, double dSize)
{
	SetRowHeight(lRow, dSize);
}

void CSpreadSheetEx::SetFieldEditEnable(long lCol, BOOL bEnable)
{
	if(lCol < 0)
		return;

	SetCol(lCol);
	for(long lRow=1; lRow<=m_nMaxRows; lRow++)
	{
		SetRow(lRow);
		if(bEnable)
			SetCellType(TYPE_DEFAULT);
		else
			SetCellType(TYPE_STATIC);
	}
}

void CSpreadSheetEx::DispSelectRow(long lRow)
{
	int nField;
	sttField dbField;
	CString strData;

	SetReDraw(FALSE);

	if(lRow<GetMaxRows() || m_nOldSelectRow != lRow)
	{
		if(m_nOldSelectRow != 0 ) // Clear Color Line.
		{
			SetRow(m_nOldSelectRow);
			for(nField=1; nField<=m_nMaxCols; nField++)
			{
				SetCol(nField);

				SetBackColor(m_rgbCellBg);
				SetForeColor(m_rgbCellFont);
			}
		}

		m_nNewSelectRow = lRow;

		SetRow(m_nNewSelectRow);
		for(nField=1; nField<=m_nMaxCols; nField++)
		{
			SetCol(nField);
			SetBackColor(0x00C0E0FF); //Bright Orange
		}
	}
	else
	{
		lRow = GetMaxRows();
		m_nNewSelectRow = lRow;

		SetRow(m_nNewSelectRow);
		for(nField=1; nField<=m_nMaxCols; nField++)
		{
			SetCol(nField);

			SetBackColor(0x00C0E0FF); //Bright Orange
		}
	}
	SetReDraw(TRUE);

	m_nOldSelectRow = m_nNewSelectRow;
}

void CSpreadSheetEx::DispSelectRow(long lRow, long lCol)
{
	int nField;
	sttField dbField;
	CString strData;

	SetReDraw(FALSE);

	if (lRow < GetMaxRows() || m_nOldSelectRow != lRow)
	{
		if (m_nOldSelectRow != 0) // Clear Color Line.
		{
			SetRow(m_nOldSelectRow);
			for (nField = 1; nField <= m_nMaxCols; nField++)
			{
				SetCol(nField);

				SetBackColor(m_rgbCellBg);
				SetForeColor(m_rgbCellFont);
			}
		}

		m_nNewSelectRow = lRow;

		SetRow(m_nNewSelectRow);
		for (nField = lCol; nField <= m_nMaxCols; nField++)
		{
			SetCol(nField);
			SetBackColor(0x00C0E0FF); //Bright Orange
		}
	}
	else
	{
		lRow = GetMaxRows();
		m_nNewSelectRow = lRow;

		SetRow(m_nNewSelectRow);
		for (nField = 1; nField <= m_nMaxCols; nField++)
		{
			SetCol(nField);

			SetBackColor(0x00C0E0FF); //Bright Orange
		}
	}
	SetReDraw(TRUE);

	m_nOldSelectRow = m_nNewSelectRow;
}

void CSpreadSheetEx::DispSelectRowOverdefectEquipment(long lRow)
{
	int nField;
	sttField dbField;
	CString strData;

	SetReDraw(FALSE);

	if(lRow<GetMaxRows() || m_nOldSelectRow != lRow)
	{
		if(m_nOldSelectRow != 0 ) // Clear Color Line.
		{
			SetRow(m_nOldSelectRow);
			for(nField=1; nField<=m_nMaxCols; nField++)
			{
				
				SetCol(nField);

				if(nField=8)
				{
					SetBackColor(RGB(128,255,255));
				}
				else if(nField==9)
				{
					SetBackColor(RGB(255,128,255));
				}
				else if(nField==10)
				{
					SetBackColor(RGB(255,128,100));
				}
				else
				{
					SetBackColor(m_rgbCellBg);
					SetForeColor(m_rgbCellFont);
				}
			}
		}

		m_nNewSelectRow = lRow;

		SetRow(m_nNewSelectRow);
		for(nField=1; nField<=m_nMaxCols; nField++)
		{
			SetCol(nField);
			SetBackColor(0x00C0E0FF); //Bright Orange
		}
	}
	else
	{
		lRow = GetMaxRows();
		m_nNewSelectRow = lRow;

		SetRow(m_nNewSelectRow);
		for(nField=1; nField<=m_nMaxCols; nField++)
		{
			SetCol(nField);

			SetBackColor(0x00C0E0FF); //Bright Orange
		}
	}
	SetReDraw(TRUE);

	m_nOldSelectRow = m_nNewSelectRow;
}


void CSpreadSheetEx::DispSpread(BOOL Field)
{
	if(m_nMaxCols < 1)
	{
		AfxMessageBox(_T("No Field."));
		//AfxMessageBox("Field∞° æ¯Ω¿¥œ¥Ÿ.");
		return;
	}
	SetReDraw(FALSE);
	SetShadowColor(m_rgbHeadBg,m_rgbHeadFont, COLOR_3DSHADOW ,COLOR_3DHILIGHT);
	//SetShadowColor((unsigned long)m_rgbHeadBg);			
	//SetShadowText((unsigned long)m_rgbHeadFont);

	SS_CELLTYPE stype;
	SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER ,128,SS_CHRSET_CHR,SS_CASE_NOCASE);

	HFONT hfont;
	hfont = CreateFont(m_dHeadFontSize,0,0,0,m_bHeadFontBold,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));
	

	//SetFontBold((BOOL)m_bHeadFontBold);
	//SetFontSize((float)m_dHeadFontSize);
	//SetFontName((LPCTSTR)m_strFontName);

	SetMaxCols((long)m_nMaxCols);
	SetMaxRows((long)m_nMaxRows);
	SetTypeHAlign(SSS_ALIGN_CENTER);

	//Column Caption
	SetRow(0L);

	SetColWidth(0,10);

	SetRowHeight(-1, m_dHeadHeight); // COL(0)'S HEIGHT

	sttField dbField;

	int nLengh;
	for(long nField = 0; nField < pArField->GetSize(); nField++)
	{
		dbField = pArField->GetAt(nField);

		SetCol(nField+1);
		
		CString str = m_FieldIDList.GetAt(nField);
		nLengh=str.GetLength();
		SetColWidth(nField+1, (double)nLengh + 5);

		//SetColWidth(nField+1, m_dCellWidth);

		if(Field) SetText(dbField.strFieldId);
		else	SetText(dbField.strAlias);
	}

	// Row Data (Record)
	CString strData;
	_variant_t vValue; 
	int nStringLenth;
	
	if(GetMaxRows() < 1)
		return;
	
	if(!m_dataSource.ExecuteQuery(m_strQuery))
	{
		AfxMessageBox(_T("Error occur at m_dataSource.ExecuteQuery()"));
		return;
	}	
	SetRow(-1);
	SetCol(-1);
	SetFontSize(9);
	//hfont = CreateFont(13,0,0,0,m_bHeadFontBold,0,0,0,0,SHIFTJIS_CHARSET,0,0,0,_T("πŸ≈¡"));
		hfont = CreateFont(13,0,0,0,m_bHeadFontBold,0,0,0,0,SHIFTJIS_CHARSET,0,0,0,_T("MS PMincho"));
	SetFont(SS_ALLCOLS,SS_ALLROWS,hfont,TRUE);

	SetCellType(SS_ALLCOLS, SS_ALLROWS, &stype);
	m_dataSource.m_pRS->MoveFirst();
	for(long lRow=1; lRow<=m_nMaxRows; lRow++)
	{
		SetRow(lRow);
		
		for(long lCol=1; lCol<=m_nMaxCols; lCol++)
		{
			SetCol(lCol);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol-1)]->Value;

			if(vValue.vt != VT_NULL)
			{

				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal;
				
				SetText(strData);

				if(lRow==1)
				{
					
					nStringLenth=strData.GetLength();

					if(nStringLenth*2 > GetColWidth(lCol))
						SetColWidth(lCol,nStringLenth+10);
				}
		
				dbField = pArField->GetAt(lCol-1);

				if(m_bSheetLock)
				{
					SetLock(TRUE);
					//SetCellType(TYPE_STATIC);				
				}
				else
				{
					if(dbField.bCombo)
					{
						SetCellType(TYPE_COMBO);
						SetTypeComboBoxList (dbField.StrComboData);
						SetTypeComboBoxEditable(TRUE);
						SetTypeComboBoxMaxDrop(10);
					}

					if(dbField.bPrimary)
					{
						//SetCellType(TYPE_STATIC);
					}	
				}
			}			
		}
		m_dataSource.m_pRS->MoveNext();
	}
	SetReDraw(TRUE);
}


void CSpreadSheetEx::InitDisplayTable(CString strQuery,CString strTableName)
{
	//m_bSheetLock = bSheetLock;
	m_strQuery = strQuery;
	m_strTable = strTableName;

	m_dataSource.InitDB(m_strServerIp, m_strCatalog, m_strAccessID, m_strPassword);

	if(!m_dataSource.ExecuteQuery((LPCTSTR)strQuery))
	{
		CString strMsg;
		strMsg.Format(_T("Error occur at m_dataSource.ExecuteQuery() at Init() on CSpreadSheetEx()\r\n%s"),m_dataSource.GetLastError());
		Log(strMsg); 
		AfxMessageBox(strMsg,MB_ICONSTOP);
		return;
	}

	m_nMaxCols = m_dataSource.m_pRS->Fields->Count ; //    //-> m_dataSource.OnGetFieldCount();
	m_nMaxRows = m_dataSource.m_pRS->RecordCount; //  OnGetRecordCount();
	if(m_nMaxRows== -1)
	{
		AfxMessageBox(_T("No Record."));
		//AfxMessageBox("«ÿ¥Á«œ¥¬ ∑πƒ⁄µÂ∞° æ¯Ω¿¥œ¥Ÿ.");
		return;
	}
	//if(lActiveRow>m_nMaxRows)
	//	lActiveRow = m_nMaxRows;
	//if(lActiveRow==0)
	//	lActiveRow = 1;

	m_arModifiedRecord.RemoveAll();
	m_arAppendedRecord.RemoveAll();
	m_arInsertedRecord.RemoveAll();
	if(pArField->GetSize())
		pArField->RemoveAll();

	//FindFieldIdOnDb();
	//FindOrgFieldIdOnDb();

	//FindPkOnDb();

	////FindComboFieldIdOnDb();
	//FindComboDataQueryOnDb();

	//FindColorFieldIdOnDb();
	//SetSpreadInfo();
	DispSpreadTable();

	//DispSelectRow(lActiveRow);
	//SetActiveCell (1,lActiveRow);

	//SetUserColAction(1);
}

void CSpreadSheetEx::DispSpreadTable(BOOL Field)
{
	if(m_nMaxCols < 1)
	{
		AfxMessageBox(_T("No Field."));
		//AfxMessageBox("Field∞° æ¯Ω¿¥œ¥Ÿ.");
		return;
	}
	SetReDraw(FALSE);
	SetShadowColor(m_rgbHeadBg,m_rgbHeadFont, COLOR_3DSHADOW ,COLOR_3DHILIGHT);
	//SetShadowColor((unsigned long)m_rgbHeadBg);			
	//SetShadowText((unsigned long)m_rgbHeadFont);

	SS_CELLTYPE stype;
	SetTypeEdit(&stype, SSS_ALIGN_VCENTER | SSS_ALIGN_CENTER ,128,SS_CHRSET_CHR,SS_CASE_NOCASE);

	HFONT hfont;
	hfont = CreateFont(m_dHeadFontSize,0,0,0,m_bHeadFontBold,0,0,0,0,0,0,0,0,_T("πŸ≈¡"));


	//SetFontBold((BOOL)m_bHeadFontBold);
	//SetFontSize((float)m_dHeadFontSize);
	//SetFontName((LPCTSTR)m_strFontName);

	SetMaxCols((long)m_nMaxCols);
	SetMaxRows((long)m_nMaxRows);
	SetTypeHAlign(SSS_ALIGN_CENTER);

	//Column Caption
	SetRow(0L);

	SetColWidth(0,10);

	SetRowHeight(-1, m_dHeadHeight); // COL(0)'S HEIGHT

	
	_variant_t vColName;

	int nLengh;
	CString str;
	
	for(long nField = 0; nField < m_nMaxCols; nField++)
	{		
		SetCol(nField+1);
		vColName = m_dataSource.m_pRS->Fields->Item[(long)nField]->GetName();
		str = vColName.bstrVal;
		nLengh=str.GetLength();
		SetColWidth(nField+1, (double)nLengh + 5);

		//SetColWidth(nField+1, m_dCellWidth);

		SetText(str);
		SetText(str);
	}

	// Row Data (Record)
	CString strData;
	_variant_t vValue; 
	int nStringLenth;

	if(GetMaxRows() < 1)
		return;

	if(!m_dataSource.ExecuteQuery(m_strQuery))
	{
		AfxMessageBox(_T("Error occur at m_dataSource.ExecuteQuery()"));
		return;
	}	

	SetRow(-1);
	SetCol(-1);
	SetFontSize(9);
	//hfont = CreateFont(13,0,0,0,m_bHeadFontBold,0,0,0,0,SHIFTJIS_CHARSET,0,0,0,_T("πŸ≈¡"));
	hfont = CreateFont(13,0,0,0,m_bHeadFontBold,0,0,0,0,SHIFTJIS_CHARSET,0,0,0,_T("MS PMincho"));
	SetFont(SS_ALLCOLS,SS_ALLROWS,hfont,TRUE);

	SetCellType(SS_ALLCOLS, SS_ALLROWS, &stype);
	m_dataSource.m_pRS->MoveFirst();

	SetRow(0);
	
	for(long lRow=1; lRow<=m_nMaxRows; lRow++)
	{
		SetRow(lRow);

		for(long lCol=1; lCol<=m_nMaxCols; lCol++)
		{
			SetCol(lCol);
			vValue = m_dataSource.m_pRS->Fields->Item[(long)(lCol-1)]->Value;

			if(vValue.vt != VT_NULL)
			{

				vValue.ChangeType(VT_BSTR);
				strData = vValue.bstrVal;

				SetText(strData);

				if(lRow==1)
				{

					nStringLenth=strData.GetLength();

					if(nStringLenth*2 > GetColWidth(lCol))
						SetColWidth(lCol,nStringLenth+10);
				}
				
			}

		}
		m_dataSource.m_pRS->MoveNext();
	}
	SetReDraw(TRUE);
}
