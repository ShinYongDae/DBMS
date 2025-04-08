//----------------------------------------------------------
//
// File: SSDLLMFC.H
//
// Copyright (C) 2004 FarPoint Technologies.
// All rights reserved.
//
//----------------------------------------------------------
         
#ifndef SSDLLMFC_H
#define SSDLLMFC_H

#include "ssdll.h"
#define TYPE_DEFAULT	1
#define TYPE_COMBO		8
#define TYPE_CHECK		10
#define TYPE_BUTTON		7
#define TYPE_STATIC		5

#define ALIGN_LEFT		0
#define ALIGN_RIGHT		1
#define ALIGN_CENTER	2

#define TYPE_NORMAL				0
#define TYPE_READ_ONLY			1
#define TYPE_ROW_MODE			2
#define TYPE_SINGLE_SELECT		3
#define TYPE_MULTI_SELECT		4
#define TYPE_EXTENDED_SELECT	5
//----------------------------------------------------------

class TSpread : public CWnd
{
  DECLARE_DYNAMIC(TSpread)

// Constructors
public:  
  TSpread() {}
  BOOL Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID);

  BOOL AddCellSpan(SS_COORD lCol, SS_COORD lRow, SS_COORD lNumCols, SS_COORD lNumRows)
	{ return SSAddCellSpan(m_hWnd, lCol, lRow, lNumCols, lNumRows); }

  BOOL AddCustomCellType(LPCTSTR lpszName, BOOL bEditable, BOOL bCanOverflow, BOOL bCanBeOverflown, BOOL bUseRendererControl, LPSS_CT_PROCS lpProcs)
  { return SSAddCustomCellType(m_hWnd, lpszName, bEditable, bCanOverflow, bCanBeOverflown, bUseRendererControl, lpProcs); }

  BOOL AddCustomFunction(LPCTSTR FunctionName, short ParameterCnt,
                         FARPROC lpfnFunctionProc)
        { return SSAddCustomFunction(m_hWnd, FunctionName,
                 ParameterCnt, lpfnFunctionProc); }

  BOOL AddCustomFunctionExt(LPCTSTR lpszFuncName, short nMinParamCnt,
                            short nMaxParamCnt, FARPROC lpfnFuncProc,
                            long lFlags)
        { return SSAddCustomFunctionExt(m_hWnd, lpszFuncName, nMinParamCnt,
                                        nMaxParamCnt, lpfnFuncProc, lFlags); }

  BOOL AddMultiSelBlocks(LPSS_SELBLOCK lpSelBlock)
        { return SSAddMultiSelBlocks(m_hWnd, lpSelBlock); }

  BOOL BuildDependencies()
        { return SSBuildDependencies(m_hWnd); }

  BOOL Clear(SS_COORD Col, SS_COORD Row)
        { return SSClear(m_hWnd, Col, Row); }

  BOOL ClearData(SS_COORD Col, SS_COORD Row)
        { return SSClearData(m_hWnd, Col, Row); }

  BOOL ClearDataRange(SS_COORD Col,  SS_COORD Row, 
                      SS_COORD Col2, SS_COORD Row2)
        { return SSClearDataRange(m_hWnd, Col, Row, Col2, Row2); }

  BOOL ClearRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2)
        { return SSClearRange(m_hWnd, Col, Row, Col2, Row2); }

  BOOL ClipboardCopy(void)
        { return (BOOL)SSClipboardCopy(m_hWnd); }

  BOOL ClipboardCut(void)
        { return SSClipboardCut(m_hWnd); }

  BOOL ClipboardPaste(void)
        { return SSClipboardPaste(m_hWnd); }

  BOOL ClipIn(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
              LPCTSTR Buffer, long BufferLen)
        { return SSClipIn(m_hWnd, Col, Row, Col2, Row2, Buffer, BufferLen); }

  GLOBALHANDLE ClipOut(SS_COORD Col, SS_COORD Row,
                       SS_COORD Col2,SS_COORD Row2)
        { return SSClipOut(m_hWnd, Col, Row, Col2, Row2); }

  BOOL ClipValueIn (SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
                    LPCTSTR Buffer, long BufferLen)
      { return SSClipValueIn(m_hWnd,Col,Row,Col2,Row2,Buffer,BufferLen); }

  GLOBALHANDLE ClipValueOut (SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2)
      { return (GLOBALHANDLE)SSClipValueOut(m_hWnd, Col,Row,Col2,Row2); }
      
  SS_COORD ColLetterToNumber(LPCTSTR lpszText)
	{ return SSColLetterToNumber(lpszText); }

  int ColWidthToLogUnits (double dfColWidth)
        { return SSColWidthToLogUnits(m_hWnd, dfColWidth); }

  LRESULT ComboBoxSendMessage (SS_COORD Col, SS_COORD Row, UINT Msg, WPARAM wParam, LPARAM lParam)
        { return SSComboBoxSendMessage(m_hWnd, Col, Row, Msg, wParam, lParam); }

  BOOL CopyRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2,
                 SS_COORD Row2, SS_COORD ColDest, SS_COORD RowDest)
        { return SSCopyRange(m_hWnd, Col, Row, Col2, Row2, ColDest, RowDest); }

  BOOL DelCol(SS_COORD Col)
        { return SSDelCol(m_hWnd, Col); }

  BOOL DelColRange(SS_COORD Col, SS_COORD Col2)
        { return SSDelColRange(m_hWnd, Col, Col2); }

  BOOL  DeleteSheets(short nSheetIndex, short nSheetCnt)
        { return SSDeleteSheets(m_hWnd, nSheetIndex, nSheetCnt); }

  BOOL DelRow(SS_COORD Row)
        { return SSDelRow(m_hWnd, Row); }

  BOOL DelRowRange(SS_COORD Row, SS_COORD Row2)
        { return SSDelRowRange(m_hWnd, Row, Row2); }

  BOOL DeSelectBlock(void)
        { return SSDeSelectBlock (m_hWnd); }

  BOOL EnumCustomCellType(LPCTSTR lpszPrevName, LPTSTR lpszName, int nLen)
  { return SSEnumCustomCellType(m_hWnd, lpszPrevName, lpszName, nLen); }

  BOOL EnumCustomCellTypeLen(LPCTSTR lpszPrevName)
  { return SSEnumCustomCellTypeLen(m_hWnd, lpszPrevName); }

  int  EnumCustomFunction(LPCTSTR lpszPrevFunc, LPTSTR lpszFunc, int nFuncLen)
		{ return SSEnumCustomFunction(m_hWnd, lpszPrevFunc, lpszFunc, nFuncLen); }
		  
  int  EnumCustomFunctionLen(LPCTSTR lpszPrevFunc)
  		{ return SSEnumCustomFunctionLen(m_hWnd, lpszPrevFunc); }

  BOOL ExportExcelBook (LPCTSTR lpszfileName, LPCTSTR lpszlogFileName)
        { return SSExportExcelBook (m_hWnd, lpszfileName, lpszlogFileName); }

  BOOL ExportExcelBookEx (LPCTSTR lpszfileName, LPCTSTR lpszlogFileName, SHORT nFlags)
        { return SSExportExcelBookEx (m_hWnd, lpszfileName, lpszlogFileName, nFlags); }

  BOOL ExportRangeToHTML (SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, LPCTSTR lpszFileName, BOOL bAppendFlag, LPCTSTR lpszLogFile)
        { return SSExportRangeToHTML (m_hWnd, Col, Row, Col2, Row2, lpszFileName, bAppendFlag, lpszLogFile); }

  BOOL ExportRangeToTextFile (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
		{ return SSExportRangeToTextFile (m_hWnd, lCol, lRow, lCol2, lRow2, pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

  BOOL ExportRangeToTextFileU (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
		{ return SSExportRangeToTextFileU (m_hWnd, lCol, lRow, lCol2, lRow2, pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

  BOOL ExportRangeToXML (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR lpszFileName, LPCTSTR pszRoot, LPCTSTR pszCollection, long lFlags, LPCTSTR lpszLogFile)
		{ return SSExportRangeToXML (m_hWnd, lCol, lRow, lCol2, lRow2, lpszFileName, pszRoot, pszCollection, lFlags, lpszLogFile); }

  BOOL ExportRangeToXMLBuffer (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR pszRoot, LPCTSTR pszCollection, HGLOBAL *hgBuff, long lFlags, LPCTSTR lpszLogFile)
		{ return SSExportRangeToXMLBuffer (m_hWnd, lCol, lRow, lCol2, lRow2, pszRoot, pszCollection, hgBuff, lFlags, lpszLogFile); }

  BOOL ExportToExcel(LPCTSTR lpszFileName, LPCTSTR lpszSheetName, LPCTSTR lpszLogFileName)
		{ return SSExportToExcel(m_hWnd, lpszFileName, lpszSheetName, lpszLogFileName); }

  BOOL ExportToExcelEx(LPCTSTR lpszFileName, LPCTSTR lpszSheetName, LPCTSTR lpszLogFileName, SHORT nFlags)
		{ return SSExportToExcelEx(m_hWnd, lpszFileName, lpszSheetName, lpszLogFileName, nFlags); }

  BOOL ExportToHTML (LPCTSTR lpszFileName, BOOL bAppendFlag, LPCTSTR lpszLogFile)
		{ return SSExportToHTML (m_hWnd, lpszFileName, bAppendFlag, lpszLogFile); }

  BOOL ExportToTextFile (LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
		{ return SSExportToTextFile (m_hWnd, pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

  BOOL ExportToTextFileU (LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
		{ return SSExportToTextFileU (m_hWnd, pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

  BOOL ExportToXML (LPCTSTR lpszFileName, LPCTSTR pszRoot, LPCTSTR pszCollection, long lFlags, LPCTSTR lpszLogFile)
		{ return SSExportToXML (m_hWnd, lpszFileName, pszRoot, pszCollection, lFlags, lpszLogFile); }

  BOOL ExportToXMLBuffer (LPCTSTR pszRoot, LPCTSTR pszCollection, HGLOBAL *hgBuff, long lFlags, LPCTSTR lpszLogFile)
		{ return SSExportToXMLBuffer (m_hWnd, pszRoot, pszCollection, hgBuff, lFlags, lpszLogFile); }

  void FloatFormat(LPSS_CELLTYPE lpCellType, double dfVal, LPTSTR lpszText, BOOL fValue)
        { SSFloatFormat(m_hWnd, lpCellType, dfVal, lpszText, fValue); }

  BOOL GetActionKey(WORD wAction, LPBOOL lpfShift, LPBOOL lpfCtrl, LPWORD lpwKey)
		{ return SSGetActionKey(m_hWnd, wAction, lpfShift, lpfCtrl, lpwKey); }

  BOOL SetActionKey(WORD wAction, BOOL fShift, BOOL fCtrl, WORD wKey)
		{ return SSSetActionKey(m_hWnd, wAction, fShift, fCtrl, wKey); }

  BOOL GetActiveCell (LPSS_COORD Col, LPSS_COORD Row)
        { return SSGetActiveCell(m_hWnd, Col, Row); }

  BOOL SetActiveCell(SS_COORD Col, SS_COORD Row)
        { return SSSetActiveCell (m_hWnd, Col, Row); }

  WORD GetActiveCellHighlightStyle ()
        { return SSGetActiveCellHighlightStyle(m_hWnd); }

  void SetActiveCellHighlightStyle(WORD wActiveCellHighlightStyle)
        { SSSetActiveCellHighlightStyle (m_hWnd, wActiveCellHighlightStyle); }

  short GetActiveSheet ()
        { return SSGetActiveSheet(m_hWnd); }

  BOOL SetActiveSheet(short nSheet)
        { return SSSetActiveSheet (m_hWnd, nSheet); }

  WORD GetAppearance()
		{ return SSGetAppearance(m_hWnd); }

  WORD SetAppearance(WORD wAppearance)
		{ return SSSetAppearance(m_hWnd, wAppearance); }
  
  BOOL GetArray(LPVOID lpArray, SS_COORD ColLeft,
				  SS_COORD RowTop, SS_COORD ArrayColCnt,
				  SS_COORD ArrayRowCnt, WORD wDataType)
		{ return SSGetArray(m_hWnd, lpArray, ColLeft, RowTop, ArrayColCnt, ArrayRowCnt, wDataType); }

  BOOL SetArray(LPVOID lpArray, SS_COORD ColLeft,
				  SS_COORD RowTop, SS_COORD ArrayColCnt,
				  SS_COORD ArrayRowCnt, WORD wDataType)
		{ return SSSetArray(m_hWnd, lpArray, ColLeft, RowTop, ArrayColCnt, ArrayRowCnt, wDataType); }

  void GetAutoSizeVisible (LPSS_COORD lpVisibleCols, LPSS_COORD lpVisibleRows)
        { SSGetAutoSizeVisible(m_hWnd, lpVisibleCols, lpVisibleRows); }

  void SetAutoSizeVisible(SS_COORD VisibleCols, SS_COORD VisibleRows)
        { SSSetAutoSizeVisible (m_hWnd, VisibleCols, VisibleRows); }

  WORD GetBackColorStyle()
        { return SSGetBackColorStyle(m_hWnd); }

  WORD SetBackColorStyle(WORD wStyle)
        { return SSSetBackColorStyle(m_hWnd, wStyle); }
 
  BOOL GetBool(short nIndex)
        { return SSGetBool(m_hWnd, nIndex); }

  BOOL SetBool(short nIndex, BOOL bNewVal)
        { return SSSetBool (m_hWnd, nIndex, bNewVal); }

  BOOL GetBorder(SS_COORD Col, SS_COORD Row, LPWORD lpwBorderType,
                 LPWORD lpwBorderStyle, LPCOLORREF lpColor)
        { return SSGetBorder(m_hWnd, Col, Row, lpwBorderType, lpwBorderStyle, lpColor); }

  BOOL SetBorder(SS_COORD Col, SS_COORD Row, WORD wBorderType,
         WORD wBorderStyle, COLORREF Color)
        { return SSSetBorder(m_hWnd, Col, Row, wBorderType, wBorderStyle, Color); }

  BOOL SetBorderRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2,
                      SS_COORD Row2, WORD wBorderType, WORD wBorderStyle, COLORREF Color)
        { return SSSetBorderRange(m_hWnd, Col, Row, Col2, Row2, wBorderType, wBorderStyle, Color); }

  BOOL GetBorderEx(SS_COORD Col, SS_COORD Row, 
					LPWORD lpwStyleLeft,	LPCOLORREF lpclrLeft, 
					LPWORD lpwStyleTop,		LPCOLORREF lpclrTop,
					LPWORD lpwStyleRight,	LPCOLORREF lpclrRight,
					LPWORD lpwStyleBottom,	LPCOLORREF lpclrBottom)
        { return SSGetBorderEx(m_hWnd, Col, Row, lpwStyleLeft, lpclrLeft, lpwStyleTop, lpclrTop,
				 lpwStyleRight, lpclrRight, lpwStyleBottom, lpclrBottom); }

  BOOL GetBottomRightCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
        { return SSGetBottomRightCell(m_hWnd, lpCol, lpRow); }

  WORD GetButtonDrawMode(void)
        { return SSGetButtonDrawMode(m_hWnd); }

  WORD SetButtonDrawMode(WORD wMode)
        { return SSSetButtonDrawMode(m_hWnd, wMode); }

  int GetButtonText(SS_COORD Col, SS_COORD Row, LPTSTR lpText)
        { return SSGetButtonText(m_hWnd, Col, Row, lpText); }

  BOOL SetButtonText(SS_COORD Col, SS_COORD Row, LPCTSTR lpText)
        { return SSSetButtonText(m_hWnd, Col, Row, lpText); }

  WNDPROC GetCallBack(void)
        { return SSGetCallBack(m_hWnd); }

  WNDPROC SetCallBack(WNDPROC lpfnCallBack)
        { return SSSetCallBack(m_hWnd, lpfnCallBack); }

  void SetCalText(LPCTSTR lpszShortDays,		LPCTSTR lpszLongDays, 
					LPCTSTR lpszShortMonths,	LPCTSTR lpszLongMonths, 
					LPCTSTR lpszOkText,			LPCTSTR lpszCancelText)
        { SSSetCalText(lpszShortDays, lpszLongDays, lpszShortMonths, lpszLongMonths, lpszOkText, lpszCancelText); }

  void GetCalTextOverride(LPTSTR lpszShortDays, LPSHORT lpnLenShortDays, 
                          LPTSTR lpszLongDays, LPSHORT lpnLenLongDays,
                          LPTSTR lpszShortMonths, LPSHORT lpnLenShortMonths, 
                          LPTSTR lpszLongMonths, LPSHORT lpnLenLongMonths,
                          LPTSTR lpszOkText, LPSHORT lpnLenOkText, 
                          LPTSTR lpszCancelText, LPSHORT lpnLenCancelText)
		{ SSGetCalTextOverride(m_hWnd, lpszShortDays, lpnLenShortDays, 
                               lpszLongDays, lpnLenLongDays,
                               lpszShortMonths, lpnLenShortMonths, 
                               lpszLongMonths, lpnLenLongMonths,
                               lpszOkText, lpnLenOkText, 
                               lpszCancelText, lpnLenCancelText); }

  void SetCalTextOverride(LPCTSTR lpszShortDays, LPCTSTR lpszLongDays,
                          LPCTSTR lpszShortMonths, LPCTSTR lpszLongMonths,
                          LPCTSTR lpszOkText, LPCTSTR lpszCancelText) 
		{ SSSetCalTextOverride(m_hWnd, lpszShortDays, lpszLongDays, lpszShortMonths,
							   lpszLongMonths, lpszOkText, lpszCancelText); }

  BOOL GetCellFromPixel(LPSS_COORD lpCol, LPSS_COORD lpRow,
                        int MouseX, int MouseY)
        { return SSGetCellFromPixel(m_hWnd, lpCol, lpRow, MouseX, MouseY); }

  short GetCellNote(SS_COORD Col, SS_COORD Row, LPTSTR Note)
		{ return SSGetCellNote(m_hWnd, Col, Row, Note); }
  
  BOOL SetCellNote(SS_COORD Col, SS_COORD Row, LPTSTR Note)
		{ return SSSetCellNote(m_hWnd, Col, Row, Note); }

  BOOL SetCellNoteRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, LPTSTR Note)
		{ return SSSetCellNoteRange(m_hWnd, Col, Row, Col2, Row2, Note); }

  WORD GetCellNoteIndicator()
		{ return SSGetCellNoteIndicator(m_hWnd); }

  void SetCellNoteIndicator(WORD wValue)
		{ SSSetCellNoteIndicator(m_hWnd, wValue); }

  BOOL GetCellRect(SS_COORD Col, SS_COORD Row, LPRECT lpRect)
        { return SSGetCellRect(m_hWnd, Col, Row, lpRect); }

  void GetCellSendingMsg(LPSS_COORD lpCol, LPSS_COORD lpRow)
        { SSGetCellSendingMsg(m_hWnd, lpCol, lpRow); }

  WORD GetCellSpan(SS_COORD lCol, SS_COORD lRow, LPSS_COORD lplColAnchor, LPSS_COORD lplRowAnchor, 	LPSS_COORD lplNumCols, LPSS_COORD lplNumRows)
	{ return SSGetCellSpan(m_hWnd, lCol, lRow, lplColAnchor, lplRowAnchor, lplNumCols, 		lplNumRows);}

  short GetCellTag(SS_COORD Col, SS_COORD Row, LPTSTR CellTag)
	{ return SSGetCellTag(m_hWnd, Col, Row, CellTag); }

  BOOL SetCellTag(SS_COORD Col, SS_COORD Row, LPCTSTR CellTag)
	{ return SSSetCellTag(m_hWnd, Col, Row, CellTag); }

  BOOL GetCellType(SS_COORD Col, SS_COORD Row, LPSS_CELLTYPE lpCellType)
        { return SSGetCellType(m_hWnd, Col, Row, lpCellType); }

  BOOL SetCellType(SS_COORD Col, SS_COORD Row, LPSS_CELLTYPE lpCellType)
        { return SSSetCellType(m_hWnd, Col, Row, lpCellType); }

  BOOL SetCellTypeRange(SS_COORD Col,  SS_COORD Row, SS_COORD Col2,
                        SS_COORD Row2, LPSS_CELLTYPE lpCellType)
        { return SSSetCellTypeRange(m_hWnd, Col, Row, Col2, Row2, lpCellType); }

  void GetClientRect(LPRECT lpRect)
        { SSGetClientRect(m_hWnd, lpRect); }

  WORD GetClipboardOptions()
		{ return SSGetClipboardOptions(m_hWnd); }

  WORD SetClipboardOptions(WORD wOptions)
		{ return SSSetClipboardOptions(m_hWnd, wOptions); }

  SS_COORD GetColFromID(LPCTSTR lpszColID)
	{ return SSGetColFromID(m_hWnd, lpszColID); }

  WORD GetColHeaderDisplay()
        { return SSGetColHeaderDisplay(m_hWnd); }

  WORD SetColHeaderDisplay(WORD wDisplay)
        { return SSSetColHeaderDisplay(m_hWnd, wDisplay); }

  SS_COORD GetColHeaderRows()
	{ return SSGetColHeaderRows(m_hWnd); }

  void SetColHeaderRows(SS_COORD lColHeaderRows)
	{ SSSetColHeaderRows(m_hWnd, lColHeaderRows); }

  SS_COORD GetColHeadersAutoTextIndex()
	{ return SSGetColHeadersAutoTextIndex(m_hWnd); }

  void SetColHeadersAutoTextIndex(SS_COORD lColHeadersAutoTextIndex)
	{ SSSetColHeadersAutoTextIndex(m_hWnd, lColHeadersAutoTextIndex); }

  SS_COORD GetColHeadersUserSortIndex()
	{ return SSGetColHeadersUserSortIndex(m_hWnd); }

  void SetColHeadersUserSortIndex(SS_COORD lColHeadersUserSortIndex)
	{ SSSetColHeadersUserSortIndex(m_hWnd, lColHeadersUserSortIndex); }

  short GetColID(SS_COORD Col, LPTSTR ColID)
	{ return SSGetColID(m_hWnd, Col, ColID); }

  BOOL  SetColID(SS_COORD Col, LPCTSTR ColID)
	{ return SSSetColID(m_hWnd, Col, ColID); }

  short GetColMerge(SS_COORD Col)
	{ return SSGetColMerge(m_hWnd, Col); }

  void SetColMerge(SS_COORD Col, short nColMerge)
	{ SSSetColMerge(m_hWnd, Col, nColMerge); }

  BOOL GetColor(SS_COORD Col, SS_COORD Row, LPCOLORREF lpBackground,
                LPCOLORREF lpForeground)
        { return SSGetColor(m_hWnd, Col, Row, lpBackground, lpForeground); }
                  
  BOOL SetColor(SS_COORD Col, SS_COORD Row, COLORREF Background,
                COLORREF Foreground)
        { return SSSetColor(m_hWnd, Col, Row, Background, Foreground); }

  BOOL SetColorRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
                     COLORREF Background, COLORREF Foreground)
        { return SSSetColorRange(m_hWnd, Col, Row, Col2, Row2, Background, Foreground); }

  BOOL GetColPageBreak(SS_COORD col)
	{ return (BOOL)SSGetColPageBreak(m_hWnd, col); }

  BOOL SetColPageBreak(SS_COORD col, BOOL fpagebreak)
        { return (BOOL)SSSetColPageBreak(m_hWnd, col, fpagebreak); }

#if defined(_WIN64) || defined(_IA64)
  BOOL GetColUserData(SS_COORD Col,LONG_PTR* lplUserData)
#else 
  BOOL GetColUserData(SS_COORD Col,long* lplUserData)
#endif
        { return SSGetColUserData(m_hWnd, Col, lplUserData); }

#if defined(_WIN64) || defined(_IA64)
  BOOL SetColUserData(SS_COORD Col, LONG_PTR lUserData)
#else
  BOOL SetColUserData(SS_COORD Col, long lUserData)
#endif
        { return SSSetColUserData(m_hWnd, Col, lUserData); }

  short GetColUserSortIndicator(SS_COORD Col)
		{ return SSGetColUserSortIndicator(m_hWnd, Col); }

  void SetColUserSortIndicator(SS_COORD Col, short nColUserSortIndicator)
		{ SSSetColUserSortIndicator(m_hWnd, Col, nColUserSortIndicator); }

  BOOL GetColWidth(SS_COORD Col, LPDOUBLE lpColWidth)
        { return SSGetColWidth(m_hWnd, Col, lpColWidth); }

  BOOL SetColWidth(SS_COORD Col, double Width)
        { return SSSetColWidth(m_hWnd, Col, Width); }

  BOOL GetColWidthInPixels(SS_COORD Col, LPINT lpdWidthPixels)
        { return SSGetColWidthInPixels(m_hWnd, Col, lpdWidthPixels); }

  BOOL SetColWidthInPixels(SS_COORD Col, int dWidth)
        { return SSSetColWidthInPixels(m_hWnd, Col, dWidth); }

  BOOL GetCurrSelBlockPos(LPSS_COORD CurrentCol, LPSS_COORD CurrentRow)
        { return SSGetCurrSelBlockPos(m_hWnd, CurrentCol, CurrentRow); }

  HCURSOR GetCursor(WORD wType)
        { return SSGetCursor(m_hWnd, wType); }

  BOOL SetCursor(WORD wType, HCURSOR hCursor)
        { return SSSetCursor(m_hWnd, wType, hCursor); }

  BOOL GetCustomCellType(LPCTSTR lpszName, LPBOOL lpbEditable, LPBOOL lpbCanOverflow, LPBOOL lpbCanBeOverflown, LPBOOL lpbUseRendererControl, LPSS_CT_PROCS lpProcs)
  { return SSGetCustomCellType(m_hWnd, lpszName, lpbEditable, lpbCanOverflow, lpbCanBeOverflown, lpbUseRendererControl, lpProcs); }

  BOOL GetCustomFunction(LPCTSTR lpszFunc, LPSHORT lpnMinParamCnt,
                         LPSHORT lpnMaxParamCnt, LPLONG lplFlags)
		{ return SSGetCustomFunction(m_hWnd, lpszFunc, lpnMinParamCnt,
									 lpnMaxParamCnt, lplFlags); }

  int GetCustomName(LPCTSTR lpszName, LPTSTR lpszValue, int nValueLen)
        { return SSGetCustomName(m_hWnd, lpszName, lpszValue, nValueLen); }

  BOOL SetCustomName(LPCTSTR lpszName, LPCTSTR lpszValue)
        { return SSSetCustomName(m_hWnd, lpszName, lpszValue); }

  int GetCustomNameLen(LPCTSTR lpszName)
        { return SSGetCustomNameLen(m_hWnd, lpszName); }

  int GetData(SS_COORD Col, SS_COORD Row, LPTSTR Data)
        { return SSGetData(m_hWnd, Col, Row, Data); }

  BOOL SetData(SS_COORD Col, SS_COORD Row, LPCTSTR lpData)
        { return SSSetData (m_hWnd, Col, Row, lpData); }

  BOOL SetDataRange(SS_COORD Col, SS_COORD Row,
                    SS_COORD Col2,SS_COORD Row2, LPCTSTR lpData)
        { return SSSetDataRange (m_hWnd, Col, Row, Col2, Row2, lpData); }

  BOOL GetDataCnt(LPSS_COORD lpColCnt, LPSS_COORD lpRowCnt)
        { return SSGetDataCnt(m_hWnd, lpColCnt, lpRowCnt); }

  int GetDataLen(SS_COORD Col, SS_COORD Row)
        { return SSGetDataLen(m_hWnd, Col, Row); }

  BOOL GetDefFloatFormat(LPSS_FLOATFORMAT lpFloatFormat)
        { return SSGetDefFloatFormat(m_hWnd, lpFloatFormat); }

  BOOL SetDefFloatFormat(LPSS_FLOATFORMAT lpFloatFormat)
        { return SSSetDefFloatFormat(m_hWnd, lpFloatFormat); }

  WORD GetEditEnterAction(void)
        { return SSGetEditEnterAction(m_hWnd); }

  WORD SetEditEnterAction(WORD wAction)
        { return SSSetEditEnterAction(m_hWnd, wAction); }

  BOOL GetEditMode(void)
        { return SSGetEditMode(m_hWnd); }
                       
  BOOL SetEditMode(BOOL fEditModeOn)
        { return SSSetEditMode(m_hWnd, fEditModeOn); }

  BOOL GetExcelSheetList(LPCTSTR lpszFileName, GLOBALHANDLE FAR *lpghList, 
						 LPSHORT lpnListCount, LPCTSTR lpszLogFileName, 
						 LPSHORT lpnWorkbookHandle, BOOL fReplace)
		{ return SSGetExcelSheetList(m_hWnd, lpszFileName, lpghList, lpnListCount, lpszLogFileName, lpnWorkbookHandle, fReplace);}

  BOOL GetFirstValidCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
        { return SSGetFirstValidCell(m_hWnd, lpCol, lpRow); }
  
  BOOL GetFloat(SS_COORD Col, SS_COORD Row, LPDOUBLE lpdfValue)
        { return SSGetFloat(m_hWnd, Col, Row, lpdfValue); }

  BOOL SetFloat(SS_COORD Col, SS_COORD Row, double dfValue)
        { return SSSetFloat(m_hWnd, Col, Row, dfValue); }

  BOOL SetFloatRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, double dfValue)
        { return SSSetFloatRange(m_hWnd, Col, Row, Col2, Row2, dfValue); }

  HFONT GetFont(SS_COORD Col, SS_COORD Row)
        { return SSGetFont(m_hWnd, Col, Row); }

  BOOL SetFont(SS_COORD Col, SS_COORD Row, HFONT hFont, BOOL fDeleteFont)
        { return SSSetFont(m_hWnd, Col, Row, hFont, fDeleteFont); }

  BOOL SetFontRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2,
                    SS_COORD Row2, HFONT hFont, BOOL fDeleteFont)
        { return SSSetFontRange(m_hWnd, Col, Row, Col2, Row2, hFont, fDeleteFont); }

  int GetFormula(SS_COORD Col, SS_COORD Row, LPTSTR lpFormula)
        { return SSGetFormula(m_hWnd, Col, Row, lpFormula); }

  BOOL SetFormula(SS_COORD Col, SS_COORD Row, LPCTSTR Formula,
                  BOOL BuildDependencies)
        { return SSSetFormula (m_hWnd, Col, Row, Formula, BuildDependencies); }

  BOOL SetFormulaRange(SS_COORD Col, SS_COORD Row,
                       SS_COORD Col2,SS_COORD Row2,
                       LPCTSTR Formula, BOOL BuildDependencies)
        { return SSSetFormulaRange(m_hWnd, Col, Row, Col2, Row2, Formula, 
                                   BuildDependencies); }

  int GetFormulaLen(SS_COORD Col, SS_COORD Row)
        { return SSGetFormulaLen(m_hWnd, Col, Row); }

  BOOL GetFreeze(LPSS_COORD lpColsFrozen, LPSS_COORD lpRowsFrozen)
        { return SSGetFreeze(m_hWnd, lpColsFrozen, lpRowsFrozen); }

  BOOL SetFreeze(SS_COORD ColsFrozen, SS_COORD RowsFrozen)
         { return SSSetFreeze (m_hWnd, ColsFrozen, RowsFrozen); }

  void GetGrayAreaColor(LPCOLORREF lpBackground, LPCOLORREF lpForeground)
        { SSGetGrayAreaColor(m_hWnd, lpBackground, lpForeground); }

  void SetGrayAreaColor(COLORREF Background, COLORREF Foreground)
        { SSSetGrayAreaColor(m_hWnd, Background, Foreground); }

  COLORREF GetGridColor(void)
        { return SSGetGridColor(m_hWnd); }

  COLORREF SetGridColor(COLORREF Color)
        { return SSSetGridColor(m_hWnd, Color); }

  WORD GetGridType(void)
        { return SSGetGridType(m_hWnd); }

  WORD SetGridType(WORD wGridType)
        { return SSSetGridType(m_hWnd, wGridType); }

  BOOL GetBuffer(SS_COORD Col, SS_COORD Row, LPVOID lpBuffer, long *plSize) 
        {return SSGetBuffer(m_hWnd, Col, Row, lpBuffer, plSize);} 
  
  BOOL SetBuffer(SS_COORD Col, SS_COORD Row, LPVOID lpBuffer, long lSize) 
        {return SSSetBuffer(m_hWnd, Col, Row, lpBuffer, lSize);}

  BOOL GetInteger(SS_COORD Col, SS_COORD Row, LPLONG lplValue)
        { return SSGetInteger(m_hWnd, Col, Row, lplValue); }

  BOOL SetInteger(SS_COORD Col, SS_COORD Row, long lValue)
        { return SSSetInteger(m_hWnd, Col, Row, lValue); }

  BOOL SetIntegerRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, long lValue)
        { return SSSetIntegerRange(m_hWnd, Col, Row, Col2, Row2, lValue); }

  BOOL GetIteration(LPSHORT lpnMaxIterations, LPDOUBLE lpdfMaxChange)
        { return SSGetIteration(m_hWnd, lpnMaxIterations, lpdfMaxChange); }

  BOOL SetIteration(BOOL bIteration, short nMaxIterations, double dfMaxChange)
        { return SSSetIteration(m_hWnd, bIteration, nMaxIterations, dfMaxChange); }

  BOOL GetLastValidCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
        { return SSGetLastValidCell(m_hWnd, lpCol, lpRow); }

  BOOL GetLock(SS_COORD Col, SS_COORD Row)
        { return SSGetLock(m_hWnd, Col, Row); }

  BOOL SetLock(SS_COORD Col, SS_COORD Row, BOOL Lock)
        { return SSSetLock (m_hWnd, Col, Row, Lock); }

  BOOL SetLockRange(SS_COORD Col,  SS_COORD Row,
                    SS_COORD Col2, SS_COORD Row2, BOOL Lock)
        { return SSSetLockRange (m_hWnd, Col, Row, Col2, Row2, Lock); }

  void GetLockColor(LPCOLORREF lpBackground, LPCOLORREF lpForeground)
        { SSGetLockColor(m_hWnd, lpBackground, lpForeground); }

  void SetLockColor(COLORREF Background, COLORREF Foreground)
        { SSSetLockColor(m_hWnd, Background, Foreground); }

  SS_COORD GetMaxCols()
        { return SSGetMaxCols (m_hWnd); }

  void SetMaxCols(SS_COORD MaxCols)
        { SSSetMaxCols(m_hWnd, MaxCols); }

  SS_COORD GetMaxRows()
        { return SSGetMaxRows(m_hWnd); }

  void SetMaxRows(SS_COORD MaxRows)
        { SSSetMaxRows(m_hWnd, MaxRows); }

  BOOL GetMaxTextCellSize(SS_COORD Col, SS_COORD Row, LPDOUBLE lpdfWidth, LPDOUBLE lpdfHeight)
        { return SSGetMaxTextCellSize(m_hWnd, Col, Row, lpdfWidth, lpdfHeight); }

  BOOL GetMaxTextColWidth(SS_COORD Col, LPDOUBLE lpdfColWidth)
        { return SSGetMaxTextColWidth(m_hWnd, Col, lpdfColWidth); }

  BOOL GetMaxTextRowHeight(SS_COORD Row, LPDOUBLE lpdfRowHeight)
        { return SSGetMaxTextRowHeight(m_hWnd, Row, lpdfRowHeight); }

  GLOBALHANDLE GetMultiSelBlocks(LPSHORT lpdSelBlockCnt)
        { return SSGetMultiSelBlocks(m_hWnd, lpdSelBlockCnt); }

  BOOL SetMultiSelBlocks(LPSS_SELBLOCK lpSelBlockList, short dSelBlockCnt)
        { return SSSetMultiSelBlocks(m_hWnd, lpSelBlockList, dSelBlockCnt); }

  long GetNextPageBreakCol(long lPrevCol)
		{ return SSGetNextPageBreakCol(m_hWnd, lPrevCol); }

  long GetNextPageBreakRow(long lPrevRow)
		{ return SSGetNextPageBreakRow(m_hWnd, lPrevRow); }

  BOOL GetOddEvenRowColor(LPCOLORREF lpclrBackOdd, LPCOLORREF lpclrForeOdd,
							LPCOLORREF lpclrBackEven, LPCOLORREF lpclrForeEven)
		{ return SSGetOddEvenRowColor(m_hWnd, lpclrBackOdd, lpclrForeOdd, lpclrBackEven, lpclrForeEven); }

  BOOL SetOddEvenRowColor(COLORREF clrBackOdd, COLORREF clrForeOdd,
							COLORREF clrBackEven, COLORREF clrForeEven)
		{ return SSSetOddEvenRowColor(m_hWnd, clrBackOdd, clrForeOdd, clrBackEven, clrForeEven); }

  WORD GetOperationMode()
        { return SSGetOperationMode(m_hWnd); }

  WORD SetOperationMode(WORD wMode)
        { return SSSetOperationMode(m_hWnd, wMode); }
  
  HWND GetOwner()
        { return SSGetOwner(m_hWnd); }

  HWND SetOwner(HWND hWndOwner)
        { return SSSetOwner(m_hWnd, hWndOwner); }

  short GetPrintAbortMsg(LPTSTR lpszText, short nLen)
        { return SSGetPrintAbortMsg(m_hWnd, lpszText, nLen); }

  BOOL SetPrintAbortMsg(LPTSTR lpszText)
        { return SSSetPrintAbortMsg(m_hWnd, lpszText); }

  short GetPrintJobName(LPTSTR lpszText, short nLen)
        { return SSGetPrintJobName(m_hWnd, lpszText, nLen); }

  BOOL SetPrintJobName(LPTSTR lpszText)
        { return SSSetPrintJobName(m_hWnd, lpszText); }

#if defined(_WIN64) || defined(_IA64)
  BOOL GetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC FAR *lpfnPrintProcAddr,LONG_PTR *lplAppData)
#else
  BOOL GetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC FAR *lpfnPrintProcAddr,long *lplAppData)
#endif
		{ return SSGetPrintOptions(m_hWnd, lpPrintFormat, lpfnPrintProcAddr, lplAppData); }

#if defined(_WIN64) || defined(_IA64)
  void SetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC lpfnPrintProc, LONG_PTR lAppData)
#else
  void SetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC lpfnPrintProc, long lAppData)
#endif
		{ SSSetPrintOptions(m_hWnd, lpPrintFormat, lpfnPrintProc, lAppData); }

#if defined(_WIN64) || defined(_IA64)
  void GetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC *lpfnPrintProc, LONG_PTR *lplAppData)
#else
  void GetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC *lpfnPrintProc, long *lplAppData)
#endif
		{ SSGetPrintOptionsEx(m_hWnd, lpPageFormat, lpfnPrintProc, lplAppData); }

#if defined(_WIN64) || defined(_IA64)
  void SetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC lpfnPrintProc, LONG_PTR lAppData)
#else
  void SetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC lpfnPrintProc, long lAppData)
#endif
		{ SSSetPrintOptionsEx(m_hWnd, lpPageFormat, lpfnPrintProc, lAppData); }

  long GetPrintPageCount()
		{ return SSGetPrintPageCount(m_hWnd); }

  short GetRefStyle()
        { return SSGetRefStyle(m_hWnd); }

  short SetRefStyle(short nStyle)
        { return SSSetRefStyle(m_hWnd, nStyle); }

  SS_COORD GetRowHeaderCols()
	{ return SSGetRowHeaderCols(m_hWnd); }

  void SetRowHeaderCols(SS_COORD lRowHeaderCols)
	{ SSSetRowHeaderCols(m_hWnd, lRowHeaderCols); }

  WORD GetRowHeaderDisplay()
        { return SSGetRowHeaderDisplay(m_hWnd); }

  WORD SetRowHeaderDisplay(WORD wDisplay)
        { return SSSetRowHeaderDisplay(m_hWnd, wDisplay); }

  SS_COORD GetRowHeadersAutoTextIndex()
	{ return SSGetRowHeadersAutoTextIndex(m_hWnd); }

  void SetRowHeadersAutoTextIndex(SS_COORD lRowHeadersAutoTextIndex)
	{ SSSetRowHeadersAutoTextIndex(m_hWnd, lRowHeadersAutoTextIndex); }

  BOOL  GetRowHeight(SS_COORD Row, LPDOUBLE lpHeight)
        { return (BOOL)SSGetRowHeight(m_hWnd, Row, lpHeight); }

  BOOL SetRowHeight(SS_COORD Row, double Height)
        { return SSSetRowHeight (m_hWnd, Row, Height); }

  BOOL GetRowHeightInPixels(SS_COORD Row, LPINT lpdHeightPixels)
        { return SSGetRowHeightInPixels(m_hWnd, Row, lpdHeightPixels); }

  BOOL SetRowHeightInPixels(SS_COORD Row, int dHeight)
        { return SSSetRowHeightInPixels(m_hWnd, Row, dHeight); }

  short GetRowMerge(SS_COORD Row)
	{ return SSGetRowMerge(m_hWnd, Row); }

  void SetRowMerge(SS_COORD Row, short nRowMerge)
	{ SSSetRowMerge(m_hWnd, Row, nRowMerge); }
 
  BOOL GetRowPageBreak(SS_COORD row)
        { return SSGetRowPageBreak(m_hWnd, row); }  

  BOOL SetRowPageBreak(SS_COORD row, BOOL fpagebreak)
        { return (BOOL)SSSetRowPageBreak(m_hWnd, row, fpagebreak); }

#if defined(_WIN64) || defined(_IA64)
  BOOL GetRowUserData(SS_COORD Row, LONG_PTR* lplUserData)
#else
  BOOL GetRowUserData(SS_COORD Row, long* lplUserData)
#endif
        { return SSGetRowUserData(m_hWnd, Row, lplUserData); }

#if defined(_WIN64) || defined(_IA64)
  BOOL SetRowUserData(SS_COORD Row, LONG_PTR lUserData)
#else
  BOOL SetRowUserData(SS_COORD Row, long lUserData)
#endif
        { return SSSetRowUserData(m_hWnd, Row, lUserData); }

  void GetScrollBarColor(LPCOLORREF lpVScrollBarColor, LPCOLORREF lpHScrollBarColor)
        { SSGetScrollBarColor(m_hWnd, lpVScrollBarColor, lpHScrollBarColor); }

  void SetScrollBarColor(COLORREF VScrollBarColor, COLORREF HScrollBarColor)
        { SSSetScrollBarColor(m_hWnd, VScrollBarColor, HScrollBarColor); }

  void GetScrollBarSize(LPSHORT lpdVScrollWidth, LPSHORT lpdHScrollHeight)
        { SSGetScrollBarSize(m_hWnd, lpdVScrollWidth, lpdHScrollHeight); }

  void SetScrollBarSize(short dVScrollWidth, short dHScrollHeight)
        { SSSetScrollBarSize(m_hWnd, dVScrollWidth, dHScrollHeight); }

  BOOL GetSel(LPINT SelStart, LPINT SelEnd)
        { return SSGetSel (m_hWnd, SelStart, SelEnd); }

  BOOL SetSel(int SelStart, int SelEnd)
        { return SSSetSel(m_hWnd, SelStart, SelEnd); }

  WORD GetSelBlockOptions()
        { return SSGetSelBlockOptions(m_hWnd); }

  WORD SetSelBlockOptions(WORD wOption)
        { return SSSetSelBlockOptions(m_hWnd, wOption); }

  void GetSelColor(LPCOLORREF lpclrBack, LPCOLORREF lpclrFore)
		{ SSGetSelColor(m_hWnd, lpclrBack, lpclrFore); }

  void SetSelColor(COLORREF clrBack, COLORREF clrFore)
		{ SSSetSelColor(m_hWnd, clrBack, clrFore); }

  BOOL GetSelectBlock(LPSS_CELLCOORD CellUL, LPSS_CELLCOORD CellLR)
        { return SSGetSelectBlock (m_hWnd, CellUL, CellLR); }

  BOOL SetSelectBlock(LPSS_CELLCOORD CellUL, LPSS_CELLCOORD CellLR)
        { return SSSetSelectBlock (m_hWnd, CellUL, CellLR); }

  GLOBALHANDLE GetSelText()
        { return SSGetSelText (m_hWnd); }

  void GetShadowColor(LPCOLORREF lpShadowColor, LPCOLORREF lpShadowText,
                      LPCOLORREF lpShadowDark,  LPCOLORREF lpShadowLight)
        { SSGetShadowColor (m_hWnd, lpShadowColor, lpShadowText,
                            lpShadowDark, lpShadowLight); }

  void SetShadowColor(COLORREF ShadowColor, COLORREF ShadowText,
                      COLORREF ShadowDark,  COLORREF ShadowLight)
        { SSSetShadowColor(m_hWnd, ShadowColor, ShadowText,
                           ShadowDark, ShadowLight); }

  short GetSheet()
		{ return SSGetSheet(m_hWnd); }

  void SetSheet(short nSheet)
		{ SSSetSheet(m_hWnd, nSheet); }

  short GetSheetCount()
		{ return SSGetSheetCount(m_hWnd); }

  BOOL SetSheetCount(short nSheetCnt)
		{ return SSSetSheetCount(m_hWnd, nSheetCnt); }

  short GetSheetSendingMsg()
		{ return SSGetSheetSendingMsg(m_hWnd); }

  short GetSheetName(short nSheet, LPTSTR lpszName, short nLen)
		{ return SSGetSheetName(m_hWnd, nSheet, lpszName, nLen); }

  BOOL SetSheetName(short nSheet, LPCTSTR lpszSheetName)
		{ return SSSetSheetName(m_hWnd, nSheet, lpszSheetName); }

  BOOL GetSheetVisible(short nSheet)
      { return SSGetSheetVisible(m_hWnd, nSheet); }

  BOOL SetSheetVisible(short nSheet, BOOL fVisible)
      { return SSSetSheetVisible(m_hWnd, nSheet, fVisible); }

  WORD GetShowScrollTips()
		{ return SSGetShowScrollTips(m_hWnd); }

  void SetShowScrollTips(WORD wScrollTips)
		{ SSSetShowScrollTips(m_hWnd, wScrollTips); }

  void GetStartingNumbers(LPSS_COORD lpStartingColNumber,
                          LPSS_COORD lpStartingRowNumber)
        { SSGetStartingNumbers(m_hWnd, lpStartingColNumber,
                                      lpStartingRowNumber); }

  void SetStartingNumbers(SS_COORD StartingColNumber,
                          SS_COORD StartingRowNumber)
        { SSSetStartingNumbers(m_hWnd, StartingColNumber,
                                      StartingRowNumber); }

  HFONT GetTabStripFont()
		{ return SSGetTabStripFont(m_hWnd); }

  BOOL SetTabStripFont(HFONT hFont)
		{ return SSSetTabStripFont(m_hWnd, hFont); }

  short GetTabStripLeftSheet()
		{ return SSGetTabStripLeftSheet(m_hWnd); }

  BOOL SetTabStripLeftSheet(short nSheet)
		{ return SSSetTabStripLeftSheet(m_hWnd, nSheet); }

  WORD GetTabStripPolicy()
		{ return SSGetTabStripPolicy(m_hWnd); }

  WORD SetTabStripPolicy(WORD wTabStripPolicy)
		{ return SSSetTabStripPolicy(m_hWnd, wTabStripPolicy); }

  double GetTabStripRatio()
		{ return SSGetTabStripRatio(m_hWnd); }

  double SetTabStripRatio(double nTabStripRatio)
		{ return SSSetTabStripRatio(m_hWnd, nTabStripRatio); }

  BOOL GetTextTip(LPWORD lpwStatus, LPLONG lplDelay, LPLOGFONT lpLogFont,
					LPCOLORREF lpclrBack, LPCOLORREF lpclrFore)
		{ return SSGetTextTip(m_hWnd, lpwStatus, lplDelay, lpLogFont, lpclrBack, lpclrFore); }

  BOOL SetTextTip(WORD wStatus, LONG lDelay, LPLOGFONT lpLogFont,
				  COLORREF clrBack, COLORREF clrFore)
		{ return SSSetTextTip(m_hWnd, wStatus, lDelay, lpLogFont,
							  clrBack, clrFore); }

  BOOL GetTopLeftCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
        { return SSGetTopLeftCell (m_hWnd, lpCol, lpRow); }

  void GetTopLeftPrev(LPSS_COORD lpCol, LPSS_COORD lpRow)
        { SSGetTopLeftPrev(m_hWnd, lpCol, lpRow); }

  short GetTwoDigitYearMax()
		{ return SSGetTwoDigitYearMax(m_hWnd); }

  BOOL SetTwoDigitYearMax(short nTwoDigitYearMax)
		{ return SSSetTwoDigitYearMax(m_hWnd, nTwoDigitYearMax); }

  WORD GetUnitType()
        { return SSGetUnitType(m_hWnd); }

  WORD SetUnitType(WORD wUnitType)
        { return SSSetUnitType(m_hWnd, wUnitType); }

  short GetUserColAction()
		{return SSGetUserColAction(m_hWnd); }

  void SetUserColAction(short nUserColAction)
		{ SSSetUserColAction(m_hWnd, nUserColAction); }

#if defined(_WIN64) || defined(_IA64)
  LONG_PTR GetUserData()
#else
  long GetUserData()
#endif
        { return SSGetUserData(m_hWnd); }

#if defined(_WIN64) || defined(_IA64)
  LONG_PTR SetUserData(LONG_PTR lUserData)
#else
  long SetUserData(long lUserData)
#endif
        { return SSSetUserData(m_hWnd, lUserData); }

  WORD GetUserResize()
        { return SSGetUserResize(m_hWnd); }

  WORD SetUserResize(WORD wUserResize)
        { return SSSetUserResize(m_hWnd, wUserResize); }

  short GetUserResizeCol(SS_COORD Col)
        { return SSGetUserResizeCol(m_hWnd, Col); }

  short SetUserResizeCol(SS_COORD Col, short dOption)
        { return SSSetUserResizeCol(m_hWnd, Col, dOption); }

  short GetUserResizeRow(SS_COORD Row)
        { return SSGetUserResizeRow(m_hWnd, Row); }

  short SetUserResizeRow(SS_COORD Row, short dOption)
        { return SSSetUserResizeRow(m_hWnd, Row, dOption); }

  int GetValue(SS_COORD Col, SS_COORD Row, LPTSTR lpBuffer)
        { return SSGetValue (m_hWnd, Col, Row, lpBuffer); }

  BOOL SetValue(SS_COORD Col, SS_COORD Row, LPCTSTR lpData)
        { return SSSetValue (m_hWnd, Col, Row, lpData); }

  BOOL SetValueRange(SS_COORD Col,  SS_COORD Row,
                     SS_COORD Col2, SS_COORD Row2, LPCTSTR lpData)
        { return SSSetValueRange (m_hWnd, Col, Row, Col2, Row2, lpData); }

  int GetValueLen(SS_COORD Col, SS_COORD Row)
        { return SSGetValueLen (m_hWnd, Col, Row); }

  BOOL GetVisible(SS_COORD Col, SS_COORD Row, short Visible)
        { return SSGetVisible (m_hWnd, Col, Row, Visible); }

  BOOL ImportExcelBook (LPCTSTR lpszfileName, LPCTSTR lpszlogFileName)
        { return SSImportExcelBook (m_hWnd, lpszfileName, lpszlogFileName); }

  BOOL ImportExcelSheet(short nWorkbookHandle, short nSheetIndex, LPCTSTR lpszSheetName)
		{ return SSImportExcelSheet(m_hWnd, nWorkbookHandle, nSheetIndex, lpszSheetName); }

  BOOL InsCol(SS_COORD Col)
        { return SSInsCol (m_hWnd, Col); }

  BOOL InsColRange(SS_COORD Col, SS_COORD Col2)
        { return SSInsColRange(m_hWnd, Col, Col2); }

  BOOL InsertSheets(short nSheetIndex, short nSheetCnt)
        { return SSInsertSheets (m_hWnd, nSheetIndex, nSheetCnt); }

  BOOL InsRow(SS_COORD Row)
        { return SSInsRow (m_hWnd, Row); }

  BOOL InsRowRange(SS_COORD Row, SS_COORD Row2)
        { return SSInsRowRange(m_hWnd, Row, Row2); }

  BOOL IsCellInSelection(SS_COORD Col, SS_COORD Row)
        { return SSIsCellInSelection(m_hWnd, Col, Row); }

  BOOL IsColHidden(SS_COORD Col)
		{ return SSIsColHidden(m_hWnd, Col); }

  int IsExcelFile(LPCTSTR lpszFileName)
		{ return SSIsExcelFile (lpszFileName);}

  BOOL IsFetchCellNote()
		{ return SSIsFetchCellNote(m_hWnd); }

  BOOL IsRowHidden(SS_COORD Row)
		{ return SSIsRowHidden(m_hWnd, Row); }

  BOOL LoadFromBuffer(GLOBALHANDLE hBuffer, long lBufferLen)
        { return SSLoadFromBuffer (m_hWnd, hBuffer, lBufferLen); }

  BOOL LoadFromFile(LPCTSTR lpszFileName)
        { return SSLoadFromFile (m_hWnd, lpszFileName); }

  HANDLE LoadPicture(LPCTSTR lpszFileName, short nPictType)
        { return SSLoadPicture(lpszFileName, nPictType); }

  HANDLE LoadPictureBuffer(HGLOBAL hGlobal, long lSize, short nPictType)
        { return SSLoadPictureBuffer(hGlobal, lSize, nPictType); }

  HANDLE LoadResPicture(HINSTANCE hInstance, LPCTSTR lpszResName, LPCTSTR lpszResType, short nPictType)
        { return SSLoadResPicture(hInstance, lpszResName, lpszResType, nPictType); }

  BOOL LoadTabFile(LPCTSTR lpszFileName)
        { return SSLoadTabFile(m_hWnd, lpszFileName); }

  BOOL LoadTextFile (LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
		{ return SSLoadTextFile (m_hWnd, pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

  void LogUnitsToColWidth(short dUnits, LPDOUBLE lpWidth)
        { SSLogUnitsToColWidth(m_hWnd, dUnits, lpWidth); }

  void LogUnitsToRowHeight(SS_COORD Row, short dUnits, LPDOUBLE lpHeight)
        { SSLogUnitsToRowHeight(m_hWnd, Row, dUnits, lpHeight); }

  BOOL MoveRange(SS_COORD Col,SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
         SS_COORD ColDest, SS_COORD RowDest)
        { return SSMoveRange(m_hWnd, Col, Row, Col2, Row2, ColDest, RowDest); }

  BOOL MoveSheets(short nSheetSrc, short nSheetCnt, short nSheetDest)
        { return SSMoveSheets(m_hWnd, nSheetSrc, nSheetCnt, nSheetDest); }

#if defined(_WIN64) || defined(_IA64)
  BOOL Print(LPCTSTR lpszPrintJobName, LPSS_PRINTFORMAT lpPrintFormat,
             SS_PRINTFUNC lpfnPrintProc, LONG_PTR lAppData,
             ABORTPROC lpfnAbortProc)
#else
  BOOL Print(LPCTSTR lpszPrintJobName, LPSS_PRINTFORMAT lpPrintFormat,
             SS_PRINTFUNC lpfnPrintProc, LONG lAppData,
             ABORTPROC lpfnAbortProc)
#endif
        { return SSPrint(m_hWnd, lpszPrintJobName, lpPrintFormat,
                 lpfnPrintProc, lAppData, lpfnAbortProc); }
  
  BOOL OwnerPrintDraw(HDC hDC, long left, long top, long right, long bottom, short page)
        { return SSOwnerPrintDraw(m_hWnd, hDC, left, top, right, bottom, page);}

  BOOL OwnerPrintPageCount(HDC hDC, long left, long top, long right, long bottom, LPSHORT pageCount)
        { return SSOwnerPrintPageCount(m_hWnd, hDC, left, top, right, bottom, pageCount);}
  int QueryCustomName(LPCTSTR lpszPrevName, LPTSTR lpszName, int nNameLen)
        { return SSQueryCustomName(m_hWnd, lpszPrevName, lpszName, nNameLen); }

  int QueryCustomNameLen(LPCTSTR lpszPrevName)
        { return SSQueryCustomNameLen(m_hWnd, lpszPrevName); }

  BOOL ReCalc()
        { return SSReCalc (m_hWnd); }

  BOOL ReCalcCell(SS_COORD lCol, SS_COORD lRow)
  		{ return SSReCalcCell(m_hWnd, lCol, lRow); }

  void RemoveCellSpan(SS_COORD lCol, SS_COORD lRow)
	{ SSRemoveCellSpan(m_hWnd, lCol, lRow); }
  		
  BOOL RemoveCustomCellType(LPCTSTR lpszName)
  { return SSRemoveCustomCellType(m_hWnd, lpszName); }

  BOOL RemoveCustomFunction(LPCTSTR lpszFuncName)
  		{ return SSRemoveCustomFunction(m_hWnd, lpszFuncName); }

  BOOL ReplaceSel(LPCTSTR lpText)
        { return SSReplaceSel (m_hWnd, lpText); }

  void Reset()
        { SSReset(m_hWnd); }

  void ResetSheet(short nSheetIndex)
        { SSResetSheet(m_hWnd, nSheetIndex); }

  int RowHeightToLogUnits(SS_COORD Row, double dfRowHeight)
        { return SSRowHeightToLogUnits (m_hWnd, Row, dfRowHeight); }

  BOOL SavePicture(HANDLE hPict, HPALETTE hPal, LPCTSTR lpszFileName, short nSaveType)
        { return SSSavePicture(hPict, hPal, lpszFileName, nSaveType); }
  
  BOOL SavePictureBuffer(HANDLE hPict, HPALETTE hPal, short nSaveType, HGLOBAL *phGlobal, long *plSize)
        { return SSSavePictureBuffer(hPict, hPal, nSaveType, phGlobal, plSize); }

  BOOL SaveTabFile(LPCTSTR lpszFileName)
        { return SSSaveTabFile(m_hWnd, lpszFileName); }

  BOOL SaveTabFileU(LPCTSTR lpszFileName)
        { return SSSaveTabFileU(m_hWnd, lpszFileName); }

  GLOBALHANDLE SaveToBuffer(LPLONG lpBufferLen, BOOL fDataOnly)
        { return SSSaveToBuffer (m_hWnd, lpBufferLen, fDataOnly); }
                    
  BOOL SaveToFile(LPCTSTR lpszFileName, BOOL fDataOnly)
        { return SSSaveToFile (m_hWnd, lpszFileName, fDataOnly); }

  SS_COORD SearchCol(SS_COORD lCol, SS_COORD lRowStart, SS_COORD lRowEnd, LPCTSTR lpszText, WORD wSearchFlags)
	{ return SSSearchCol(m_hWnd, lCol, lRowStart, lRowEnd, lpszText, wSearchFlags); }

  SS_COORD SearchRow(SS_COORD lRow, SS_COORD lColStart, SS_COORD lColEnd, LPCTSTR lpszText, WORD wSearchFlags)
	{ return SSSearchRow(m_hWnd, lRow, lColStart, lColEnd, lpszText, wSearchFlags); }

#if defined(_WIN64) || defined(_IA64)
  long SelModeSendMessage(UINT uMsg, long lParam1, LONG_PTR lParam2, long lParam3)
#else
  long SelModeSendMessage(UINT uMsg, long lParam1, long lParam2, long lParam3)
#endif
        { return SSSelModeSendMessage(m_hWnd, uMsg, lParam1, lParam2, lParam3); }

  LPSS_CELLTYPE SetTypeButton(LPSS_CELLTYPE lpCellType, long Style,
                              LPCTSTR Text, LPCTSTR Picture, 
                              short nPictureType, LPCTSTR PictureDown,
                              short nPictureDownType, short ButtonType,
                              short nShadowSize, LPSUPERBTNCOLOR lpColor)
        { return SSSetTypeButton(m_hWnd, lpCellType, Style, 
                                 Text, Picture, nPictureType,
                                 PictureDown, nPictureDownType,
                                 ButtonType, nShadowSize, lpColor); }
 
  LPSS_CELLTYPE SetTypeCheckBox(LPSS_CELLTYPE lpCellType,
                long lStyle, LPCTSTR lpText,
                LPCTSTR lpPictUp, WORD wPictUpType,
                LPCTSTR lpPictDown, WORD wPictDownType,
                LPCTSTR lpPictFocusUp, WORD wPictFocusUpType,
                LPCTSTR lpPictFocusDown, WORD wPictFocusDownType,
                LPCTSTR lpPictGray, WORD wPictGrayType,
                LPCTSTR lpPictFocusGray, WORD wPictFocusGrayType)
        { return SSSetTypeCheckBox(m_hWnd, lpCellType, lStyle, lpText, lpPictUp, wPictUpType,
                   lpPictDown, wPictDownType, lpPictFocusUp, wPictFocusUpType,
                   lpPictFocusDown, wPictFocusDownType, lpPictGray, wPictGrayType,
                   lpPictFocusGray, wPictFocusGrayType); }

  LPSS_CELLTYPE SetTypeComboBox(LPSS_CELLTYPE lpCellType, long Style, LPCTSTR lpItems)
        { return SSSetTypeComboBox(m_hWnd, lpCellType, Style, lpItems); }
 

  LPSS_CELLTYPE SetTypeComboBoxEx(LPSS_CELLTYPE lpCellType, long Style, 
  								  LPCTSTR lpItems, short dMaxRows, short dMaxEditLen,
                                  short dComboWidth, HWND hWndDropDown, BYTE bAutoSearch)
        { return SSSetTypeComboBoxEx(m_hWnd, lpCellType, Style, 
  								     lpItems, dMaxRows, dMaxEditLen,
                                     dComboWidth, hWndDropDown, bAutoSearch); }

  LPSS_CELLTYPE SetTypeCurrency(LPSS_CELLTYPE lpCellType, long style, BYTE Right, double Min, double Max, BYTE fShowSeparator)
	{ return SSSetTypeCurrency(m_hWnd, lpCellType, style, Right, Min, Max, fShowSeparator); }

  LPSS_CELLTYPE SetTypeCurrencyEx(LPSS_CELLTYPE lpCellType, long Style, BYTE Right, double Min, double Max, BYTE fShowSeparator, BYTE fShowCurrencySymbol, LPTSTR lpszDecimal, LPTSTR lpszSeparator, LPTSTR lpszCurrency, BYTE bLeadingZero, BYTE bNegCurrencyStyle, BYTE bPosCurrencyStyle, BYTE fSpin, BYTE fSpinWrap, double SpinInc)
	{ return SSSetTypeCurrencyEx(m_hWnd, lpCellType, Style, Right, Min, Max, fShowSeparator, fShowCurrencySymbol, lpszDecimal, lpszSeparator, lpszCurrency, bLeadingZero, bNegCurrencyStyle, bPosCurrencyStyle, fSpin, fSpinWrap, SpinInc); } 

 LPSS_CELLTYPE  SetTypeCustom(LPSS_CELLTYPE lpCellType, LPCTSTR lpszName, long Style, LPSS_CT_VALUE lpItemData)
	{ return SSSetTypeCustom(m_hWnd, lpCellType, lpszName, Style, lpItemData);}

 LPSS_CELLTYPE SetTypeDate(LPSS_CELLTYPE lpCellType, long Style,
                            LPSS_DATEFORMAT lpFormat, LPSS_DATE lpMin, 
                            LPSS_DATE lpMax)
        { return SSSetTypeDate(m_hWnd, lpCellType, Style, lpFormat, lpMin, lpMax); }

  LPSS_CELLTYPE SetTypeEdit(LPSS_CELLTYPE lpCellType, long Style, short Len,
                            short ChrSet, short ChrCase)
        { return SSSetTypeEdit(m_hWnd, lpCellType, Style, Len, ChrSet, ChrCase); }

  LPSS_CELLTYPE SetTypeFloat(LPSS_CELLTYPE lpCellType, long Style,
                             short Left, short Right, double Min, 
                             double Max)
        { return SSSetTypeFloat(m_hWnd, lpCellType, Style, Left, Right, Min, Max); }

  LPSS_CELLTYPE SetTypeFloatExt(LPSS_CELLTYPE lpCellType, long Style,
                             short Left, short Right, double Min, 
                             double Max, LPSS_FLOATFORMAT lpFormat)
        { return SSSetTypeFloatExt(m_hWnd, lpCellType, Style, Left, Right, Min, Max, lpFormat); }

  LPSS_CELLTYPE SetTypeInteger(LPSS_CELLTYPE lpCellType, long Min, long Max)
        { return SSSetTypeInteger(m_hWnd, lpCellType, Min, Max); }

  LPSS_CELLTYPE SetTypeIntegerExt(LPSS_CELLTYPE lpCellType, long lStyle, long lMin,
                  long lMax, BOOL fSpinWrap, long lSpinInc)
        { return SSSetTypeIntegerExt(m_hWnd, lpCellType, lStyle, lMin, lMax, fSpinWrap, lSpinInc); }

  LPSS_CELLTYPE SetTypeNumber(LPSS_CELLTYPE lpCellType, long lStyle, BYTE Right, double Min, double Max, BYTE fShowSeparator)
	{ return SSSetTypeNumber(m_hWnd, lpCellType, lStyle, Right, Min, Max, fShowSeparator); }

  LPSS_CELLTYPE SetTypeNumberEx(LPSS_CELLTYPE lpCellType, long lStyle, BYTE Right, double Min, double Max, BYTE fShowSeparator, LPTSTR lpszDecimal, LPTSTR lpszSeparator, BYTE bLeadingZero, BYTE bNegStyle, BYTE fSpin, BYTE fSpinWrap, double SpinInc)
	{ return SSSetTypeNumberEx(m_hWnd, lpCellType, lStyle, Right, Min, Max, fShowSeparator, lpszDecimal, lpszSeparator, bLeadingZero, bNegStyle, fSpin, fSpinWrap, SpinInc); } 							 
 
  LPSS_CELLTYPE SetTypeOwnerDraw(LPSS_CELLTYPE lpCellType, long lStyle)
        { return SSSetTypeOwnerDraw(m_hWnd, lpCellType, lStyle); }

  LPSS_CELLTYPE SetTypePercent(LPSS_CELLTYPE lpCellType, long Style, BYTE Right, double Min, double Max)
	{ return SSSetTypePercent(m_hWnd, lpCellType, Style, Right, Min, Max); }

  LPSS_CELLTYPE SetTypePercentEx(LPSS_CELLTYPE lpCellType, long Style, BYTE Right, double Min, double Max, LPTSTR lpszDecimal,  BYTE bLeadingZero, BYTE bNegStyle, BYTE fSpin, BYTE fSpinWrap, double SpinInc)
	{ return SSSetTypePercentEx(m_hWnd, lpCellType, Style, Right, Min, Max, lpszDecimal, bLeadingZero, bNegStyle, fSpin, fSpinWrap, SpinInc); }

  LPSS_CELLTYPE SetTypePic(LPSS_CELLTYPE lpCellType, long Style, LPCTSTR Mask)
        { return SSSetTypePic(m_hWnd, lpCellType, Style, Mask); }

  LPSS_CELLTYPE SetTypePicture(LPSS_CELLTYPE lpCellType, long Style,
                               LPCTSTR PictName)
        { return SSSetTypePicture(m_hWnd, lpCellType, Style, PictName); }

  LPSS_CELLTYPE SetTypePictureHandle(LPSS_CELLTYPE lpCellType, long Style, HANDLE hPict,
                     HPALETTE hPal, BOOL fDeleteHandle)
        { return SSSetTypePictureHandle(m_hWnd, lpCellType, Style, hPict, hPal, fDeleteHandle); }

  LPSS_CELLTYPE SetTypeScientific(LPSS_CELLTYPE lpCellType, long lStyle, BYTE Right, double Min, double Max, LPTSTR lpszDecimal)
	{ return SSSetTypeScientific(m_hWnd, lpCellType, lStyle, Right, Min, Max, lpszDecimal); }

  LPSS_CELLTYPE SetTypeStaticText(LPSS_CELLTYPE lpCellType, long TextStyle)
        { return SSSetTypeStaticText(m_hWnd, lpCellType, TextStyle); }

  LPSS_CELLTYPE SetTypeTime(LPSS_CELLTYPE lpCellType, long Style,
                            LPSS_TIMEFORMAT lpFormat, LPSS_TIME lpMin,
                            LPSS_TIME lpMax)
        { return SSSetTypeTime (m_hWnd, lpCellType, Style, lpFormat, lpMin, lpMax); }

  BOOL ShowActiveCell(short Position)
        { return SSShowActiveCell (m_hWnd, Position); }

  BOOL ShowCell(SS_COORD Col, SS_COORD Row, short Position)
        { return SSShowCell (m_hWnd, Col, Row, Position); }

  BOOL ShowCol(SS_COORD Col, BOOL fShow)
        { return SSShowCol(m_hWnd, Col, fShow); }

  BOOL ShowRow(SS_COORD Row, BOOL fShow)
        { return SSShowRow(m_hWnd, Row, fShow); }

  BOOL Sort(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, LPSS_SORT lpSort)
        { return SSSort(m_hWnd, Col, Row, Col2, Row2, lpSort); }

  BOOL SortEx(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
				WORD wSortBy, LPSS_SORTKEY lpSortKeys, short nSortKeyCnt)
		{ return SSSortEx(m_hWnd, Col, Row, Col2, Row2, wSortBy, lpSortKeys, nSortKeyCnt); }

  BOOL SwapRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
         SS_COORD ColDest, SS_COORD RowDest)
        { return SSSwapRange(m_hWnd, Col, Row, Col2, Row2, ColDest, RowDest); }

  BOOL ValidateFormula (LPCTSTR lpszFormula)
        { return SSValidateFormula (m_hWnd, lpszFormula); }

  void VGetBufferSize(LPSS_COORD lpdBufferSize, LPSS_COORD lpdOverlap)
        { SSVGetBufferSize(m_hWnd, lpdBufferSize, lpdOverlap); }

  void VSetBufferSize(SS_COORD dBufferSize, SS_COORD dOverlap)
        { SSVSetBufferSize(m_hWnd, dBufferSize, dOverlap); }

  SS_COORD VGetMax()
        { return SSVGetMax(m_hWnd); }

  SS_COORD VSetMax(SS_COORD Max)
        { return SSVSetMax(m_hWnd, Max); }

  SS_COORD VGetPhysBufferSize(void)
        { return SSVGetPhysBufferSize(m_hWnd); }

  SS_COORD VGetPhysBufferTop(void)
        { return SSVGetPhysBufferTop(m_hWnd); }

  LONG VGetStyle(void)
        { return SSVGetStyle(m_hWnd); }

  LONG VSetStyle(LONG lStyle)
        { return SSVSetStyle(m_hWnd, lStyle); }
 
  BOOL VRefreshBuffer()
        { return SSVRefreshBuffer(m_hWnd); }

  BOOL VScrollGetSpecial(LPWORD lpwOptions)
        { return SSVScrollGetSpecial(m_hWnd, lpwOptions); }

  BOOL VScrollSetSpecial(BOOL fUseSpecialVScroll, WORD wOptions)
        { return SSVScrollSetSpecial(m_hWnd, fUseSpecialVScroll, wOptions); }


  WORD GetTabStripButtonPolicy(void)
		{return SSGetTabStripButtonPolicy(m_hWnd);}

  WORD SetTabStripButtonPolicy(WORD wTabStripButtonPolicy)
		{return SSSetTabStripButtonPolicy(m_hWnd, wTabStripButtonPolicy);}

  BOOL IsVerticalScrollBarVisible(void)
		{return SSIsVerticalScrollBarVisible(m_hWnd);}

  BOOL IsHorizontalScrollBarVisible(void)
		{return SSIsHorizontalScrollBarVisible(m_hWnd);}

  WORD SetAppearanceStyle(WORD nStyle)
		{return SSSetAppearanceStyle(m_hWnd, nStyle);}

  WORD GetAppearanceStyle(void)
		{return SSGetAppearanceStyle(m_hWnd);}

  WORD SetScrollBarStyle(WORD nStyle)
		{return SSSetScrollBarStyle(m_hWnd, nStyle);}

  WORD GetScrollBarStyle(void)
		{return SSGetScrollBarStyle(m_hWnd);}

  WORD SetUseVisualStyles(WORD nStyle)
		{return SSSetUseVisualStyles(m_hWnd, nStyle);}

  WORD GetUseVisualStyles(void)
		{return SSGetUseVisualStyles(m_hWnd);}
  
  BOOL SetEnhanceStaticCells(BOOL nStyle)
		{return SSSetEnhanceStaticCells(m_hWnd, nStyle);}

  BOOL GetEnhanceStaticCells(void)
		{return SSGetEnhanceStaticCells(m_hWnd);}

  WORD SetTabEnhancedShape(WORD nStyle)
		{return SSSetTabEnhancedShape(m_hWnd, nStyle);}

  WORD GetTabEnhancedShape(void)
		{return SSGetTabEnhancedShape(m_hWnd);}

  WORD GetCellNoteIndicatorShape(void)
		{return SSGetCellNoteIndicatorShape(m_hWnd);}

  WORD SetCellNoteIndicatorShape(WORD nStyle)
		{return SSSetCellNoteIndicatorShape(m_hWnd, nStyle);}

  COLORREF GetCellNoteIndicatorColor(void)
		{return SSGetCellNoteIndicatorColor(m_hWnd);}

  COLORREF SetCellNoteIndicatorColor(COLORREF Color)
		{return SSSetCellNoteIndicatorColor(m_hWnd, Color);}

  BOOL SetEnhancedCornerColors(COLORREF clrBackColor, COLORREF clrHoverColor, COLORREF clrTriangle, COLORREF clrTriangleHover, COLORREF clrTriangleBorderColor, COLORREF clrTriangleHoverBorderColor)
		{return SSSetEnhancedCornerColors(m_hWnd, clrBackColor, clrHoverColor, clrTriangle, clrTriangleHover, clrTriangleBorderColor, clrTriangleHoverBorderColor);}

  BOOL SetEnhancedColumnHeaderColors(COLORREF clrSelectedHoverUpperColor, COLORREF clrSelectedHoverLowerColor, COLORREF clrSelectedUpperColor,
											  COLORREF clrSelectedLowerColor, COLORREF clrHoverUpperColor, COLORREF clrHoverLowerColor, COLORREF clrUpperColor,
											  COLORREF clrLowerColor, COLORREF clrSelectedBorderColor, COLORREF clrBorderColor)
		{return SSSetEnhancedColumnHeaderColors(m_hWnd, clrSelectedHoverUpperColor, clrSelectedHoverLowerColor, clrSelectedUpperColor,
											  clrSelectedLowerColor, clrHoverUpperColor, clrHoverLowerColor, clrUpperColor,
											  clrLowerColor, clrSelectedBorderColor, clrBorderColor);}


  BOOL SetEnhancedRowHeaderColors(COLORREF clrSelectedHoverUpperColor, COLORREF clrSelectedHoverLowerColor, COLORREF clrSelectedUpperColor,
											  COLORREF clrSelectedLowerColor, COLORREF clrHoverUpperColor, COLORREF clrHoverLowerColor, COLORREF clrUpperColor,
											  COLORREF clrLowerColor, COLORREF clrSelectedBorderColor, COLORREF clrBorderColor)
		{return SSSetEnhancedRowHeaderColors(m_hWnd, clrSelectedHoverUpperColor, clrSelectedHoverLowerColor, clrSelectedUpperColor,
											  clrSelectedLowerColor, clrHoverUpperColor, clrHoverLowerColor, clrUpperColor,
											  clrLowerColor, clrSelectedBorderColor, clrBorderColor);}

  BOOL SetEnhancedSheetTabColors(COLORREF clrNormalUpperStartColor, COLORREF clrNormalUpperEndColor,
								   COLORREF clrNormalLowerStartColor, COLORREF clrNormalLowerEndColor, COLORREF clrHoverUpperStartColor,
								   COLORREF clrHoverUpperEndColor, COLORREF clrHoverLowerStartColor, COLORREF clrHoverLowerEndColor,
								   COLORREF clrTabOuterBorderColor, COLORREF clrTabInnerBorderColor, COLORREF clrButtonUpperNormalStartColor,
								   COLORREF clrButtonUpperNormalEndColor, COLORREF clrButtonLowerNormalStartColor, COLORREF clrButtonLowerNormalEndColor,
								   COLORREF clrButtonUpperHoverStartColor, COLORREF clrButtonUpperHoverEndColor, COLORREF clrButtonLowerHoverStartColor,
								   COLORREF clrButtonLowerHoverEndColor,COLORREF clrButtonUpperPushedStartColor, COLORREF clrButtonUpperPushedEndColor,
								   COLORREF clrButtonLowerPushedStartColor, COLORREF clrButtonLowerPushedEndColor, COLORREF clrArrowColor,
								   COLORREF clrBackColorStartColor, COLORREF clrBackColorEndColor, COLORREF clrSplitBoxStartColor,
								   COLORREF clrSplitBoxEndColor, COLORREF clrForeColor, COLORREF clrActiveForeColor)
		{return SSSetEnhancedSheetTabColors(m_hWnd, clrNormalUpperStartColor, clrNormalUpperEndColor,
								   clrNormalLowerStartColor, clrNormalLowerEndColor, clrHoverUpperStartColor,
								   clrHoverUpperEndColor, clrHoverLowerStartColor, clrHoverLowerEndColor,
								   clrTabOuterBorderColor, clrTabInnerBorderColor, clrButtonUpperNormalStartColor,
								   clrButtonUpperNormalEndColor, clrButtonLowerNormalStartColor, clrButtonLowerNormalEndColor,
								   clrButtonUpperHoverStartColor, clrButtonUpperHoverEndColor, clrButtonLowerHoverStartColor,
								   clrButtonLowerHoverEndColor, clrButtonUpperPushedStartColor, clrButtonUpperPushedEndColor,
								   clrButtonLowerPushedStartColor, clrButtonLowerPushedEndColor, clrArrowColor,
								   clrBackColorStartColor, clrBackColorEndColor, clrSplitBoxStartColor,
								   clrSplitBoxEndColor, clrForeColor, clrActiveForeColor);}




  BOOL SetEnhancedScrollBarColors( COLORREF clrTrackColor, COLORREF clrArrowColor, COLORREF clrUpperNormalStartColor, COLORREF clrUpperNormalEndColor,
										   COLORREF clrLowerNormalStartColor, COLORREF clrLowerNormalEndColor, COLORREF clrUpperPushedStartColor, COLORREF clrUpperPushedEndColor,
										   COLORREF clrLowerPushedStartColor, COLORREF clrLowerPushedEndColor, COLORREF clrUpperHoverStartColor, COLORREF clrUpperHoverEndColor,
										   COLORREF clrLowerHoverStartColor, COLORREF clrLowerHoverEndColor, COLORREF clrHoverButtonBorderColor, COLORREF clrButtonBorderColor)
		{return SSSetEnhancedScrollBarColors(m_hWnd, clrTrackColor, clrArrowColor, clrUpperNormalStartColor, clrUpperNormalEndColor,
										   clrLowerNormalStartColor, clrLowerNormalEndColor, clrUpperPushedStartColor, clrUpperPushedEndColor,
										   clrLowerPushedStartColor, clrLowerPushedEndColor, clrUpperHoverStartColor, clrUpperHoverEndColor,
										   clrLowerHoverStartColor, clrLowerHoverEndColor, clrHoverButtonBorderColor, clrButtonBorderColor);}

  WORD SetHighlightHeaders(WORD wOption)
		{return SSSetHighlightHeaders(m_hWnd, wOption);}

  WORD GetHighlightHeaders(void)
		{return SSGetHighlightHeaders(m_hWnd);}

  WORD SetHighlightStyle(WORD wStyle)
		{return SSSetHighlightStyle(m_hWnd, wStyle);}

  WORD GetHighlightStyle(void)
		{return SSGetHighlightStyle(m_hWnd);}

  BOOL SetHighlightAlphaBlend(COLORREF clrAlphaBlend, short nAlphaBlend)
		{return SSSetHighlightAlphaBlend(m_hWnd, clrAlphaBlend, nAlphaBlend);}

  BOOL GetHighlightAlphaBlend(LPCOLORREF lpclrAlphaBlend, LPSHORT lpnAlphaBlend)
		{return SSGetHighlightAlphaBlend(m_hWnd, lpclrAlphaBlend, lpnAlphaBlend);}

  BOOL ExportToHTMLEx (LPCTSTR lpszFileName, 
							  BOOL bAppendFlag, LPCTSTR lpszLogFile, LPCTSTR lpszTitle, LPCTSTR lpszCssFile, LPCTSTR lpszTableClass, LPCTSTR lpszTableRowClass, LPCTSTR lpszTableDefinitionClass, LPCTSTR lpszTableHeaderClass)
		{return SSExportToHTMLEx (m_hWnd, lpszFileName, bAppendFlag, lpszLogFile, lpszTitle, lpszCssFile, lpszTableClass, lpszTableRowClass, lpszTableDefinitionClass, lpszTableHeaderClass);}

  BOOL ExportRangeToHTMLEx (SS_COORD col, SS_COORD row, SS_COORD col2, SS_COORD row2, 
								   LPCTSTR lpszFileName, BOOL bAppendFlag, 
								   LPCTSTR lpszLogFile, LPCTSTR lpszTitle,
								   LPCTSTR lpszCssFile, LPCTSTR lpszTableClass, LPCTSTR lpszTableRowClass, LPCTSTR lpszTableDefinitionClass, LPCTSTR lpszTableHeaderClass)
  {return SSExportRangeToHTMLEx (m_hWnd, col, row, col2, row2, lpszFileName, bAppendFlag, lpszLogFile, lpszTitle, lpszCssFile,
			lpszTableClass, lpszTableRowClass, lpszTableDefinitionClass, lpszTableHeaderClass);}
  BOOL SaveExcel2007File(LPCTSTR lpszFileName, LPCTSTR lpszPassword, short nFlags, LPCTSTR lpszLogFileName)
  {
     return SSSaveExcel2007File(m_hWnd, lpszFileName, lpszPassword, nFlags, lpszLogFileName);
  }
  BOOL OpenExcel2007File(LPCTSTR lpszFileName, LPCTSTR lpszPassword, short nSheet, short nExcelSheet, LPCTSTR lpszLogFileName)
  {
     return SSOpenExcel2007File(m_hWnd, lpszFileName, lpszPassword, nSheet, nExcelSheet, lpszLogFileName);
  }
  short IsExcel2007File(LPCTSTR lpszFileName)
  {
     return SSIsExcel2007File(m_hWnd, lpszFileName);
  }
  BOOL LoadBlockFromBuffer(SS_COORD lCol, SS_COORD lRow,
                           SS_COORD lCol2, SS_COORD lRow2,
                           LPBYTE lpBuff, long lBuffLen)
  {
     return SSLoadBlockFromBuffer(m_hWnd, lCol, lRow, lCol2, lRow2, lpBuff, lBuffLen);
  }
  HGLOBAL SaveBlockToBuffer(SS_COORD lCol, SS_COORD lRow,
                            SS_COORD lCol2, SS_COORD lRow2,
                            LPLONG lplBuffLen, BOOL  bDataOnly)
  {
     return SSSaveBlockToBuffer(m_hWnd, lCol, lRow, lCol2, lRow2, lplBuffLen, bDataOnly);
  }

  HGLOBAL SaveBlockToBufferExt(SS_COORD lCol, SS_COORD lRow,
                               SS_COORD lCol2, SS_COORD lRow2,
                               LPLONG lplBuffLen, BOOL  bDataOnly, BOOL fSaveColInfo, BOOL fSaveRowInfo)
  {
     return SSSaveBlockToBufferExt(m_hWnd, lCol, lRow, lCol2, lRow2, lplBuffLen, bDataOnly, fSaveColInfo, fSaveRowInfo);
  }
  
  // Implementation
  virtual ~TSpread() {};
  virtual WNDPROC* GetSuperWndProcAddr();

}; // end class TSpread

#define ConvertTSpread(RCID) ((TSpread*)GetDlgItem(RCID))
class TSpreadEx 
{
	public:
		TSpreadEx(){m_nCols=0,m_nRows=0;m_pSpread=NULL;m_nClickedCol=0,m_nClickedRow=0;m_nRCID=0;m_pParent=NULL;m_pParentReal=NULL;}
	CWnd* m_pParent;
	CWnd* m_pParentReal;
	int m_nRCID;
	
	BOOL SetWindowPos(const CWnd* pWndInsertAfter, int x, int y, int cx,int cy, UINT nFlags)
	{
		BOOL nflag;
		return nflag = m_pSpread->SetWindowPos(pWndInsertAfter, x, y, cx, cy, nFlags);
	}

	void SetNoBorder(BOOL bEnable)
	{
		m_pSpread->SetBool(SSB_NOBORDER,bEnable);
	}
	void Attach(TSpread* pWnd)
	{
		m_pParent = pWnd->GetParent();
		m_pSpread = pWnd;
	}

	void SetParentWnd(CWnd* pWnd,int nRC)
	{	
		m_pParentReal = pWnd;
		m_nRCID = nRC;
	}
	virtual ~TSpreadEx(){}
	SS_COORD m_nCols;
	SS_COORD m_nRows;

	SS_COORD GetRow()
	{
		return m_nRows;
	}

	SS_COORD GetCol()
	{
		return m_nCols;
	}

	SS_COORD m_nClickedCol;
	SS_COORD m_nClickedRow;

	TSpread* m_pSpread;

	BOOL AddCellSpan(SS_COORD lCol, SS_COORD lRow, SS_COORD lNumCols, SS_COORD lNumRows)
	{ 
		return m_pSpread->AddCellSpan(lCol,lRow,lNumCols,lNumRows);
	}

	BOOL AddCustomCellType(LPCTSTR lpszName, BOOL bEditable, BOOL bCanOverflow, BOOL bCanBeOverflown, BOOL bUseRendererControl, LPSS_CT_PROCS lpProcs)
	{ 
		return m_pSpread->AddCustomCellType(lpszName, bEditable, bCanOverflow, bCanBeOverflown, bUseRendererControl, lpProcs); 
	}

	BOOL AddCustomFunction(LPCTSTR FunctionName, short ParameterCnt,FARPROC lpfnFunctionProc)
	{ 
		return m_pSpread->AddCustomFunction(FunctionName, ParameterCnt,lpfnFunctionProc); 
	}

	BOOL AddCustomFunctionExt(LPCTSTR lpszFuncName, short nMinParamCnt,
		short nMaxParamCnt, FARPROC lpfnFuncProc,
		long lFlags)
	{ return m_pSpread->AddCustomFunctionExt(lpszFuncName, nMinParamCnt, nMaxParamCnt,lpfnFuncProc, lFlags); }

	BOOL AddMultiSelBlocks(LPSS_SELBLOCK lpSelBlock)
	{ return m_pSpread->AddMultiSelBlocks(lpSelBlock); }

	BOOL BuildDependencies()
	{ return m_pSpread->BuildDependencies(); }

	BOOL Clear(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->Clear(Col, Row); }

	BOOL ClearData(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->ClearData( Col, Row); }

	BOOL ClearDataRange(SS_COORD Col,  SS_COORD Row, 
		SS_COORD Col2, SS_COORD Row2)
	{ return m_pSpread->ClearDataRange(Col, Row, Col2, Row2); }

	BOOL ClearRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2)
	{ return m_pSpread->ClearRange( Col, Row, Col2, Row2); }

	BOOL ClipboardCopy(void)
	{ return m_pSpread->ClipboardCopy(); }

	BOOL ClipboardCut(void)
	{ return m_pSpread->ClipboardCut(); }

	BOOL ClipboardPaste(void)
	{ return m_pSpread->ClipboardPaste(); }

	BOOL ClipIn(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
		LPCTSTR Buffer, long BufferLen)
	{ return m_pSpread->ClipIn(Col, Row, Col2, Row2, Buffer, BufferLen); }

	GLOBALHANDLE ClipOut(SS_COORD Col, SS_COORD Row,
		SS_COORD Col2,SS_COORD Row2)
	{ return m_pSpread->ClipOut(Col, Row, Col2, Row2); }

	BOOL ClipValueIn (SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
		LPCTSTR Buffer, long BufferLen)
	{ return m_pSpread->ClipValueIn(Col,Row,Col2,Row2,Buffer,BufferLen); }

	GLOBALHANDLE ClipValueOut (SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2)
	{ return m_pSpread->ClipValueOut( Col,Row,Col2,Row2); }

	SS_COORD ColLetterToNumber(LPCTSTR lpszText)
	{ return m_pSpread->ColLetterToNumber(lpszText); }

	int ColWidthToLogUnits (double dfColWidth)
	{ return m_pSpread->ColWidthToLogUnits( dfColWidth); }

	LRESULT ComboBoxSendMessage (SS_COORD Col, SS_COORD Row, UINT Msg, WPARAM wParam, LPARAM lParam)
	{ return m_pSpread->ComboBoxSendMessage( Col, Row, Msg, wParam, lParam); }

	BOOL CopyRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2,
		SS_COORD Row2, SS_COORD ColDest, SS_COORD RowDest)
	{ return m_pSpread->CopyRange( Col, Row, Col2, Row2, ColDest, RowDest); }

	BOOL DelCol(SS_COORD Col)
	{ return m_pSpread->DelCol( Col); }

	BOOL DelColRange(SS_COORD Col, SS_COORD Col2)
	{ return m_pSpread->DelColRange( Col, Col2); }

	BOOL  DeleteSheets(short nSheetIndex, short nSheetCnt)
	{ return m_pSpread->DeleteSheets( nSheetIndex, nSheetCnt); }

	BOOL DelRow(SS_COORD Row)
	{ return m_pSpread->DelRow( Row); }

	BOOL DelRowRange(SS_COORD Row, SS_COORD Row2)
	{ return m_pSpread->DelRowRange( Row, Row2); }

	BOOL DeSelectBlock(void)
	{ return m_pSpread->DeSelectBlock(); }

	BOOL EnumCustomCellType(LPCTSTR lpszPrevName, LPTSTR lpszName, int nLen)
	{ return m_pSpread->EnumCustomCellType( lpszPrevName, lpszName, nLen); }

	BOOL EnumCustomCellTypeLen(LPCTSTR lpszPrevName)
	{ return m_pSpread->EnumCustomCellTypeLen( lpszPrevName); }

	int  EnumCustomFunction(LPCTSTR lpszPrevFunc, LPTSTR lpszFunc, int nFuncLen)
	{ return m_pSpread->EnumCustomFunction( lpszPrevFunc, lpszFunc, nFuncLen); }

	int  EnumCustomFunctionLen(LPCTSTR lpszPrevFunc)
	{ return m_pSpread->EnumCustomFunctionLen( lpszPrevFunc); }

	BOOL ExportExcelBook (LPCTSTR lpszfileName, LPCTSTR lpszlogFileName)
	{ return m_pSpread->ExportExcelBook ( lpszfileName, lpszlogFileName); }

	BOOL ExportExcelBookEx (LPCTSTR lpszfileName, LPCTSTR lpszlogFileName, SHORT nFlags)
	{ return m_pSpread->ExportExcelBookEx ( lpszfileName, lpszlogFileName, nFlags); }

	BOOL ExportRangeToHTML (SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, LPCTSTR lpszFileName, BOOL bAppendFlag, LPCTSTR lpszLogFile)
	{ return m_pSpread->ExportRangeToHTML ( Col, Row, Col2, Row2, lpszFileName, bAppendFlag, lpszLogFile); }

	BOOL ExportRangeToTextFile (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
	{ return m_pSpread->ExportRangeToTextFile ( lCol, lRow, lCol2, lRow2, pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

	BOOL ExportRangeToTextFileU (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
	{ return m_pSpread->ExportRangeToTextFileU ( lCol, lRow, lCol2, lRow2, pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

	BOOL ExportRangeToXML (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR lpszFileName, LPCTSTR pszRoot, LPCTSTR pszCollection, long lFlags, LPCTSTR lpszLogFile)
	{ return m_pSpread->ExportRangeToXML ( lCol, lRow, lCol2, lRow2, lpszFileName, pszRoot, pszCollection, lFlags, lpszLogFile); }

	BOOL ExportRangeToXMLBuffer (SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2, LPCTSTR pszRoot, LPCTSTR pszCollection, HGLOBAL *hgBuff, long lFlags, LPCTSTR lpszLogFile)
	{ return m_pSpread->ExportRangeToXMLBuffer ( lCol, lRow, lCol2, lRow2, pszRoot, pszCollection, hgBuff, lFlags, lpszLogFile); }

	BOOL ExportToExcel(LPCTSTR lpszFileName, LPCTSTR lpszSheetName, LPCTSTR lpszLogFileName)
	{ return m_pSpread->ExportToExcel( lpszFileName, lpszSheetName, lpszLogFileName); }

	BOOL ExportToExcelEx(LPCTSTR lpszFileName, LPCTSTR lpszSheetName, LPCTSTR lpszLogFileName, SHORT nFlags)
	{ return m_pSpread->ExportToExcelEx( lpszFileName, lpszSheetName, lpszLogFileName, nFlags); }

	BOOL ExportToHTML (LPCTSTR lpszFileName, BOOL bAppendFlag, LPCTSTR lpszLogFile)
	{ return m_pSpread->ExportToHTML ( lpszFileName, bAppendFlag, lpszLogFile); }

	BOOL ExportToTextFile (LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
	{ return m_pSpread->ExportToTextFile ( pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

	BOOL ExportToTextFileU (LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
	{ return m_pSpread->ExportToTextFileU ( pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

	BOOL ExportToXML (LPCTSTR lpszFileName, LPCTSTR pszRoot, LPCTSTR pszCollection, long lFlags, LPCTSTR lpszLogFile)
	{ return m_pSpread->ExportToXML ( lpszFileName, pszRoot, pszCollection, lFlags, lpszLogFile); }

	BOOL ExportToXMLBuffer (LPCTSTR pszRoot, LPCTSTR pszCollection, HGLOBAL *hgBuff, long lFlags, LPCTSTR lpszLogFile)
	{ return m_pSpread->ExportToXMLBuffer ( pszRoot, pszCollection, hgBuff, lFlags, lpszLogFile); }

	void FloatFormat(LPSS_CELLTYPE lpCellType, double dfVal, LPTSTR lpszText, BOOL fValue)
	{ m_pSpread->FloatFormat( lpCellType, dfVal, lpszText, fValue); }

	BOOL GetActionKey(WORD wAction, LPBOOL lpfShift, LPBOOL lpfCtrl, LPWORD lpwKey)
	{ return m_pSpread->GetActionKey( wAction, lpfShift, lpfCtrl, lpwKey); }

	BOOL SetActionKey(WORD wAction, BOOL fShift, BOOL fCtrl, WORD wKey)
	{ return m_pSpread->SetActionKey( wAction, fShift, fCtrl, wKey); }

	BOOL GetActiveCell (LPSS_COORD Col, LPSS_COORD Row)
	{ return m_pSpread->GetActiveCell( Col, Row); }

	BOOL SetActiveCell(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->SetActiveCell ( Col, Row); }

	WORD GetActiveCellHighlightStyle ()
	{ return m_pSpread->GetActiveCellHighlightStyle(); }

	void SetActiveCellHighlightStyle(WORD wActiveCellHighlightStyle)
	{ m_pSpread->SetActiveCellHighlightStyle ( wActiveCellHighlightStyle); }

	short GetActiveSheet ()
	{ return m_pSpread->GetActiveSheet(); }

	BOOL SetActiveSheet(short nSheet)
	{ return m_pSpread->SetActiveSheet ( nSheet); }

	WORD GetAppearance()
	{ return m_pSpread->GetAppearance(); }

	WORD SetAppearance(WORD wAppearance)
	{ return m_pSpread->SetAppearance( wAppearance); }

	BOOL GetArray(LPVOID lpArray, SS_COORD ColLeft,
		SS_COORD RowTop, SS_COORD ArrayColCnt,
		SS_COORD ArrayRowCnt, WORD wDataType)
	{ return m_pSpread->GetArray( lpArray, ColLeft, RowTop, ArrayColCnt, ArrayRowCnt, wDataType); }

	BOOL SetArray(LPVOID lpArray, SS_COORD ColLeft,
		SS_COORD RowTop, SS_COORD ArrayColCnt,
		SS_COORD ArrayRowCnt, WORD wDataType)
	{ return m_pSpread->SetArray( lpArray, ColLeft, RowTop, ArrayColCnt, ArrayRowCnt, wDataType); }

	void GetAutoSizeVisible (LPSS_COORD lpVisibleCols, LPSS_COORD lpVisibleRows)
	{ m_pSpread->GetAutoSizeVisible( lpVisibleCols, lpVisibleRows); }

	void SetAutoSizeVisible(SS_COORD VisibleCols, SS_COORD VisibleRows)
	{ m_pSpread->SetAutoSizeVisible ( VisibleCols, VisibleRows); }

	WORD GetBackColorStyle()
	{ return m_pSpread->GetBackColorStyle(); }

	WORD SetBackColorStyle(WORD wStyle)
	{ return m_pSpread->SetBackColorStyle( wStyle); }

	BOOL GetBool(short nIndex)
	{ return m_pSpread->GetBool( nIndex); }

	BOOL SetBool(short nIndex, BOOL bNewVal)
	{ return m_pSpread->SetBool ( nIndex, bNewVal); }

	BOOL GetBorder(SS_COORD Col, SS_COORD Row, LPWORD lpwBorderType,
		LPWORD lpwBorderStyle, LPCOLORREF lpColor)
	{ return m_pSpread->GetBorder( Col, Row, lpwBorderType, lpwBorderStyle, lpColor); }

	BOOL SetBorder(SS_COORD Col, SS_COORD Row, WORD wBorderType,
		WORD wBorderStyle, COLORREF Color)
	{ return m_pSpread->SetBorder( Col, Row, wBorderType, wBorderStyle, Color); }

	BOOL SetBorderRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2,
		SS_COORD Row2, WORD wBorderType, WORD wBorderStyle, COLORREF Color)
	{ return m_pSpread->SetBorderRange( Col, Row, Col2, Row2, wBorderType, wBorderStyle, Color); }

	BOOL GetBorderEx(SS_COORD Col, SS_COORD Row, 
		LPWORD lpwStyleLeft,	LPCOLORREF lpclrLeft, 
		LPWORD lpwStyleTop,		LPCOLORREF lpclrTop,
		LPWORD lpwStyleRight,	LPCOLORREF lpclrRight,
		LPWORD lpwStyleBottom,	LPCOLORREF lpclrBottom)
	{ return m_pSpread->GetBorderEx( Col, Row, lpwStyleLeft, lpclrLeft, lpwStyleTop, lpclrTop,
	lpwStyleRight, lpclrRight, lpwStyleBottom, lpclrBottom); }

	BOOL GetBottomRightCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
	{ return m_pSpread->GetBottomRightCell( lpCol, lpRow); }

	WORD GetButtonDrawMode(void)
	{ return m_pSpread->GetButtonDrawMode(); }

	WORD SetButtonDrawMode(WORD wMode)
	{ return m_pSpread->SetButtonDrawMode( wMode); }

	int GetButtonText(SS_COORD Col, SS_COORD Row, LPTSTR lpText)
	{ return m_pSpread->GetButtonText( Col, Row, lpText); }

	BOOL SetButtonText(SS_COORD Col, SS_COORD Row, LPCTSTR lpText)
	{ return m_pSpread->SetButtonText( Col, Row, lpText); }

	WNDPROC GetCallBack(void)
	{ return m_pSpread->GetCallBack(); }

	WNDPROC SetCallBack(WNDPROC lpfnCallBack)
	{ return m_pSpread->SetCallBack( lpfnCallBack); }

	void SetCalText(LPCTSTR lpszShortDays,		LPCTSTR lpszLongDays, 
		LPCTSTR lpszShortMonths,	LPCTSTR lpszLongMonths, 
		LPCTSTR lpszOkText,			LPCTSTR lpszCancelText)
	{ m_pSpread->SetCalText(lpszShortDays, lpszLongDays, lpszShortMonths, lpszLongMonths, lpszOkText, lpszCancelText); }

	void GetCalTextOverride(LPTSTR lpszShortDays, LPSHORT lpnLenShortDays, 
		LPTSTR lpszLongDays, LPSHORT lpnLenLongDays,
		LPTSTR lpszShortMonths, LPSHORT lpnLenShortMonths, 
		LPTSTR lpszLongMonths, LPSHORT lpnLenLongMonths,
		LPTSTR lpszOkText, LPSHORT lpnLenOkText, 
		LPTSTR lpszCancelText, LPSHORT lpnLenCancelText)
	{ m_pSpread->GetCalTextOverride( lpszShortDays, lpnLenShortDays, 
	lpszLongDays, lpnLenLongDays,
	lpszShortMonths, lpnLenShortMonths, 
	lpszLongMonths, lpnLenLongMonths,
	lpszOkText, lpnLenOkText, 
	lpszCancelText, lpnLenCancelText); }

	void SetCalTextOverride(LPCTSTR lpszShortDays, LPCTSTR lpszLongDays,
		LPCTSTR lpszShortMonths, LPCTSTR lpszLongMonths,
		LPCTSTR lpszOkText, LPCTSTR lpszCancelText) 
	{ m_pSpread->SetCalTextOverride( lpszShortDays, lpszLongDays, lpszShortMonths,
	lpszLongMonths, lpszOkText, lpszCancelText); }

	BOOL GetCellFromPixel(LPSS_COORD lpCol, LPSS_COORD lpRow,
		int MouseX, int MouseY)
	{ return m_pSpread->GetCellFromPixel( lpCol, lpRow, MouseX, MouseY); }

	short GetCellNote(SS_COORD Col, SS_COORD Row, LPTSTR Note)
	{ return m_pSpread->GetCellNote( Col, Row, Note); }

	BOOL SetCellNote(SS_COORD Col, SS_COORD Row, LPTSTR Note)
	{ return m_pSpread->SetCellNote( Col, Row, Note); }

	BOOL SetCellNoteRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, LPTSTR Note)
	{ return m_pSpread->SetCellNoteRange( Col, Row, Col2, Row2, Note); }

	WORD GetCellNoteIndicator()
	{ return m_pSpread->GetCellNoteIndicator(); }

	void SetCellNoteIndicator(WORD wValue)
	{ m_pSpread->SetCellNoteIndicator( wValue); }

	BOOL GetCellRect(SS_COORD Col, SS_COORD Row, LPRECT lpRect)
	{ return m_pSpread->GetCellRect( Col, Row, lpRect); }

	void GetCellSendingMsg(LPSS_COORD lpCol, LPSS_COORD lpRow)
	{ m_pSpread->GetCellSendingMsg( lpCol, lpRow); }

	WORD GetCellSpan(SS_COORD lCol, SS_COORD lRow, LPSS_COORD lplColAnchor, LPSS_COORD lplRowAnchor, 	LPSS_COORD lplNumCols, LPSS_COORD lplNumRows)
	{ return m_pSpread->GetCellSpan( lCol, lRow, lplColAnchor, lplRowAnchor, lplNumCols, 		lplNumRows);}

	short GetCellTag(SS_COORD Col, SS_COORD Row, LPTSTR CellTag)
	{ return m_pSpread->GetCellTag( Col, Row, CellTag); }

	BOOL SetCellTag(SS_COORD Col, SS_COORD Row, LPCTSTR CellTag)
	{ return m_pSpread->SetCellTag( Col, Row, CellTag); }

	BOOL GetCellType(SS_COORD Col, SS_COORD Row, LPSS_CELLTYPE lpCellType)
	{ return m_pSpread->GetCellType( Col, Row, lpCellType); }

	BOOL SetCellType(SS_COORD Col, SS_COORD Row, LPSS_CELLTYPE lpCellType)
	{ return m_pSpread->SetCellType( Col, Row, lpCellType); }

	BOOL SetCellTypeRange(SS_COORD Col,  SS_COORD Row, SS_COORD Col2,
		SS_COORD Row2, LPSS_CELLTYPE lpCellType)
	{ return m_pSpread->SetCellTypeRange( Col, Row, Col2, Row2, lpCellType); }

	void GetClientRect(LPRECT lpRect)
	{ m_pSpread->GetClientRect( lpRect); }

	WORD GetClipboardOptions()
	{ return m_pSpread->GetClipboardOptions(); }

	WORD SetClipboardOptions(WORD wOptions)
	{ return m_pSpread->SetClipboardOptions( wOptions); }

	SS_COORD GetColFromID(LPCTSTR lpszColID)
	{ return m_pSpread->GetColFromID( lpszColID); }

	WORD GetColHeaderDisplay()
	{ return m_pSpread->GetColHeaderDisplay(); }

	WORD SetColHeaderDisplay(WORD wDisplay)
	{ return m_pSpread->SetColHeaderDisplay( wDisplay); }

	SS_COORD GetColHeaderRows()
	{ return m_pSpread->GetColHeaderRows(); }

	void SetColHeaderRows(SS_COORD lColHeaderRows)
	{ m_pSpread->SetColHeaderRows( lColHeaderRows); }

	SS_COORD GetColHeadersAutoTextIndex()
	{ return m_pSpread->GetColHeadersAutoTextIndex(); }

	void SetColHeadersAutoTextIndex(SS_COORD lColHeadersAutoTextIndex)
	{ m_pSpread->SetColHeadersAutoTextIndex( lColHeadersAutoTextIndex); }

	SS_COORD GetColHeadersUserSortIndex()
	{ return m_pSpread->GetColHeadersUserSortIndex(); }

	void SetColHeadersUserSortIndex(SS_COORD lColHeadersUserSortIndex)
	{ m_pSpread->SetColHeadersUserSortIndex( lColHeadersUserSortIndex); }

	short GetColID(SS_COORD Col, LPTSTR ColID)
	{ return m_pSpread->GetColID( Col, ColID); }

	BOOL  SetColID(SS_COORD Col, LPCTSTR ColID)
	{ return m_pSpread->SetColID( Col, ColID); }

	short GetColMerge(SS_COORD Col)
	{ return m_pSpread->GetColMerge( Col); }

	void SetColMerge(SS_COORD Col, short nColMerge)
	{ m_pSpread->SetColMerge( Col, nColMerge); }

	BOOL GetColor(SS_COORD Col, SS_COORD Row, LPCOLORREF lpBackground,
		LPCOLORREF lpForeground)
	{ return m_pSpread->GetColor( Col, Row, lpBackground, lpForeground); }

	BOOL SetColor(SS_COORD Col, SS_COORD Row, COLORREF Background,
		COLORREF Foreground)
	{ return m_pSpread->SetColor( Col, Row, Background, Foreground); }

	BOOL SetColorRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
		COLORREF Background, COLORREF Foreground)
	{ return m_pSpread->SetColorRange( Col, Row, Col2, Row2, Background, Foreground); }

	BOOL GetColPageBreak(SS_COORD col)
	{ return m_pSpread->GetColPageBreak( col); }

	BOOL SetColPageBreak(SS_COORD col, BOOL fpagebreak)
	{ return m_pSpread->SetColPageBreak( col, fpagebreak); }

#if defined(_WIN64) || defined(_IA64)
	BOOL GetColUserData(SS_COORD Col,LONG_PTR* lplUserData)
#else 
	BOOL GetColUserData(SS_COORD Col,long* lplUserData)
#endif
	{ return m_pSpread->GetColUserData( Col, lplUserData); }

#if defined(_WIN64) || defined(_IA64)
	BOOL SetColUserData(SS_COORD Col, LONG_PTR lUserData)
#else
	BOOL SetColUserData(SS_COORD Col, long lUserData)
#endif
	{ return m_pSpread->SetColUserData( Col, lUserData); }

	short GetColUserSortIndicator(SS_COORD Col)
	{ return m_pSpread->GetColUserSortIndicator( Col); }

	void SetColUserSortIndicator(SS_COORD Col, short nColUserSortIndicator)
	{ m_pSpread->SetColUserSortIndicator( Col, nColUserSortIndicator); }

	BOOL GetColWidth(SS_COORD Col, LPDOUBLE lpColWidth)
	{ return m_pSpread->GetColWidth( Col, lpColWidth); }

	BOOL SetColWidth(SS_COORD Col, double Width)
	{ return m_pSpread->SetColWidth( Col, Width); }

	BOOL GetColWidthInPixels(SS_COORD Col, LPINT lpdWidthPixels)
	{ return m_pSpread->GetColWidthInPixels( Col, lpdWidthPixels); }

	BOOL SetColWidthInPixels(SS_COORD Col, int dWidth)
	{ return m_pSpread->SetColWidthInPixels( Col, dWidth); }

	BOOL GetCurrSelBlockPos(LPSS_COORD CurrentCol, LPSS_COORD CurrentRow)
	{ return m_pSpread->GetCurrSelBlockPos( CurrentCol, CurrentRow); }

	HCURSOR GetCursor(WORD wType)
	{ return m_pSpread->GetCursor( wType); }

	BOOL SetCursor(WORD wType, HCURSOR hCursor)
	{ return m_pSpread->SetCursor( wType, hCursor); }

	BOOL GetCustomCellType(LPCTSTR lpszName, LPBOOL lpbEditable, LPBOOL lpbCanOverflow, LPBOOL lpbCanBeOverflown, LPBOOL lpbUseRendererControl, LPSS_CT_PROCS lpProcs)
	{ return m_pSpread->GetCustomCellType( lpszName, lpbEditable, lpbCanOverflow, lpbCanBeOverflown, lpbUseRendererControl, lpProcs); }

	BOOL GetCustomFunction(LPCTSTR lpszFunc, LPSHORT lpnMinParamCnt,
		LPSHORT lpnMaxParamCnt, LPLONG lplFlags)
	{ return m_pSpread->GetCustomFunction( lpszFunc, lpnMinParamCnt,
	lpnMaxParamCnt, lplFlags); }

	int GetCustomName(LPCTSTR lpszName, LPTSTR lpszValue, int nValueLen)
	{ return m_pSpread->GetCustomName( lpszName, lpszValue, nValueLen); }

	BOOL SetCustomName(LPCTSTR lpszName, LPCTSTR lpszValue)
	{ return m_pSpread->SetCustomName( lpszName, lpszValue); }

	int GetCustomNameLen(LPCTSTR lpszName)
	{ return m_pSpread->GetCustomNameLen( lpszName); }

	int GetData(SS_COORD Col, SS_COORD Row, LPTSTR Data)
	{ return m_pSpread->GetData( Col, Row, Data); }

	BOOL SetData(SS_COORD Col, SS_COORD Row, LPCTSTR lpData)
	{ return m_pSpread->SetData ( Col, Row, lpData); }

	BOOL SetDataRange(SS_COORD Col, SS_COORD Row,
		SS_COORD Col2,SS_COORD Row2, LPCTSTR lpData)
	{ return m_pSpread->SetDataRange ( Col, Row, Col2, Row2, lpData); }

	BOOL GetDataCnt(LPSS_COORD lpColCnt, LPSS_COORD lpRowCnt)
	{ return m_pSpread->GetDataCnt( lpColCnt, lpRowCnt); }

	int GetDataLen(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->GetDataLen( Col, Row); }

	BOOL GetDefFloatFormat(LPSS_FLOATFORMAT lpFloatFormat)
	{ return m_pSpread->GetDefFloatFormat( lpFloatFormat); }

	BOOL SetDefFloatFormat(LPSS_FLOATFORMAT lpFloatFormat)
	{ return m_pSpread->SetDefFloatFormat( lpFloatFormat); }

	WORD GetEditEnterAction(void)
	{ return m_pSpread->GetEditEnterAction(); }

	WORD SetEditEnterAction(WORD wAction)
	{ return m_pSpread->SetEditEnterAction( wAction); }

	BOOL GetEditMode(void)
	{ return m_pSpread->GetEditMode(); }

	BOOL SetEditMode(BOOL fEditModeOn)
	{ 
	//	m_pSpread->SetBool(SSB_EDITMODEREPLACE,fEditModeOn);
	//	AFX_MANAGE_STATE(AfxGetStaticModuleState());
		return m_pSpread->SetEditMode( fEditModeOn); }

	BOOL GetExcelSheetList(LPCTSTR lpszFileName, GLOBALHANDLE FAR *lpghList, 
		LPSHORT lpnListCount, LPCTSTR lpszLogFileName, 
		LPSHORT lpnWorkbookHandle, BOOL fReplace)
	{ return m_pSpread->GetExcelSheetList( lpszFileName, lpghList, lpnListCount, lpszLogFileName, lpnWorkbookHandle, fReplace);}

	BOOL GetFirstValidCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
	{ return m_pSpread->GetFirstValidCell( lpCol, lpRow); }

	BOOL GetFloat(SS_COORD Col, SS_COORD Row, LPDOUBLE lpdfValue)
	{ return m_pSpread->GetFloat( Col, Row, lpdfValue); }

	BOOL SetFloat(SS_COORD Col, SS_COORD Row, double dfValue)
	{ return m_pSpread->SetFloat( Col, Row, dfValue); }

	BOOL SetFloatRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, double dfValue)
	{ return m_pSpread->SetFloatRange( Col, Row, Col2, Row2, dfValue); }

	HFONT GetFont(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->GetFont( Col, Row); }

	BOOL SetFont(SS_COORD Col, SS_COORD Row, HFONT hFont, BOOL fDeleteFont)
	{ 
		if(hFont)
			return m_pSpread->SetFont( Col, Row, hFont, fDeleteFont);
		else
			return FALSE;
	}

	BOOL SetFontRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2,
		SS_COORD Row2, HFONT hFont, BOOL fDeleteFont)
	{ 
		if(hFont)
			return m_pSpread->SetFontRange( Col, Row, Col2, Row2, hFont, fDeleteFont); 
		else
			return FALSE;
	}

	int GetFormula(SS_COORD Col, SS_COORD Row, LPTSTR lpFormula)
	{ return m_pSpread->GetFormula( Col, Row, lpFormula); }

	BOOL SetFormula(SS_COORD Col, SS_COORD Row, LPCTSTR Formula,
		BOOL BuildDependencies)
	{ return m_pSpread->SetFormula ( Col, Row, Formula, BuildDependencies); }

	BOOL SetFormulaRange(SS_COORD Col, SS_COORD Row,
		SS_COORD Col2,SS_COORD Row2,
		LPCTSTR Formula, BOOL BuildDependencies)
	{ return m_pSpread->SetFormulaRange( Col, Row, Col2, Row2, Formula, 
	BuildDependencies); }

	int GetFormulaLen(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->GetFormulaLen( Col, Row); }

	BOOL GetFreeze(LPSS_COORD lpColsFrozen, LPSS_COORD lpRowsFrozen)
	{ return m_pSpread->GetFreeze( lpColsFrozen, lpRowsFrozen); }

	BOOL SetFreeze(SS_COORD ColsFrozen, SS_COORD RowsFrozen)
	{ return m_pSpread->SetFreeze ( ColsFrozen, RowsFrozen); }

	void GetGrayAreaColor(LPCOLORREF lpBackground, LPCOLORREF lpForeground)
	{ m_pSpread->GetGrayAreaColor( lpBackground, lpForeground); }

	void SetGrayAreaColor(COLORREF Background, COLORREF Foreground)
	{ m_pSpread->SetGrayAreaColor( Background, Foreground); }

	COLORREF GetGridColor(void)
	{ return m_pSpread->GetGridColor(); }

	COLORREF SetGridColor(COLORREF Color)
	{ return m_pSpread->SetGridColor( Color); }

	WORD GetGridType(void)
	{ return m_pSpread->GetGridType(); }

	WORD SetGridType(WORD wGridType)
	{ return m_pSpread->SetGridType( wGridType); }

	BOOL GetBuffer(SS_COORD Col, SS_COORD Row, LPVOID lpBuffer, long *plSize) 
	{return m_pSpread->GetBuffer( Col, Row, lpBuffer, plSize);} 

	BOOL SetBuffer(SS_COORD Col, SS_COORD Row, LPVOID lpBuffer, long lSize) 
	{return m_pSpread->SetBuffer( Col, Row, lpBuffer, lSize);}

	BOOL GetInteger(SS_COORD Col, SS_COORD Row, LPLONG lplValue)
	{ return m_pSpread->GetInteger( Col, Row, lplValue); }

	BOOL SetInteger(SS_COORD Col, SS_COORD Row, long lValue)
	{ return m_pSpread->SetInteger( Col, Row, lValue); }

	BOOL SetIntegerRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, long lValue)
	{ return m_pSpread->SetIntegerRange( Col, Row, Col2, Row2, lValue); }

	BOOL GetIteration(LPSHORT lpnMaxIterations, LPDOUBLE lpdfMaxChange)
	{ return m_pSpread->GetIteration( lpnMaxIterations, lpdfMaxChange); }

	BOOL SetIteration(BOOL bIteration, short nMaxIterations, double dfMaxChange)
	{ return m_pSpread->SetIteration( bIteration, nMaxIterations, dfMaxChange); }

	BOOL GetLastValidCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
	{ return m_pSpread->GetLastValidCell( lpCol, lpRow); }

	BOOL GetLock(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->GetLock( Col, Row); }

	BOOL SetLock(SS_COORD Col, SS_COORD Row, BOOL Lock)
	{ return m_pSpread->SetLock ( Col, Row, Lock); }

	BOOL SetLockRange(SS_COORD Col,  SS_COORD Row,
		SS_COORD Col2, SS_COORD Row2, BOOL Lock)
	{ return m_pSpread->SetLockRange ( Col, Row, Col2, Row2, Lock); }

	void GetLockColor(LPCOLORREF lpBackground, LPCOLORREF lpForeground)
	{ m_pSpread->GetLockColor( lpBackground, lpForeground); }

	void SetLockColor(COLORREF Background, COLORREF Foreground)
	{ m_pSpread->SetLockColor( Background, Foreground); }

	SS_COORD GetMaxCols()
	{ return m_pSpread->GetMaxCols (); }

	void SetMaxCols(SS_COORD MaxCols)
	{ m_pSpread->SetMaxCols( MaxCols); }

	SS_COORD GetMaxRows()
	{ return m_pSpread->GetMaxRows(); }

	void SetMaxRows(SS_COORD MaxRows)
	{ m_pSpread->SetMaxRows( MaxRows); }

	BOOL GetMaxTextCellSize(SS_COORD Col, SS_COORD Row, LPDOUBLE lpdfWidth, LPDOUBLE lpdfHeight)
	{ return m_pSpread->GetMaxTextCellSize( Col, Row, lpdfWidth, lpdfHeight); }

	BOOL GetMaxTextColWidth(SS_COORD Col, LPDOUBLE lpdfColWidth)
	{ return m_pSpread->GetMaxTextColWidth( Col, lpdfColWidth); }

	BOOL GetMaxTextRowHeight(SS_COORD Row, LPDOUBLE lpdfRowHeight)
	{ return m_pSpread->GetMaxTextRowHeight( Row, lpdfRowHeight); }

	GLOBALHANDLE GetMultiSelBlocks(LPSHORT lpdSelBlockCnt)
	{ return m_pSpread->GetMultiSelBlocks( lpdSelBlockCnt); }

	BOOL SetMultiSelBlocks(LPSS_SELBLOCK lpSelBlockList, short dSelBlockCnt)
	{ return m_pSpread->SetMultiSelBlocks( lpSelBlockList, dSelBlockCnt); }

	long GetNextPageBreakCol(long lPrevCol)
	{ return m_pSpread->GetNextPageBreakCol( lPrevCol); }

	long GetNextPageBreakRow(long lPrevRow)
	{ return m_pSpread->GetNextPageBreakRow( lPrevRow); }

	BOOL GetOddEvenRowColor(LPCOLORREF lpclrBackOdd, LPCOLORREF lpclrForeOdd,
		LPCOLORREF lpclrBackEven, LPCOLORREF lpclrForeEven)
	{ return m_pSpread->GetOddEvenRowColor( lpclrBackOdd, lpclrForeOdd, lpclrBackEven, lpclrForeEven); }

	BOOL SetOddEvenRowColor(COLORREF clrBackOdd, COLORREF clrForeOdd,
		COLORREF clrBackEven, COLORREF clrForeEven)
	{ return m_pSpread->SetOddEvenRowColor( clrBackOdd, clrForeOdd, clrBackEven, clrForeEven); }

	WORD GetOperationMode()
	{ return m_pSpread->GetOperationMode(); }

	WORD SetOperationMode(WORD wMode)
	{ return m_pSpread->SetOperationMode( wMode); }

	HWND GetOwner()
	{ return m_pSpread->GetOwner(); }

	HWND SetOwner(HWND hWndOwner)
	{ return m_pSpread->SetOwner( hWndOwner); }

	short GetPrintAbortMsg(LPTSTR lpszText, short nLen)
	{ return m_pSpread->GetPrintAbortMsg( lpszText, nLen); }

	BOOL SetPrintAbortMsg(LPTSTR lpszText)
	{ return m_pSpread->SetPrintAbortMsg( lpszText); }

	short GetPrintJobName(LPTSTR lpszText, short nLen)
	{ return m_pSpread->GetPrintJobName( lpszText, nLen); }

	BOOL SetPrintJobName(LPTSTR lpszText)
	{ return m_pSpread->SetPrintJobName( lpszText); }

#if defined(_WIN64) || defined(_IA64)
	BOOL GetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC FAR *lpfnPrintProcAddr,LONG_PTR *lplAppData)
#else
	BOOL GetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC FAR *lpfnPrintProcAddr,long *lplAppData)
#endif
	{ return m_pSpread->GetPrintOptions( lpPrintFormat, lpfnPrintProcAddr, lplAppData); }

#if defined(_WIN64) || defined(_IA64)
	void SetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC lpfnPrintProc, LONG_PTR lAppData)
#else
	void SetPrintOptions(LPSS_PRINTFORMAT lpPrintFormat, SS_PRINTFUNC lpfnPrintProc, long lAppData)
#endif
	{ m_pSpread->SetPrintOptions( lpPrintFormat, lpfnPrintProc, lAppData); }

#if defined(_WIN64) || defined(_IA64)
	void GetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC *lpfnPrintProc, LONG_PTR *lplAppData)
#else
	void GetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC *lpfnPrintProc, long *lplAppData)
#endif
	{ m_pSpread->GetPrintOptionsEx( lpPageFormat, lpfnPrintProc, lplAppData); }

#if defined(_WIN64) || defined(_IA64)
	void SetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC lpfnPrintProc, LONG_PTR lAppData)
#else
	void SetPrintOptionsEx(LPSS_PAGEFORMAT lpPageFormat, SS_PRINTFUNC lpfnPrintProc, long lAppData)
#endif
	{ m_pSpread->SetPrintOptionsEx( lpPageFormat, lpfnPrintProc, lAppData); }

	long GetPrintPageCount()
	{ return m_pSpread->GetPrintPageCount(); }

	short GetRefStyle()
	{ return m_pSpread->GetRefStyle(); }

	short SetRefStyle(short nStyle)
	{ return m_pSpread->SetRefStyle( nStyle); }

	SS_COORD GetRowHeaderCols()
	{ return m_pSpread->GetRowHeaderCols(); }

	void SetRowHeaderCols(SS_COORD lRowHeaderCols)
	{ m_pSpread->SetRowHeaderCols( lRowHeaderCols); }

	WORD GetRowHeaderDisplay()
	{ return m_pSpread->GetRowHeaderDisplay(); }

	WORD SetRowHeaderDisplay(WORD wDisplay)
	{ return m_pSpread->SetRowHeaderDisplay( wDisplay); }

	SS_COORD GetRowHeadersAutoTextIndex()
	{ return m_pSpread->GetRowHeadersAutoTextIndex(); }

	void SetRowHeadersAutoTextIndex(SS_COORD lRowHeadersAutoTextIndex)
	{ m_pSpread->SetRowHeadersAutoTextIndex( lRowHeadersAutoTextIndex); }

	BOOL  GetRowHeight(SS_COORD Row, LPDOUBLE lpHeight)
	{ return m_pSpread->GetRowHeight( Row, lpHeight); }

	BOOL SetRowHeight(SS_COORD Row, double Height)
	{ return m_pSpread->SetRowHeight ( Row, Height); }

	BOOL GetRowHeightInPixels(SS_COORD Row, LPINT lpdHeightPixels)
	{ return m_pSpread->GetRowHeightInPixels( Row, lpdHeightPixels); }

	BOOL SetRowHeightInPixels(SS_COORD Row, int dHeight)
	{ return m_pSpread->SetRowHeightInPixels( Row, dHeight); }

	short GetRowMerge(SS_COORD Row)
	{ return m_pSpread->GetRowMerge( Row); }

	void SetRowMerge(SS_COORD Row, short nRowMerge)
	{ m_pSpread->SetRowMerge( Row, nRowMerge); }

	BOOL GetRowPageBreak(SS_COORD row)
	{ return m_pSpread->GetRowPageBreak( row); }  

	BOOL SetRowPageBreak(SS_COORD row, BOOL fpagebreak)
	{ return m_pSpread->SetRowPageBreak( row, fpagebreak); }

#if defined(_WIN64) || defined(_IA64)
	BOOL GetRowUserData(SS_COORD Row, LONG_PTR* lplUserData)
#else
	BOOL GetRowUserData(SS_COORD Row, long* lplUserData)
#endif
	{ return m_pSpread->GetRowUserData( Row, lplUserData); }

#if defined(_WIN64) || defined(_IA64)
	BOOL SetRowUserData(SS_COORD Row, LONG_PTR lUserData)
#else
	BOOL SetRowUserData(SS_COORD Row, long lUserData)
#endif
	{ return m_pSpread->SetRowUserData( Row, lUserData); }

	void GetScrollBarColor(LPCOLORREF lpVScrollBarColor, LPCOLORREF lpHScrollBarColor)
	{ m_pSpread->GetScrollBarColor( lpVScrollBarColor, lpHScrollBarColor); }

	void SetScrollBarColor(COLORREF VScrollBarColor, COLORREF HScrollBarColor)
	{ m_pSpread->SetScrollBarColor( VScrollBarColor, HScrollBarColor); }

	void GetScrollBarSize(LPSHORT lpdVScrollWidth, LPSHORT lpdHScrollHeight)
	{ m_pSpread->GetScrollBarSize( lpdVScrollWidth, lpdHScrollHeight); }

	void SetScrollBarSize(short dVScrollWidth, short dHScrollHeight)
	{ m_pSpread->SetScrollBarSize( dVScrollWidth, dHScrollHeight); }

	BOOL GetSel(LPINT SelStart, LPINT SelEnd)
	{ return m_pSpread->GetSel ( SelStart, SelEnd); }

	BOOL SetSel(int SelStart, int SelEnd)
	{ return m_pSpread->SetSel( SelStart, SelEnd); }

	WORD GetSelBlockOptions()
	{ return m_pSpread->GetSelBlockOptions(); }

	WORD SetSelBlockOptions(WORD wOption)
	{ return m_pSpread->SetSelBlockOptions( wOption); }

	void GetSelColor(LPCOLORREF lpclrBack, LPCOLORREF lpclrFore)
	{ m_pSpread->GetSelColor( lpclrBack, lpclrFore); }

	void SetSelColor(COLORREF clrBack, COLORREF clrFore)
	{ m_pSpread->SetSelColor( clrBack, clrFore); }

	BOOL GetSelectBlock(LPSS_CELLCOORD CellUL, LPSS_CELLCOORD CellLR)
	{ return m_pSpread->GetSelectBlock ( CellUL, CellLR); }

	BOOL SetSelectBlock(LPSS_CELLCOORD CellUL, LPSS_CELLCOORD CellLR)
	{ return m_pSpread->SetSelectBlock ( CellUL, CellLR); }

	GLOBALHANDLE GetSelText()
	{ return m_pSpread->GetSelText (); }

	void GetShadowColor(LPCOLORREF lpShadowColor, LPCOLORREF lpShadowText,
		LPCOLORREF lpShadowDark,  LPCOLORREF lpShadowLight)
	{ m_pSpread->GetShadowColor ( lpShadowColor, lpShadowText,
	lpShadowDark, lpShadowLight); }

	void SetShadowColor(COLORREF ShadowColor, COLORREF ShadowText,
		COLORREF ShadowDark,  COLORREF ShadowLight)
	{ m_pSpread->SetShadowColor( ShadowColor, ShadowText,
	ShadowDark, ShadowLight); }

	short GetSheet()
	{ return m_pSpread->GetSheet(); }

	void SetSheet(short nSheet)
	{ m_pSpread->SetSheet( nSheet); }

	short GetSheetCount()
	{ return m_pSpread->GetSheetCount(); }

	BOOL SetSheetCount(short nSheetCnt)
	{ return m_pSpread->SetSheetCount( nSheetCnt); }

	short GetSheetSendingMsg()
	{ return m_pSpread->GetSheetSendingMsg(); }

	short GetSheetName(short nSheet, LPTSTR lpszName, short nLen)
	{ return m_pSpread->GetSheetName( nSheet, lpszName, nLen); }

	BOOL SetSheetName(short nSheet, LPCTSTR lpszSheetName)
	{ return m_pSpread->SetSheetName( nSheet, lpszSheetName); }

	BOOL GetSheetVisible(short nSheet)
	{ return m_pSpread->GetSheetVisible( nSheet); }

	BOOL SetSheetVisible(short nSheet, BOOL fVisible)
	{ return m_pSpread->SetSheetVisible( nSheet, fVisible); }

	WORD GetShowScrollTips()
	{ return m_pSpread->GetShowScrollTips(); }

	void SetShowScrollTips(WORD wScrollTips)
	{ m_pSpread->SetShowScrollTips( wScrollTips); }

	void GetStartingNumbers(LPSS_COORD lpStartingColNumber,
		LPSS_COORD lpStartingRowNumber)
	{ m_pSpread->GetStartingNumbers( lpStartingColNumber,
	lpStartingRowNumber); }

	void SetStartingNumbers(SS_COORD StartingColNumber,
		SS_COORD StartingRowNumber)
	{ m_pSpread->SetStartingNumbers( StartingColNumber,
	StartingRowNumber); }

	HFONT GetTabStripFont()
	{ return m_pSpread->GetTabStripFont(); }

	BOOL SetTabStripFont(HFONT hFont)
	{ return m_pSpread->SetTabStripFont( hFont); }

	short GetTabStripLeftSheet()
	{ return m_pSpread->GetTabStripLeftSheet(); }

	BOOL SetTabStripLeftSheet(short nSheet)
	{ return m_pSpread->SetTabStripLeftSheet( nSheet); }

	WORD GetTabStripPolicy()
	{ return m_pSpread->GetTabStripPolicy(); }

	WORD SetTabStripPolicy(WORD wTabStripPolicy)
	{ return m_pSpread->SetTabStripPolicy( wTabStripPolicy); }

	double GetTabStripRatio()
	{ return m_pSpread->GetTabStripRatio(); }

	double SetTabStripRatio(double nTabStripRatio)
	{ return m_pSpread->SetTabStripRatio( nTabStripRatio); }

	BOOL GetTextTip(LPWORD lpwStatus, LPLONG lplDelay, LPLOGFONT lpLogFont,
		LPCOLORREF lpclrBack, LPCOLORREF lpclrFore)
	{ return m_pSpread->GetTextTip( lpwStatus, lplDelay, lpLogFont, lpclrBack, lpclrFore); }

	BOOL SetTextTip(WORD wStatus, LONG lDelay, LPLOGFONT lpLogFont,
		COLORREF clrBack, COLORREF clrFore)
	{ return m_pSpread->SetTextTip( wStatus, lDelay, lpLogFont,
	clrBack, clrFore); }

	BOOL GetTopLeftCell(LPSS_COORD lpCol, LPSS_COORD lpRow)
	{ return m_pSpread->GetTopLeftCell ( lpCol, lpRow); }

	void GetTopLeftPrev(LPSS_COORD lpCol, LPSS_COORD lpRow)
	{ m_pSpread->GetTopLeftPrev( lpCol, lpRow); }

	short GetTwoDigitYearMax()
	{ return m_pSpread->GetTwoDigitYearMax(); }

	BOOL SetTwoDigitYearMax(short nTwoDigitYearMax)
	{ return m_pSpread->SetTwoDigitYearMax( nTwoDigitYearMax); }

	WORD GetUnitType()
	{ return m_pSpread->GetUnitType(); }

	WORD SetUnitType(WORD wUnitType)
	{ return m_pSpread->SetUnitType( wUnitType); }

	short GetUserColAction()
	{return m_pSpread->GetUserColAction(); }

	void SetUserColAction(short nUserColAction)
	{ m_pSpread->SetUserColAction( nUserColAction); }

#if defined(_WIN64) || defined(_IA64)
	LONG_PTR GetUserData()
#else
	long GetUserData()
#endif
	{ return m_pSpread->GetUserData(); }

#if defined(_WIN64) || defined(_IA64)
	LONG_PTR SetUserData(LONG_PTR lUserData)
#else
	long SetUserData(long lUserData)
#endif
	{ return m_pSpread->SetUserData( lUserData); }

	WORD GetUserResize()
	{ return m_pSpread->GetUserResize(); }

	WORD SetUserResize(WORD wUserResize)
	{ return m_pSpread->SetUserResize( wUserResize); }

	short GetUserResizeCol(SS_COORD Col)
	{ return m_pSpread->GetUserResizeCol( Col); }

	short SetUserResizeCol(SS_COORD Col, short dOption)
	{ return m_pSpread->SetUserResizeCol( Col, dOption); }

	short GetUserResizeRow(SS_COORD Row)
	{ return m_pSpread->GetUserResizeRow( Row); }

	short SetUserResizeRow(SS_COORD Row, short dOption)
	{ return m_pSpread->SetUserResizeRow( Row, dOption); }

	int GetValue(SS_COORD Col, SS_COORD Row, LPTSTR lpBuffer)
	{ return m_pSpread->GetValue ( Col, Row, lpBuffer); }

	BOOL SetValue(SS_COORD Col, SS_COORD Row, LPCTSTR lpData)
	{ return m_pSpread->SetValue ( Col, Row, lpData); }

	BOOL SetValueRange(SS_COORD Col,  SS_COORD Row,
		SS_COORD Col2, SS_COORD Row2, LPCTSTR lpData)
	{ return m_pSpread->SetValueRange ( Col, Row, Col2, Row2, lpData); }

	int GetValueLen(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->GetValueLen ( Col, Row); }

	BOOL GetVisible(SS_COORD Col, SS_COORD Row, short Visible)
	{ return m_pSpread->GetVisible ( Col, Row, Visible); }

	BOOL ImportExcelBook (LPCTSTR lpszfileName, LPCTSTR lpszlogFileName)
	{ return m_pSpread->ImportExcelBook ( lpszfileName, lpszlogFileName); }

	BOOL ImportExcelSheet(short nWorkbookHandle, short nSheetIndex, LPCTSTR lpszSheetName)
	{ return m_pSpread->ImportExcelSheet( nWorkbookHandle, nSheetIndex, lpszSheetName); }

	BOOL InsCol(SS_COORD Col)
	{ return m_pSpread->InsCol ( Col); }

	BOOL InsColRange(SS_COORD Col, SS_COORD Col2)
	{ return m_pSpread->InsColRange( Col, Col2); }

	BOOL InsertSheets(short nSheetIndex, short nSheetCnt)
	{ return m_pSpread->InsertSheets ( nSheetIndex, nSheetCnt); }

	BOOL InsRow(SS_COORD Row)
	{ return m_pSpread->InsRow ( Row); }

	BOOL InsRowRange(SS_COORD Row, SS_COORD Row2)
	{ return m_pSpread->InsRowRange( Row, Row2); }

	BOOL IsCellInSelection(SS_COORD Col, SS_COORD Row)
	{ return m_pSpread->IsCellInSelection( Col, Row); }

	BOOL IsColHidden(SS_COORD Col)
	{ return m_pSpread->IsColHidden( Col); }

	int IsExcelFile(LPCTSTR lpszFileName)
	{ return m_pSpread->IsExcelFile (lpszFileName);}

	BOOL IsFetchCellNote()
	{ return m_pSpread->IsFetchCellNote(); }

	BOOL IsRowHidden(SS_COORD Row)
	{ return m_pSpread->IsRowHidden( Row); }

	BOOL LoadFromBuffer(GLOBALHANDLE hBuffer, long lBufferLen)
	{ return m_pSpread->LoadFromBuffer ( hBuffer, lBufferLen); }

	BOOL LoadFromFile(LPCTSTR lpszFileName)
	{ return m_pSpread->LoadFromFile ( lpszFileName); }

	HANDLE LoadPicture(LPCTSTR lpszFileName, short nPictType)
	{ return m_pSpread->LoadPicture(lpszFileName, nPictType); }

	HANDLE LoadPictureBuffer(HGLOBAL hGlobal, long lSize, short nPictType)
	{ return m_pSpread->LoadPictureBuffer(hGlobal, lSize, nPictType); }

	HANDLE LoadResPicture(HINSTANCE hInstance, LPCTSTR lpszResName, LPCTSTR lpszResType, short nPictType)
	{ return m_pSpread->LoadResPicture(hInstance, lpszResName, lpszResType, nPictType); }

	BOOL LoadTabFile(LPCTSTR lpszFileName)
	{ return m_pSpread->LoadTabFile( lpszFileName); }

	BOOL LoadTextFile (LPCTSTR pcszFileName, LPCTSTR pcszCellDelim, LPCTSTR pcszColDelim, LPCTSTR pcszRowDelim, long lFlags, LPCTSTR pcszLogFile)
	{ return m_pSpread->LoadTextFile ( pcszFileName, pcszCellDelim, pcszColDelim, pcszRowDelim, lFlags, pcszLogFile); }

	void LogUnitsToColWidth(short dUnits, LPDOUBLE lpWidth)
	{ m_pSpread->LogUnitsToColWidth( dUnits, lpWidth); }

	void LogUnitsToRowHeight(SS_COORD Row, short dUnits, LPDOUBLE lpHeight)
	{ m_pSpread->LogUnitsToRowHeight( Row, dUnits, lpHeight); }

	BOOL MoveRange(SS_COORD Col,SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
		SS_COORD ColDest, SS_COORD RowDest)
	{ return m_pSpread->MoveRange( Col, Row, Col2, Row2, ColDest, RowDest); }

	BOOL MoveSheets(short nSheetSrc, short nSheetCnt, short nSheetDest)
	{ return m_pSpread->MoveSheets( nSheetSrc, nSheetCnt, nSheetDest); }

#if defined(_WIN64) || defined(_IA64)
	BOOL Print(LPCTSTR lpszPrintJobName, LPSS_PRINTFORMAT lpPrintFormat,
		SS_PRINTFUNC lpfnPrintProc, LONG_PTR lAppData,
		ABORTPROC lpfnAbortProc)
#else
	BOOL Print(LPCTSTR lpszPrintJobName, LPSS_PRINTFORMAT lpPrintFormat,
		SS_PRINTFUNC lpfnPrintProc, LONG lAppData,
		ABORTPROC lpfnAbortProc)
#endif
	{ return m_pSpread->Print( lpszPrintJobName, lpPrintFormat,
	lpfnPrintProc, lAppData, lpfnAbortProc); }

	BOOL OwnerPrintDraw(HDC hDC, long left, long top, long right, long bottom, short page)
	{ return m_pSpread->OwnerPrintDraw( hDC, left, top, right, bottom, page);}

	BOOL OwnerPrintPageCount(HDC hDC, long left, long top, long right, long bottom, LPSHORT pageCount)
	{ return m_pSpread->OwnerPrintPageCount( hDC, left, top, right, bottom, pageCount);}
	int QueryCustomName(LPCTSTR lpszPrevName, LPTSTR lpszName, int nNameLen)
	{ return m_pSpread->QueryCustomName( lpszPrevName, lpszName, nNameLen); }

	int QueryCustomNameLen(LPCTSTR lpszPrevName)
	{ return m_pSpread->QueryCustomNameLen( lpszPrevName); }

	BOOL ReCalc()
	{ return m_pSpread->ReCalc (); }

	BOOL ReCalcCell(SS_COORD lCol, SS_COORD lRow)
	{ return m_pSpread->ReCalcCell( lCol, lRow); }

	void RemoveCellSpan(SS_COORD lCol, SS_COORD lRow)
	{ m_pSpread->RemoveCellSpan( lCol, lRow); }

	BOOL RemoveCustomCellType(LPCTSTR lpszName)
	{ return m_pSpread->RemoveCustomCellType( lpszName); }

	BOOL RemoveCustomFunction(LPCTSTR lpszFuncName)
	{ return m_pSpread->RemoveCustomFunction( lpszFuncName); }

	BOOL ReplaceSel(LPCTSTR lpText)
	{ return m_pSpread->ReplaceSel ( lpText); }

	void Reset()
	{ m_pSpread->Reset(); }

	void ResetSheet(short nSheetIndex)
	{ m_pSpread->ResetSheet( nSheetIndex); }

	int RowHeightToLogUnits(SS_COORD Row, double dfRowHeight)
	{ return m_pSpread->RowHeightToLogUnits ( Row, dfRowHeight); }

	BOOL SavePicture(HANDLE hPict, HPALETTE hPal, LPCTSTR lpszFileName, short nSaveType)
	{ return m_pSpread->SavePicture(hPict, hPal, lpszFileName, nSaveType); }

	BOOL SavePictureBuffer(HANDLE hPict, HPALETTE hPal, short nSaveType, HGLOBAL *phGlobal, long *plSize)
	{ return m_pSpread->SavePictureBuffer(hPict, hPal, nSaveType, phGlobal, plSize); }

	BOOL SaveTabFile(LPCTSTR lpszFileName)
	{ return m_pSpread->SaveTabFile( lpszFileName); }

	BOOL SaveTabFileU(LPCTSTR lpszFileName)
	{ return m_pSpread->SaveTabFileU( lpszFileName); }

	GLOBALHANDLE SaveToBuffer(LPLONG lpBufferLen, BOOL fDataOnly)
	{ return m_pSpread->SaveToBuffer ( lpBufferLen, fDataOnly); }

	BOOL SaveToFile(LPCTSTR lpszFileName, BOOL fDataOnly)
	{ return m_pSpread->SaveToFile ( lpszFileName, fDataOnly); }

	SS_COORD SearchCol(SS_COORD lCol, SS_COORD lRowStart, SS_COORD lRowEnd, LPCTSTR lpszText, WORD wSearchFlags)
	{ return m_pSpread->SearchCol( lCol, lRowStart, lRowEnd, lpszText, wSearchFlags); }

	SS_COORD SearchRow(SS_COORD lRow, SS_COORD lColStart, SS_COORD lColEnd, LPCTSTR lpszText, WORD wSearchFlags)
	{ return m_pSpread->SearchRow( lRow, lColStart, lColEnd, lpszText, wSearchFlags); }

#if defined(_WIN64) || defined(_IA64)
	long SelModeSendMessage(UINT uMsg, long lParam1, LONG_PTR lParam2, long lParam3)
#else
	long SelModeSendMessage(UINT uMsg, long lParam1, long lParam2, long lParam3)
#endif
	{ return m_pSpread->SelModeSendMessage( uMsg, lParam1, lParam2, lParam3); }

	LPSS_CELLTYPE SetTypeButton(LPSS_CELLTYPE lpCellType, long Style,
		LPCTSTR Text, LPCTSTR Picture, 
		short nPictureType, LPCTSTR PictureDown,
		short nPictureDownType, short ButtonType,
		short nShadowSize, LPSUPERBTNCOLOR lpColor)
	{ return m_pSpread->SetTypeButton( lpCellType, Style, 
	Text, Picture, nPictureType,
	PictureDown, nPictureDownType,
	ButtonType, nShadowSize, lpColor); }

	LPSS_CELLTYPE SetTypeCheckBox(LPSS_CELLTYPE lpCellType,
		long lStyle, LPCTSTR lpText,
		LPCTSTR lpPictUp, WORD wPictUpType,
		LPCTSTR lpPictDown, WORD wPictDownType,
		LPCTSTR lpPictFocusUp, WORD wPictFocusUpType,
		LPCTSTR lpPictFocusDown, WORD wPictFocusDownType,
		LPCTSTR lpPictGray, WORD wPictGrayType,
		LPCTSTR lpPictFocusGray, WORD wPictFocusGrayType)
	{ return m_pSpread->SetTypeCheckBox( lpCellType, lStyle, lpText, lpPictUp, wPictUpType,
	lpPictDown, wPictDownType, lpPictFocusUp, wPictFocusUpType,
	lpPictFocusDown, wPictFocusDownType, lpPictGray, wPictGrayType,
	lpPictFocusGray, wPictFocusGrayType); }

	LPSS_CELLTYPE SetTypeComboBox(LPSS_CELLTYPE lpCellType, long Style, LPCTSTR lpItems)
	{ return m_pSpread->SetTypeComboBox( lpCellType, Style, lpItems); }


	LPSS_CELLTYPE SetTypeComboBoxEx(LPSS_CELLTYPE lpCellType, long Style, 
		LPCTSTR lpItems, short dMaxRows, short dMaxEditLen,
		short dComboWidth, HWND hWndDropDown, BYTE bAutoSearch)
	{ return m_pSpread->SetTypeComboBoxEx( lpCellType, Style, 
	lpItems, dMaxRows, dMaxEditLen,
	dComboWidth, hWndDropDown, bAutoSearch); }

	LPSS_CELLTYPE SetTypeCurrency(LPSS_CELLTYPE lpCellType, long style, BYTE Right, double Min, double Max, BYTE fShowSeparator)
	{ return m_pSpread->SetTypeCurrency( lpCellType, style, Right, Min, Max, fShowSeparator); }

	LPSS_CELLTYPE SetTypeCurrencyEx(LPSS_CELLTYPE lpCellType, long Style, BYTE Right, double Min, double Max, BYTE fShowSeparator, BYTE fShowCurrencySymbol, LPTSTR lpszDecimal, LPTSTR lpszSeparator, LPTSTR lpszCurrency, BYTE bLeadingZero, BYTE bNegCurrencyStyle, BYTE bPosCurrencyStyle, BYTE fSpin, BYTE fSpinWrap, double SpinInc)
	{ return m_pSpread->SetTypeCurrencyEx( lpCellType, Style, Right, Min, Max, fShowSeparator, fShowCurrencySymbol, lpszDecimal, lpszSeparator, lpszCurrency, bLeadingZero, bNegCurrencyStyle, bPosCurrencyStyle, fSpin, fSpinWrap, SpinInc); } 

	LPSS_CELLTYPE  SetTypeCustom(LPSS_CELLTYPE lpCellType, LPCTSTR lpszName, long Style, LPSS_CT_VALUE lpItemData)
	{ return m_pSpread->SetTypeCustom( lpCellType, lpszName, Style, lpItemData);}

	LPSS_CELLTYPE SetTypeDate(LPSS_CELLTYPE lpCellType, long Style,
		LPSS_DATEFORMAT lpFormat, LPSS_DATE lpMin, 
		LPSS_DATE lpMax)
	{ return m_pSpread->SetTypeDate( lpCellType, Style, lpFormat, lpMin, lpMax); }

	LPSS_CELLTYPE SetTypeEdit(LPSS_CELLTYPE lpCellType, long Style, short Len,
		short ChrSet, short ChrCase)
	{ return m_pSpread->SetTypeEdit( lpCellType, Style, Len, ChrSet, ChrCase); }

	LPSS_CELLTYPE SetTypeFloat(LPSS_CELLTYPE lpCellType, long Style,
		short Left, short Right, double Min, 
		double Max)
	{ return m_pSpread->SetTypeFloat( lpCellType, Style, Left, Right, Min, Max); }

	LPSS_CELLTYPE SetTypeFloatExt(LPSS_CELLTYPE lpCellType, long Style,
		short Left, short Right, double Min, 
		double Max, LPSS_FLOATFORMAT lpFormat)
	{ return m_pSpread->SetTypeFloatExt( lpCellType, Style, Left, Right, Min, Max, lpFormat); }

	LPSS_CELLTYPE SetTypeInteger(LPSS_CELLTYPE lpCellType, long Min, long Max)
	{ return m_pSpread->SetTypeInteger( lpCellType, Min, Max); }

	LPSS_CELLTYPE SetTypeIntegerExt(LPSS_CELLTYPE lpCellType, long lStyle, long lMin,
		long lMax, BOOL fSpinWrap, long lSpinInc)
	{ return m_pSpread->SetTypeIntegerExt( lpCellType, lStyle, lMin, lMax, fSpinWrap, lSpinInc); }

	LPSS_CELLTYPE SetTypeNumber(LPSS_CELLTYPE lpCellType, long lStyle, BYTE Right, double Min, double Max, BYTE fShowSeparator)
	{ return m_pSpread->SetTypeNumber( lpCellType, lStyle, Right, Min, Max, fShowSeparator); }

	LPSS_CELLTYPE SetTypeNumberEx(LPSS_CELLTYPE lpCellType, long lStyle, BYTE Right, double Min, double Max, BYTE fShowSeparator, LPTSTR lpszDecimal, LPTSTR lpszSeparator, BYTE bLeadingZero, BYTE bNegStyle, BYTE fSpin, BYTE fSpinWrap, double SpinInc)
	{ return m_pSpread->SetTypeNumberEx( lpCellType, lStyle, Right, Min, Max, fShowSeparator, lpszDecimal, lpszSeparator, bLeadingZero, bNegStyle, fSpin, fSpinWrap, SpinInc); } 							 

	LPSS_CELLTYPE SetTypeOwnerDraw(LPSS_CELLTYPE lpCellType, long lStyle)
	{ return m_pSpread->SetTypeOwnerDraw( lpCellType, lStyle); }

	LPSS_CELLTYPE SetTypePercent(LPSS_CELLTYPE lpCellType, long Style, BYTE Right, double Min, double Max)
	{ return m_pSpread->SetTypePercent( lpCellType, Style, Right, Min, Max); }

	LPSS_CELLTYPE SetTypePercentEx(LPSS_CELLTYPE lpCellType, long Style, BYTE Right, double Min, double Max, LPTSTR lpszDecimal,  BYTE bLeadingZero, BYTE bNegStyle, BYTE fSpin, BYTE fSpinWrap, double SpinInc)
	{ return m_pSpread->SetTypePercentEx( lpCellType, Style, Right, Min, Max, lpszDecimal, bLeadingZero, bNegStyle, fSpin, fSpinWrap, SpinInc); }

	LPSS_CELLTYPE SetTypePic(LPSS_CELLTYPE lpCellType, long Style, LPCTSTR Mask)
	{ return m_pSpread->SetTypePic( lpCellType, Style, Mask); }

	LPSS_CELLTYPE SetTypePicture(LPSS_CELLTYPE lpCellType, long Style,
		LPCTSTR PictName)
	{ return m_pSpread->SetTypePicture( lpCellType, Style, PictName); }

	LPSS_CELLTYPE SetTypePictureHandle(LPSS_CELLTYPE lpCellType, long Style, HANDLE hPict,
		HPALETTE hPal, BOOL fDeleteHandle)
	{ return m_pSpread->SetTypePictureHandle( lpCellType, Style, hPict, hPal, fDeleteHandle); }

	LPSS_CELLTYPE SetTypeScientific(LPSS_CELLTYPE lpCellType, long lStyle, BYTE Right, double Min, double Max, LPTSTR lpszDecimal)
	{ return m_pSpread->SetTypeScientific( lpCellType, lStyle, Right, Min, Max, lpszDecimal); }

	LPSS_CELLTYPE SetTypeStaticText(LPSS_CELLTYPE lpCellType, long TextStyle)
	{ return m_pSpread->SetTypeStaticText( lpCellType, TextStyle); }

	LPSS_CELLTYPE SetTypeTime(LPSS_CELLTYPE lpCellType, long Style,
		LPSS_TIMEFORMAT lpFormat, LPSS_TIME lpMin,
		LPSS_TIME lpMax)
	{ return m_pSpread->SetTypeTime ( lpCellType, Style, lpFormat, lpMin, lpMax); }

	BOOL ShowActiveCell(short Position)
	{ return m_pSpread->ShowActiveCell ( Position); }

	BOOL ShowCell(SS_COORD Col, SS_COORD Row, short Position)
	{ return m_pSpread->ShowCell ( Col, Row, Position); }

	BOOL ShowCol(SS_COORD Col, BOOL fShow)
	{ return m_pSpread->ShowCol( Col, fShow); }

	BOOL ShowRow(SS_COORD Row, BOOL fShow)
	{ return m_pSpread->ShowRow( Row, fShow); }

	BOOL Sort(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2, LPSS_SORT lpSort)
	{ return m_pSpread->Sort( Col, Row, Col2, Row2, lpSort); }

	BOOL SortEx(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
		WORD wSortBy, LPSS_SORTKEY lpSortKeys, short nSortKeyCnt)
	{ return m_pSpread->SortEx( Col, Row, Col2, Row2, wSortBy, lpSortKeys, nSortKeyCnt); }

	BOOL SwapRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,
		SS_COORD ColDest, SS_COORD RowDest)
	{ return m_pSpread->SwapRange( Col, Row, Col2, Row2, ColDest, RowDest); }

	BOOL ValidateFormula (LPCTSTR lpszFormula)
	{ return m_pSpread->ValidateFormula ( lpszFormula); }

	void VGetBufferSize(LPSS_COORD lpdBufferSize, LPSS_COORD lpdOverlap)
	{ m_pSpread->VGetBufferSize( lpdBufferSize, lpdOverlap); }

	void VSetBufferSize(SS_COORD dBufferSize, SS_COORD dOverlap)
	{ m_pSpread->VSetBufferSize( dBufferSize, dOverlap); }

	SS_COORD VGetMax()
	{ return m_pSpread->VGetMax(); }

	SS_COORD VSetMax(SS_COORD Max)
	{ return m_pSpread->VSetMax( Max); }

	SS_COORD VGetPhysBufferSize(void)
	{ return m_pSpread->VGetPhysBufferSize(); }

	SS_COORD VGetPhysBufferTop(void)
	{ return m_pSpread->VGetPhysBufferTop(); }

	LONG VGetStyle(void)
	{ return m_pSpread->VGetStyle(); }

	LONG VSetStyle(LONG lStyle)
	{ return m_pSpread->VSetStyle( lStyle); }

	BOOL VRefreshBuffer()
	{ return m_pSpread->VRefreshBuffer(); }

	BOOL VScrollGetSpecial(LPWORD lpwOptions)
	{ return m_pSpread->VScrollGetSpecial( lpwOptions); }

	BOOL VScrollSetSpecial(BOOL fUseSpecialVScroll, WORD wOptions)
	{ return m_pSpread->VScrollSetSpecial( fUseSpecialVScroll, wOptions); }


	WORD GetTabStripButtonPolicy(void)
	{return m_pSpread->GetTabStripButtonPolicy();}

	WORD SetTabStripButtonPolicy(WORD wTabStripButtonPolicy)
	{return m_pSpread->SetTabStripButtonPolicy( wTabStripButtonPolicy);}

	BOOL IsVerticalScrollBarVisible(void)
	{return m_pSpread->IsVerticalScrollBarVisible();}

	BOOL IsHorizontalScrollBarVisible(void)
	{return m_pSpread->IsHorizontalScrollBarVisible();}

	WORD SetAppearanceStyle(WORD nStyle)
	{return m_pSpread->SetAppearanceStyle( nStyle);}

	WORD GetAppearanceStyle(void)
	{return m_pSpread->GetAppearanceStyle();}

	WORD SetScrollBarStyle(WORD nStyle)
	{return m_pSpread->SetScrollBarStyle( nStyle);}

	WORD GetScrollBarStyle(void)
	{return m_pSpread->GetScrollBarStyle();}

	WORD SetUseVisualStyles(WORD nStyle)
	{return m_pSpread->SetUseVisualStyles( nStyle);}

	WORD GetUseVisualStyles(void)
	{return m_pSpread->GetUseVisualStyles();}

	BOOL SetEnhanceStaticCells(BOOL nStyle)
	{return m_pSpread->SetEnhanceStaticCells( nStyle);}

	BOOL GetEnhanceStaticCells(void)
	{return m_pSpread->GetEnhanceStaticCells();}

	WORD SetTabEnhancedShape(WORD nStyle)
	{return m_pSpread->SetTabEnhancedShape( nStyle);}

	WORD GetTabEnhancedShape(void)
	{return m_pSpread->GetTabEnhancedShape();}

	WORD GetCellNoteIndicatorShape(void)
	{return m_pSpread->GetCellNoteIndicatorShape();}

	WORD SetCellNoteIndicatorShape(WORD nStyle)
	{return m_pSpread->SetCellNoteIndicatorShape( nStyle);}

	COLORREF GetCellNoteIndicatorColor(void)
	{return m_pSpread->GetCellNoteIndicatorColor();}

	COLORREF SetCellNoteIndicatorColor(COLORREF Color)
	{return m_pSpread->SetCellNoteIndicatorColor( Color);}

	BOOL SetEnhancedCornerColors(COLORREF clrBackColor, COLORREF clrHoverColor, COLORREF clrTriangle, COLORREF clrTriangleHover, COLORREF clrTriangleBorderColor, COLORREF clrTriangleHoverBorderColor)
	{return m_pSpread->SetEnhancedCornerColors( clrBackColor, clrHoverColor, clrTriangle, clrTriangleHover, clrTriangleBorderColor, clrTriangleHoverBorderColor);}

	BOOL SetEnhancedColumnHeaderColors(COLORREF clrSelectedHoverUpperColor, COLORREF clrSelectedHoverLowerColor, COLORREF clrSelectedUpperColor,
		COLORREF clrSelectedLowerColor, COLORREF clrHoverUpperColor, COLORREF clrHoverLowerColor, COLORREF clrUpperColor,
		COLORREF clrLowerColor, COLORREF clrSelectedBorderColor, COLORREF clrBorderColor)
	{return m_pSpread->SetEnhancedColumnHeaderColors( clrSelectedHoverUpperColor, clrSelectedHoverLowerColor, clrSelectedUpperColor,
	clrSelectedLowerColor, clrHoverUpperColor, clrHoverLowerColor, clrUpperColor,
	clrLowerColor, clrSelectedBorderColor, clrBorderColor);}


	BOOL SetEnhancedRowHeaderColors(COLORREF clrSelectedHoverUpperColor, COLORREF clrSelectedHoverLowerColor, COLORREF clrSelectedUpperColor,
		COLORREF clrSelectedLowerColor, COLORREF clrHoverUpperColor, COLORREF clrHoverLowerColor, COLORREF clrUpperColor,
		COLORREF clrLowerColor, COLORREF clrSelectedBorderColor, COLORREF clrBorderColor)
	{return m_pSpread->SetEnhancedRowHeaderColors( clrSelectedHoverUpperColor, clrSelectedHoverLowerColor, clrSelectedUpperColor,
	clrSelectedLowerColor, clrHoverUpperColor, clrHoverLowerColor, clrUpperColor,
	clrLowerColor, clrSelectedBorderColor, clrBorderColor);}

	BOOL SetEnhancedSheetTabColors(COLORREF clrNormalUpperStartColor, COLORREF clrNormalUpperEndColor,
		COLORREF clrNormalLowerStartColor, COLORREF clrNormalLowerEndColor, COLORREF clrHoverUpperStartColor,
		COLORREF clrHoverUpperEndColor, COLORREF clrHoverLowerStartColor, COLORREF clrHoverLowerEndColor,
		COLORREF clrTabOuterBorderColor, COLORREF clrTabInnerBorderColor, COLORREF clrButtonUpperNormalStartColor,
		COLORREF clrButtonUpperNormalEndColor, COLORREF clrButtonLowerNormalStartColor, COLORREF clrButtonLowerNormalEndColor,
		COLORREF clrButtonUpperHoverStartColor, COLORREF clrButtonUpperHoverEndColor, COLORREF clrButtonLowerHoverStartColor,
		COLORREF clrButtonLowerHoverEndColor,COLORREF clrButtonUpperPushedStartColor, COLORREF clrButtonUpperPushedEndColor,
		COLORREF clrButtonLowerPushedStartColor, COLORREF clrButtonLowerPushedEndColor, COLORREF clrArrowColor,
		COLORREF clrBackColorStartColor, COLORREF clrBackColorEndColor, COLORREF clrSplitBoxStartColor,
		COLORREF clrSplitBoxEndColor, COLORREF clrForeColor, COLORREF clrActiveForeColor)
	{return m_pSpread->SetEnhancedSheetTabColors( clrNormalUpperStartColor, clrNormalUpperEndColor,
	clrNormalLowerStartColor, clrNormalLowerEndColor, clrHoverUpperStartColor,
	clrHoverUpperEndColor, clrHoverLowerStartColor, clrHoverLowerEndColor,
	clrTabOuterBorderColor, clrTabInnerBorderColor, clrButtonUpperNormalStartColor,
	clrButtonUpperNormalEndColor, clrButtonLowerNormalStartColor, clrButtonLowerNormalEndColor,
	clrButtonUpperHoverStartColor, clrButtonUpperHoverEndColor, clrButtonLowerHoverStartColor,
	clrButtonLowerHoverEndColor, clrButtonUpperPushedStartColor, clrButtonUpperPushedEndColor,
	clrButtonLowerPushedStartColor, clrButtonLowerPushedEndColor, clrArrowColor,
	clrBackColorStartColor, clrBackColorEndColor, clrSplitBoxStartColor,
	clrSplitBoxEndColor, clrForeColor, clrActiveForeColor);}




	BOOL SetEnhancedScrollBarColors( COLORREF clrTrackColor, COLORREF clrArrowColor, COLORREF clrUpperNormalStartColor, COLORREF clrUpperNormalEndColor,
		COLORREF clrLowerNormalStartColor, COLORREF clrLowerNormalEndColor, COLORREF clrUpperPushedStartColor, COLORREF clrUpperPushedEndColor,
		COLORREF clrLowerPushedStartColor, COLORREF clrLowerPushedEndColor, COLORREF clrUpperHoverStartColor, COLORREF clrUpperHoverEndColor,
		COLORREF clrLowerHoverStartColor, COLORREF clrLowerHoverEndColor, COLORREF clrHoverButtonBorderColor, COLORREF clrButtonBorderColor)
	{return m_pSpread->SetEnhancedScrollBarColors( clrTrackColor, clrArrowColor, clrUpperNormalStartColor, clrUpperNormalEndColor,
	clrLowerNormalStartColor, clrLowerNormalEndColor, clrUpperPushedStartColor, clrUpperPushedEndColor,
	clrLowerPushedStartColor, clrLowerPushedEndColor, clrUpperHoverStartColor, clrUpperHoverEndColor,
	clrLowerHoverStartColor, clrLowerHoverEndColor, clrHoverButtonBorderColor, clrButtonBorderColor);}

	WORD SetHighlightHeaders(WORD wOption)
	{return m_pSpread->SetHighlightHeaders( wOption);}

	WORD GetHighlightHeaders(void)
	{return m_pSpread->GetHighlightHeaders();}

	WORD SetHighlightStyle(WORD wStyle)
	{return m_pSpread->SetHighlightStyle( wStyle);}

	WORD GetHighlightStyle(void)
	{return m_pSpread->GetHighlightStyle();}

	BOOL SetHighlightAlphaBlend(COLORREF clrAlphaBlend, short nAlphaBlend)
	{return m_pSpread->SetHighlightAlphaBlend( clrAlphaBlend, nAlphaBlend);}

	BOOL GetHighlightAlphaBlend(LPCOLORREF lpclrAlphaBlend, LPSHORT lpnAlphaBlend)
	{return m_pSpread->GetHighlightAlphaBlend( lpclrAlphaBlend, lpnAlphaBlend);}

	BOOL ExportToHTMLEx (LPCTSTR lpszFileName, 
		BOOL bAppendFlag, LPCTSTR lpszLogFile, LPCTSTR lpszTitle, LPCTSTR lpszCssFile, LPCTSTR lpszTableClass, LPCTSTR lpszTableRowClass, LPCTSTR lpszTableDefinitionClass, LPCTSTR lpszTableHeaderClass)
	{return m_pSpread->ExportToHTMLEx ( lpszFileName, bAppendFlag, lpszLogFile, lpszTitle, lpszCssFile, lpszTableClass, lpszTableRowClass, lpszTableDefinitionClass, lpszTableHeaderClass);}

	BOOL ExportRangeToHTMLEx (SS_COORD col, SS_COORD row, SS_COORD col2, SS_COORD row2, 
		LPCTSTR lpszFileName, BOOL bAppendFlag, 
		LPCTSTR lpszLogFile, LPCTSTR lpszTitle,
		LPCTSTR lpszCssFile, LPCTSTR lpszTableClass, LPCTSTR lpszTableRowClass, LPCTSTR lpszTableDefinitionClass, LPCTSTR lpszTableHeaderClass)
	{return m_pSpread->ExportRangeToHTMLEx ( col, row, col2, row2, lpszFileName, bAppendFlag, lpszLogFile, lpszTitle, lpszCssFile,
	lpszTableClass, lpszTableRowClass, lpszTableDefinitionClass, lpszTableHeaderClass);}
	BOOL SaveExcel2007File(LPCTSTR lpszFileName, LPCTSTR lpszPassword, short nFlags, LPCTSTR lpszLogFileName)
	{
		return m_pSpread->SaveExcel2007File( lpszFileName, lpszPassword, nFlags, lpszLogFileName);
	}
	BOOL OpenExcel2007File(LPCTSTR lpszFileName, LPCTSTR lpszPassword, short nSheet, short nExcelSheet, LPCTSTR lpszLogFileName)
	{
		return m_pSpread->OpenExcel2007File( lpszFileName, lpszPassword, nSheet, nExcelSheet, lpszLogFileName);
	}
	short IsExcel2007File(LPCTSTR lpszFileName)
	{
		return m_pSpread->IsExcel2007File( lpszFileName);
	}
	BOOL LoadBlockFromBuffer(SS_COORD lCol, SS_COORD lRow,
		SS_COORD lCol2, SS_COORD lRow2,
		LPBYTE lpBuff, long lBuffLen)
	{
		return m_pSpread->LoadBlockFromBuffer( lCol, lRow, lCol2, lRow2, lpBuff, lBuffLen);
	}
	HGLOBAL SaveBlockToBuffer(SS_COORD lCol, SS_COORD lRow,
		SS_COORD lCol2, SS_COORD lRow2,
		LPLONG lplBuffLen, BOOL  bDataOnly)
	{
		return m_pSpread->SaveBlockToBuffer( lCol, lRow, lCol2, lRow2, lplBuffLen, bDataOnly);
	}

	HGLOBAL SaveBlockToBufferExt(SS_COORD lCol, SS_COORD lRow,
		SS_COORD lCol2, SS_COORD lRow2,
		LPLONG lplBuffLen, BOOL  bDataOnly, BOOL fSaveColInfo, BOOL fSaveRowInfo)
	{
		return m_pSpread->SaveBlockToBufferExt( lCol, lRow, lCol2, lRow2, lplBuffLen, bDataOnly, fSaveColInfo, fSaveRowInfo);
	}


	void GetCellPos (long Col, long Row, long* x, long* y, long* Width, long* Height)
  {
	  RECT rt;
	  m_pSpread->GetCellRect(Col,Row,&rt);

	  *x=rt.left;
	  *y=rt.top;
	  *Width=rt.right-rt.left;
	  *Height=rt.bottom-rt.top;
  }

  void ClearSelection()
  {
	  m_pSpread->DeSelectBlock();
  }

  void SetRetainSelBlock(BOOL bValue)
  {
	  m_pSpread->SetBool(SSB_RETAINSELBLOCK,bValue);
  }

  void SetCellNote(LPTSTR str)
  {
	  m_pSpread->SetCellNote(m_nCols,m_nRows,str);
  }

  void SetValueEx(LPCTSTR str)
  {
	  m_pSpread->SetValue(m_nCols,m_nRows,str);
  }

  CString GetText()
  {
	  TCHAR strVal[1024];
	  SS_CELLTYPE stype;
	
	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);

	  if(stype.Type == SS_TYPE_COMBOBOX)
		  m_pSpread->GetData(m_nCols,m_nRows,strVal);
	  else
		  m_pSpread->GetValue(m_nCols,m_nRows,strVal);

	  return CString(strVal);
  }

  void SetTypeButtonText(LPCTSTR str)
  {
	  m_pSpread->SetButtonText(m_nCols,m_nRows,str);
  }

  BOOL ClearRange(SS_COORD Col, SS_COORD Row, SS_COORD Col2, SS_COORD Row2,BOOL bDataOnly)
  { return m_pSpread->ClearRange( Col, Row, Col2, Row2); }

  void SetLock(BOOL bLock)
  {
	  m_pSpread->SetLock(m_nCols,m_nRows,bLock);
  }

  void SetTypeComboBoxEditable(BOOL bValue)
  {
	  //dll에는 없다.
	  return;
  }

  void SetTypeComboBoxMaxDrop(int nVal)
  {
	  SS_CELLTYPE stype;
	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);

	  stype.Spec.ComboBox.dMaxRows = nVal;
	  m_pSpread->SetCellType(m_nCols,m_nRows,&stype);
  }

  void SetForeColor(COLORREF fore)
  {
	  //return;
	  COLORREF oldfore;
	  COLORREF oldbk;
	  m_pSpread->GetColor(m_nCols,m_nRows,&oldbk,&oldfore);
	  m_pSpread->SetColor(m_nCols,m_nRows,oldbk,fore);
  }

  void SetSelBackColor(COLORREF bk)
  {
	  COLORREF fore;
	  COLORREF oldbk;
	  m_pSpread->GetSelColor(&oldbk,&fore);
	  m_pSpread->SetSelColor(bk,fore);
  }


  void SetBackColor(COLORREF ref)
  {
	//  AFX_MANAGE_STATE(AfxGetStaticModuleState());
	  //return;
	  COLORREF fore;
	  COLORREF oldbk;
	  m_pSpread->GetColor(m_nCols,m_nRows,&oldbk,&fore);
	  m_pSpread->SetColor(m_nCols,m_nRows,ref,fore);
  }
  COLORREF GetBackColor()
  {
	  COLORREF fore;
	  COLORREF oldbk;
	  m_pSpread->GetColor(m_nCols,m_nRows,&oldbk,&fore);

	  return oldbk;
  }

  void SetTypeComboBoxList(LPCTSTR str)
  {
	  SS_CELLTYPE stype;
	  m_pSpread->SetTypeComboBox(&stype,SSS_ALIGN_CENTER  | SS_CB_DROPDOWN,str);
	  m_pSpread->SetCellType(m_nCols, m_nRows, &stype);
  }

  void SetTypeVAlign(long Val)
  {
	  SS_CELLTYPE stype;

	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);
	  stype.Style |=Val;
	  //m_pSpread->SetTypeStaticText()
	  //  m_pSpread->SetTypeEdit(&stype,Val,256,SS_CHRSET_CHR,SS_CASE_NOCASE);
	  m_pSpread->SetCellType(m_nCols, m_nRows, &stype);
  }

  void SetTypeHAlign(long Val)
  {
	  SS_CELLTYPE stype;

	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);
	  stype.Style |=Val;
	  //m_pSpread->SetTypeStaticText()
	  //  m_pSpread->SetTypeEdit(&stype,Val,256,SS_CHRSET_CHR,SS_CASE_NOCASE);
	  m_pSpread->SetCellType(m_nCols, m_nRows, &stype);
  }

  void SetCellType(long nType,long Val=0,LPCTSTR str=NULL)
  {
	  SS_CELLTYPE CellType;
	  /*
	  #define TYPE_COMBO		8
	  #define TYPE_CHECK		10
	  #define TYPE_BUTTON		7
	  #define TYPE_STATIC		5
	  */

	  switch(nType)
	  {
	  case 1:
		  m_pSpread->SetTypeEdit(&CellType,Val,128,SS_CHRSET_CHR,SS_CASE_NOCASE);
		  break;

	  case 8:
		  m_pSpread->SetTypeComboBox(&CellType, 0, str);
		  break;

	  case 5:
		  m_pSpread->SetTypeStaticText(&CellType, Val | SS_TEXT_WORDWRAP | SS_TEXT_VCENTER);
		  break;

	  case 7:
		  m_pSpread->SetTypeButton (
			  &CellType,
			  SBS_TEXTLEFT_PICTRIGHT, str,
			  NULL, SUPERBTN_PICT_NONE,
			  NULL, SUPERBTN_PICT_NONE,
			  SUPERBTN_TYPE_NORMAL,
			  2, NULL
			  );
		  break;

	  case 10:
		  m_pSpread->SetTypeCheckBox (
			  &CellType,
			  0, _T("Click Me"),
			  NULL, BT_NONE,
			  NULL, BT_NONE,
			  NULL, BT_NONE,
			  NULL, BT_NONE,
			  NULL, BT_NONE,
			  NULL, BT_NONE
			  );			
		  break;

	  }
	  m_pSpread->SetCellType(m_nCols, m_nRows, &CellType);
  }

  BOOL SetCellBorder(SS_COORD Col, SS_COORD Row, SS_COORD Col2,
	  SS_COORD Row2, WORD wBorderType, COLORREF Color,WORD wBorderStyle)
  { return m_pSpread->SetBorderRange( Col, Row, Col2, Row2, wBorderType, wBorderStyle, Color); }

  void SetFontBold(BOOL bValue)
  {
	  return;
  }

  void SetFontSize(int nSize)
  {

	  return;
  }

  void SetTypeStaticAlign(long Val)
  {
	  // m_pSpread->GetCellType(m
	  // m_pSpread->SetTypeStaticText()
	  SS_CELLTYPE stype;
	  m_pSpread->SetTypeStaticText(&stype,Val);
	  m_pSpread->SetCellType(m_nCols, m_nRows, &stype);
  }

  void SetTypeAlign(long Val)
  {
	  SS_CELLTYPE stype;

	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);
	  stype.Style |=Val;
	  //m_pSpread->SetTypeStaticText()
	//  m_pSpread->SetTypeEdit(&stype,Val,256,SS_CHRSET_CHR,SS_CASE_NOCASE);
	  m_pSpread->SetCellType(m_nCols, m_nRows, &stype);
  }

  void SetTypeEditMultiLine(BOOL bValue)
  {
	  SS_CELLTYPE stype;

	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);

	  if(bValue)
	  {
		stype.Style |=ES_MULTILINE;
	  }
	  else
	  {
		  long bit = 0;

		  // or로 ㅣ비트를 준비한다
		  bit = bit | ES_MULTILINE;
		  //반전 시킨다.
		  bit = ~bit;
		  stype.Style = stype.Style & bit;
	  }
	  m_pSpread->SetCellType(m_nCols, m_nRows, &stype);
  }
  
  long GetCellType()
  {
	  SS_CELLTYPE stype;

	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);

	  return stype.Type;
  }

  void SetAutoCalc(BOOL bValue)
  {
	  SetBool(SSB_AUTOCALC,bValue);
  }
  void SetAllowMultiBlocks(BOOL bValue)
  {
	  SetBool(SSB_ALLOWMULTIBLOCKS,bValue);
  }

  void SetEditModeReplace(BOOL bValue)
  {
	  SetBool(SSB_EDITMODEREPLACE,bValue);
  }

  void SetCol(SS_COORD Col)
  {
	  m_nCols = Col;
	  //TRACE("Col%d\n",Col);
  }

  void SetRow(SS_COORD Row)
  {
	 // if(Row == 17)
	//	  int kk=0;
	  m_nRows = Row;
	//  TRACE("Row%d\n",m_nRows);
  }

// Operations
  //160221 lgh add custom
  void SetReDraw(BOOL bValue)
  {
	  m_pSpread->SetBool(SSB_REDRAW,bValue);
  }

  void SetText(LPCTSTR strData)
  {
	  SS_CELLTYPE stype;

	  m_pSpread->GetCellType(m_nCols,m_nRows,&stype);

	  if(stype.Type == SS_TYPE_COMBOBOX)
		  m_pSpread->SetData(m_nCols,m_nRows,strData);
	  else
		  m_pSpread->SetData(m_nCols,m_nRows,strData);

	  //return;
	//  return;
	//  m_pSpread->SetValue(m_nCols,m_nRows,strData);
  }

  double GetColWidth(SS_COORD nCol)
  {
	//  return 10;
	  double dColGetWidth2 = 0;
	  m_pSpread->GetColWidth(nCol,&dColGetWidth2);
	  return dColGetWidth2;
  }

  void SetFocus()
  {
	  if(m_pParentReal)
		  m_pSpread=(TSpread*)m_pParentReal->GetDlgItem(m_nRCID);
	  m_pSpread->SetFocus();
  }

  void AutoAttach()
  {
	  if(m_pParentReal)
		  m_pSpread=(TSpread*)m_pParentReal->GetDlgItem(m_nRCID);
  }

  void SetSelection(SS_COORD lCol, SS_COORD lRow, SS_COORD lCol2, SS_COORD lRow2)
  {
	  SS_CELLCOORD start, end;

	  start.Col=lCol;
	  start.Row=lRow;

	  end.Col=lCol2;
	  end.Row=lRow2;

	  m_pSpread->SetSelectBlock(&start,&end);
  }

  double GetMaxTextCellHeight()
  {
	  double dWidth=0,dHeight=0;
	  m_pSpread->GetMaxTextCellSize(m_nCols,m_nRows,&dWidth,&dHeight);
	  return dHeight;
  }

  double GetMaxTextCellWidth()
  {
	  double dWidth=0,dHeight=0;
	  m_pSpread->GetMaxTextCellSize(m_nCols,m_nRows,&dWidth,&dHeight);
	  return dWidth;
  }

  void SetRowHeadersShow(BOOL bVal)
  {
	  SetBool(SSB_SHOWROWHEADERS,bVal);
  }

  void SetColHeadersShow(BOOL bVal)
  {
	  SetBool(SSB_SHOWCOLHEADERS,bVal);
  }

  void TwipsToRowHeight(SS_COORD Row, SS_COORD Twips, double* RowHeight)
  {
	  m_pSpread->LogUnitsToRowHeight(Row,Twips,RowHeight);
  }

  void TwipsToColWidth(SS_COORD Twips, double* ColWidth)
  {
	  m_pSpread->LogUnitsToColWidth(Twips,ColWidth);
  }

  void SetRowMerge(short value)
  {
	  m_pSpread->SetRowMerge(m_nRows,value);
  }

  void SetColMerge(short Value)
  {
	  m_pSpread->SetColMerge(m_nCols,Value);
  }

  void SetGrayAreaBackColor(COLORREF value)
  {
	  COLORREF oldbk,oldfore;
	  m_pSpread->GetGrayAreaColor(&oldbk,&oldfore);
	  m_pSpread->SetGrayAreaColor(oldbk,oldfore);
  }

  CString GetValue()
  {
	  TCHAR strbuff[1024];
	  m_pSpread->GetValue(m_nCols,m_nRows,strbuff);

	  return CString(strbuff);
  }

  void SetValue(CString strBuff)
  {
	  m_pSpread->SetValue(m_nCols,m_nRows,strBuff);
  }

  void SetTopRow(SS_COORD nRow)
  {
	  AutoAttach();

	  m_pSpread->ShowCell(SS_HEADER,nRow,SS_SHOW_TOPLEFT);
  }

  HWND GetHwnd()
  {
	  return m_pSpread->m_hWnd;
  }  

 };



//----------------------------------------------------------
class TPreview : public CWnd
{
  DECLARE_DYNAMIC(TPreview)

// Constructors
public:  
  TPreview() {}
  BOOL Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID);

  BOOL GetAllowUserZoom(void)
     {BOOL value; SpvGetAllowUserZoom(m_hWnd, &value); return value;}

  int SetAllowUserZoom(BOOL value)
     { return SpvSetAllowUserZoom ( m_hWnd, value); }

  COLORREF GetGrayAreaColor(void)
     {COLORREF value; SpvGetGrayAreaColor(m_hWnd, &value); return value;}

  int SetGrayAreaColor(COLORREF value)
	{ return SpvSetGrayAreaColor ( m_hWnd, value); }

  long GetGrayAreaMarginH(void)
     {long value; SpvGetGrayAreaMarginH(m_hWnd, &value); return value;}

  int SetGrayAreaMarginH(long value)
     { return SpvSetGrayAreaMarginH ( m_hWnd, value); }

  short GetGrayAreaMarginType(void)
     {short value; SpvGetGrayAreaMarginType(m_hWnd, &value); return value;}

  int SetGrayAreaMarginType(short value)
     { return SpvSetGrayAreaMarginType ( m_hWnd, value); }

  long GetGrayAreaMarginV(void)
     {long value; SpvGetGrayAreaMarginV(m_hWnd, &value); return value;}

  int SetGrayAreaMarginV(long value)
     { return SpvSetGrayAreaMarginV ( m_hWnd, value); }

  HWND GethWndSpread(void)
     {HWND value; SpvGethWndSpread(m_hWnd, &value); return value;}

  int SethWndSpread(HWND value)
     { return SpvSethWndSpread ( m_hWnd, value); }

  COLORREF GetPageBorderColor(void)
     {COLORREF value; SpvGetPageBorderColor(m_hWnd, &value); return value;}

  int SetPageBorderColor(COLORREF value)
     { return SpvSetPageBorderColor ( m_hWnd, value); }

  short GetPageBorderWidth(void)
     {short value; SpvGetPageBorderWidth(m_hWnd, &value); return value;}

  int SetPageBorderWidth(short value)
     { return SpvSetPageBorderWidth ( m_hWnd, value); }

  long GetPageCurrent(void)
     {long value; SpvGetPageCurrent(m_hWnd, &value); return value;}

  int SetPageCurrent(long value)
     { return SpvSetPageCurrent ( m_hWnd, value); }

  long GetPageGutterH(void)
     {long value; SpvGetPageGutterH(m_hWnd, &value); return value;}

  int SetPageGutterH(long value)
     { return SpvSetPageGutterH ( m_hWnd, value); }

  long GetPageGutterV(void)
     {long value; SpvGetPageGutterV(m_hWnd, &value); return value;}

  int SetPageGutterV(long value)
     { return SpvSetPageGutterV ( m_hWnd, value); }

  short GetPageMultiCntH(void)
     {short value; SpvGetPageMultiCntH(m_hWnd, &value); return value;}

  int SetPageMultiCntH(short value)
     { return SpvSetPageMultiCntH ( m_hWnd, value); }

  short GetPageMultiCntV(void)
     {short value; SpvGetPageMultiCntV(m_hWnd, &value); return value;}

  int SetPageMultiCntV(short value)
     { return SpvSetPageMultiCntV ( m_hWnd, value); }

  short GetPagePercentageActual(void)
     {short value; SpvGetPagePercentageActual(m_hWnd, &value); return value;}

  COLORREF GetPageShadowColor(void)
     {COLORREF value; SpvGetPageShadowColor(m_hWnd, &value); return value;}

  int SetPageShadowColor(COLORREF value)
     { return SpvSetPageShadowColor ( m_hWnd, value); }

  short GetPageShadowWidth(void)
     {short value; SpvGetPageShadowWidth(m_hWnd, &value); return value;}

  int SetPageShadowWidth(short value)
     { return SpvSetPageShadowWidth ( m_hWnd, value); }

  short GetPagesPerScreen(void)
     {short value; SpvGetPagesPerScreen(m_hWnd, &value); return value;}

  short GetPageViewPercentage(void)
     {short value; SpvGetPageViewPercentage(m_hWnd, &value); return value;}

  int SetPageViewPercentage(short value)
     { return SpvSetPageViewPercentage ( m_hWnd, value); }

  short GetPageViewType(void)
     {short value; SpvGetPageViewType(m_hWnd, &value); return value;}

  int SetPageViewType(short value)
     { return SpvSetPageViewType ( m_hWnd, value); }

  short GetScrollBarH(void)
     {short value; SpvGetScrollBarH(m_hWnd, &value); return value;}

  int SetScrollBarH(short value)
     { return SpvSetScrollBarH ( m_hWnd, value); }

  short GetScrollBarV(void)
     {short value; SpvGetScrollBarV(m_hWnd, &value); return value;}

  int SetScrollBarV(short value)
     { return SpvSetScrollBarV ( m_hWnd, value); }

  long GetScrollIncH(void)
     {long value; SpvGetScrollIncH(m_hWnd, &value); return value;}

  int SetScrollIncH(long value)
     { return SpvSetScrollIncH ( m_hWnd, value); }

  long GetScrollIncV(void)
     {long value; SpvGetScrollIncV(m_hWnd, &value); return value;}

  int SetScrollIncV(long value)
     { return SpvSetScrollIncV ( m_hWnd, value); }

  short GetZoomState(void)
     {short value; SpvGetZoomState(m_hWnd, &value); return value;}

  int SetZoomState(short value)
     { return SpvSetZoomState ( m_hWnd, value); }

// Implementation
  virtual ~TPreview() {};
  virtual WNDPROC* GetSuperWndProcAddr();

}; // end class TPreview

//----------------------------------------------------------

#endif

