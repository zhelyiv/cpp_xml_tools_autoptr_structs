<!-- .hmmessage P { margin:0px; padding:0px } body.hmmessage { font-size: 12pt; font-family:Calibri } --> 
#include "stdafx.h"
#include "SmartTools.h" 
#include <stdio.h>
#include <string.h>
#include <fstream> // std::ifstream
#include <iterator> // std::istream_iterator
 
using namespace std; 
#include <algorithm> // std::find
 namespace smart_tools
 { 
	extern CString ToISODateTime(const TIMESTAMP_STRUCT dt)
 {
	//YYYY-MM-DDThh:mm:ss
	CString strISODateTime = GetStringSmall("%d-%02d-%02dT%02d:%02d:%02d", dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second);
	return strISODateTime;
 } 
	extern CString GetCurDateTime(INT_PTR nFormat)
 {
 CString strTime;
 COleDateTime dt = COleDateTime::GetCurrentTime();

	if (nFormat == ISODate)//ISO 8601
	{
	//YYYY-MM-DD //2014-05-23
	strTime = GetStringSmall("%d-%02d-%02d", dt.GetYear(), dt.GetMonth(), dt.GetDay());
 } 
	else if (nFormat == ISODateTime)//ISO 8601
	{
	//YYYY-MM-DDThh:mm:ss//2014-05-23T02:40:53 
	strTime = GetStringSmall("%d-%02d-%02dT%02d:%02d:%02d", dt.GetYear(), dt.GetMonth(), dt.GetDay(), dt.GetHour(), dt.GetMinute(), dt.GetSecond()); 
 } 
	else if (nFormat == DB_DateTime)
 {
	//SQL DateTime
	strTime = dt.Format("%Y-%m-%d %H:%M:%S");
 } 
	else if (nFormat == DT_Regular) 
 { 
	//custom
	strTime = dt.Format("%m/%d/%Y %H:%M:%S"); 
	}//else
 	return strTime;
 } 
 CString GetStringSmall(const char *pszFormat, ...)
 {
 CString strResult;
	char szFormated[1024] = { 0 };
 va_list arglist;
 va_start(arglist, pszFormat);
	_vsnprintf_s(szFormated, sizeof(szFormated), _TRUNCATE, pszFormat, arglist);
 strResult = szFormated;
	return strResult;
 }
	CString ToCString(char* pszData)
 {
	return GetStringSmall("%s", pszData);
 }

	CString ToCString(char cData)
 {
	return GetStringSmall("%c", cData);
 }
 CString ToCString(INT_PTR nData)
 {
	return GetStringSmall("%d", nData);
 }
	CString ToCString(long lData)
 {
	return GetStringSmall("%ld", lData);
 }
 CString ToCString(ULONG lData)
 {
	return GetStringSmall("%ld", lData);
 }
	CString ToCString(double dData)
 {
	return GetStringSmall("%.2f", dData);
 }
 CString EscapeSpecials(CString strString)
 {
	// strString.Replace('/', '//');
	// 
	// // TODO други специални символи
 	return strString;
 }
	std::vector<BYTE> ReadFileBinary(const char* pszFileName)
 { 
	//file opening
	std::streampos sSize;
 std::ifstream file(pszFileName, std::ios::binary);

	//get size
	file.seekg(0, std::ios::end);
 sSize = file.tellg();
 file.seekg(0, std::ios::beg);

	//read the data
	std::vector<BYTE> vFileData((size_t)sSize);
	file.read((char*) &vFileData[0], sSize);

	return vFileData;
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// CXmlParser
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	
	BOOL CXmlParser::InitXml(const CString strXml, INT_PTR nEncoding, MSXML2::IXMLDOMDocumentPtr& pXMLDom)
 {
	try
	{
 _bstr_t bstrxml;
 _bstr_t bstrBuff;
	if (nEncoding == ENCODING_UTF_16LE)
 {
 LPCWSTR pwstr = CT2W(strXml);
	int nLen = wcslen(pwstr);
	LPSTR lpsz = new char[nLen];
	auto_ptr<char> pAutoMem(lpsz);
 INT_PTR nResult = WideCharToMultiByte(0, 0, pwstr, nLen, lpsz, nLen, 0, 0);
	if (nResult <= 0)
 {
	return FALSE;
	}//if
 bstrBuff = (WCHAR*)lpsz;
	}//if
	else
	{
 bstrxml = strXml;
	}//else 
 	HRESULT hr = pXMLDom.CreateInstance(__uuidof(DOMDocument30));
	if (FAILED(hr))
 {
	return FALSE;
	}//if
 	pXMLDom->async = false;
 VARIANT_BOOL vbResult = pXMLDom->loadXML(bstrxml);
	if (vbResult == VARIANT_FALSE)
 {
	return FALSE;
	}//if
 CString strLoadedXml = pXMLDom->xml;
	if (pXMLDom->parseError->errorCode != 0)
 {
	if (pXMLDom)
 pXMLDom.Release();
	return FALSE;
	}//if
 	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
	return FALSE;
	}//catch
 	return TRUE;
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	
 CString CXmlParser::GetRootNodeName( CString strXml, INT_PTR nEncoding)
 {
 CString strRootNode;
	try
	{ 
	if (!InitXml(strXml, nEncoding, m_pXMLDom))
	return _T("");

 CString strResult = m_pXMLDom->documentElement->nodeName;
 strRootNode = strResult;
	}//try
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__); 
	return _T("");
	}//catch
 	return strRootNode; 
 } 
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 HRESULT CXmlParser::LoadAndParse(CString strPathToXmlFile, INT_PTR nEncoding)
 { 
	if (strPathToXmlFile.IsEmpty())
	return S_FALSE;
	if (!PathFileExists(strPathToXmlFile))
 {
	LAPPERROR("[%s] The file does not exists : %s", __FUNCTION__, strPathToXmlFile);
	return S_FALSE;
	}//if
 m_strFileName = PathFindFileName(strPathToXmlFile);

 CString strBuff; 
 CStdioFile oStdioFile(strPathToXmlFile, CStdioFile::modeRead);

	while (oStdioFile.ReadString(strBuff))
 m_strXmlOriginal += strBuff;
	return Parse(m_strXmlOriginal, nEncoding);
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 HRESULT CXmlParser::Parse( CString strXml, INT_PTR nEncoding)
 {
 m_nEncoding = nEncoding;
 CString strRootNodeName = GetRootNodeName(strXml);
 INT_PTR nRecursionDeepLevel = 0;
	return Parse(strXml, strRootNodeName, _T(""), nRecursionDeepLevel); 
 }

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 HRESULT CXmlParser::Parse( CString strXml, CString strNodeName, CString strParentNodeName, INT_PTR& nRecursionDeepLevel, INT_PTR nEncoding)
 { 
	m_vNodeNames.push_back(strNodeName);// вкарваме във вектора името на текущия node по който се двиижим в рамките на фунцкията
 CString strResult;
 HRESULT hr = S_OK;
	try
	{
 nRecursionDeepLevel++;
 INT_PTR nRecursionLevelBackup = nRecursionDeepLevel;

 MSXML2::IXMLDOMDocumentPtr pXMLDom;
 MSXML2::IXMLDOMNodePtr pChild;
 MSXML2::IXMLDOMNodePtr pRootNode;
 BOOL bResult = InitXml(strXml, nEncoding, pXMLDom);
	if (!bResult)
	return S_FALSE;

 _bstr_t bstrNode = strNodeName;
 pRootNode = pXMLDom->selectSingleNode(bstrNode);
	if (!pRootNode)
 {
	if (pXMLDom)
 pXMLDom.Release();
	//TODO
	hr = S_FALSE;
	goto LBL_RETURN;
	}//if
 CString strRootSubXml = pRootNode->xml;
	//LPCSTR lpstrKeyRef = _T("");
 	CString strMapKey = GetStringSmall("%s_%s", strParentNodeName, strNodeName); 
	CString strMapKeyFragm = strMapKey;// GetStringSmall("%s", strNodeName);
 	if (m_bCashFragments)
 DoStackMappingFragments(strMapKeyFragm, strRootSubXml);
	if (nRecursionLevelBackup == 1)
	TRACE("Trace CONTROL of recursion level 1");
	for (pChild = pRootNode->firstChild; pChild; pChild = pChild->nextSibling)
 {
	if (nRecursionLevelBackup == 2)
	TRACE("Trace CONTROL of recursion level 2");
 CString strNode = pChild->nodeName;
 CString strChildSubXml = pChild->xml;
 HRESULT hrParse = Parse(strChildSubXml, strNode, strNodeName, nRecursionDeepLevel);

	if (hrParse == S_FALSE && strNode == "#text")
 {
	DoStackMappingValues(strMapKey ,strChildSubXml);// опашка стойности от дървото(т.е чистите данни)
	
MSXML2::IXMLDOMNamedNodeMapPtr pXMLAttributeMap = pRootNode->Getattributes();
	if (pXMLAttributeMap)
 {
	long lCount = pXMLAttributeMap->Getlength();
	for (long i = 0; i < lCount; i++)
 {
 CString strAttributeName = pXMLAttributeMap->Getitem(i)->GetnodeName();
 CString strAttributeValue = pXMLAttributeMap->Getitem(i)->Gettext();
	CString strAttrMapKey = GetStringSmall("%s_%s", strNodeName, strAttributeName );
 DoStackMappingAttributeValues(strAttrMapKey, strAttributeValue);
	}//for
 
	}//if
	
	}//if
 	}//for
 	}//try
	catch (_com_error& ex)
 {
_bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("Exception in [%s] recursion level %d ParentNode %s ChildNode %s " __FUNCTION__, nRecursionDeepLevel, strParentNodeName, strNodeName);
	}//catch
 LBL_RETURN:
	return hr;
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 	void CXmlParser::DoStackMappingFragments(CString strKey, CString strValue)
 {
 CStringStack* pStack = (CStringStack*)m_mapStackElements[strKey];
	// Ако не съществува мапинг на фрагмента го създаваме, ако вече има, само добавяме новата стойност накрая в опашакат
	if (!pStack)
 {
	pStack = new CStringStack(FALSE); // Искаме да си пазим елементите в този стек
	pStack->PushBack(strValue);
 m_mapStackElements.SetAt(strKey, pStack);
	}//if
	else
	{
 pStack->PushBack(strValue);
	}//else
	}
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 	void CXmlParser::DoStackMappingValues(CString strKey, CString strValue)
 {
 CStringStack* pStack = (CStringStack*)m_mapStackNodeValues[strKey];
	// Ако не съществува мапинг на фрагмента го създаваме, ако вече има, само добавяме новата стойност накрая в опашакат
	if (!pStack)
 {
	pStack = new CStringStack();
 pStack->PushBack(strValue);
 m_mapStackNodeValues.SetAt(strKey, pStack);
	}//if
	else
	{
 pStack->PushBack(strValue);
	}//else
	}
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 	void CXmlParser::DoStackMappingAttributeValues(CString strKey, CString strValue)
 {
 CStringStack* pStack = (CStringStack*)m_mapStackNodeAttributeValues[strKey];
	// Ако не съществува мапинг на фрагмента го създаваме, ако вече има, само добавяме новата стойност накрая в опашакат
	if (!pStack)
 {
	pStack = new CStringStack();
 pStack->PushBack(strValue);
 m_mapStackNodeAttributeValues.SetAt(strKey, pStack);
	}//if
	else
	{
 pStack->PushBack(strValue);
	}//else
	}
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	
 CString CXmlParser::GetXmlFragment( CString strParentNodeName, CString strChildNodeName)
 { 
	CString strMapKey = GetStringSmall("%s_%s", strParentNodeName, strChildNodeName);
 CStringStack* pStack = (CStringStack*)m_mapStackElements[strMapKey];
	if (!pStack)
	return _T("");
 CString strFragment = pStack->PopFront();
	return strFragment;
 }

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 INT_PTR CXmlParser::GetCountOfFragment( CString strChildNodeName, CString strParentNodeName)
 {
	CString strMapKey = GetStringSmall("%s_%s", strParentNodeName, strChildNodeName);
 CStringStack* pStack = (CStringStack*)m_mapStackElements[strMapKey];
	if (!pStack)
	return 0;
	return pStack->GetCount();
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 INT_PTR CXmlParser::GetCountOfValues( CString strChildNodeName, CString strParentNodeName)
 {
	CString strMapKey = GetStringSmall("%s_%s", strParentNodeName, strChildNodeName);
CStringStack* pStack = (CStringStack*)m_mapStackNodeValues[strMapKey];
	if (!pStack)
	return 0;
	return pStack->GetCount();
 }

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 BOOL CXmlParser::DoesNodeExists( CString strNodeName)
 {
	if (std::find(m_vNodeNames.begin(), m_vNodeNames.end(), strNodeName) != m_vNodeNames.end())
	return TRUE;
	return FALSE;
 }

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 CString CXmlParser::GetValue( CString strChildNodeName, CString strParentNodeName, CString strParentParentNodeName)
 {
	//return _GetValue(strChildNodeName, strParentNodeName);// алтернатива на долното (по бързо е, но резултатът не е сигурен)
	
CString strXmlFragment = GetXmlFragment(strParentParentNodeName, strParentNodeName);
	if (strXmlFragment.IsEmpty())
	return _T("");
 CString strResult = CXmlParser::SttGetValue(strXmlFragment, strChildNodeName, strParentNodeName, m_nEncoding);
	return strResult;
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 CString CXmlParser::GetAttributeValue( CString strNodeName, CString strParentNodeName, CString strParentParentNodeName, CString strAttributeName)
 {
	//return _GetAttributeValue(strNodeName, strAttributeName);// алтернатива на долното (по бързо е, но резултатът не е сигурен)
 CString strXmlFragment = GetXmlFragment(strParentNodeName, strNodeName);

	if (strXmlFragment.IsEmpty())
	return _T("");
 CString strResult = CXmlParser::SttGetAttributeValue(strXmlFragment, strNodeName, strAttributeName);

	return strResult;
 }

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 CString CXmlParser::SttGetValue( CString strXmlFragment, CString strSubNodeName, CString strParentNode, INT_PTR nEncoding)
 { 
 CXmlParser oParser;
 oParser.SetCashFragments(FALSE);
 HRESULT hr = oParser.Parse(strXmlFragment, nEncoding);
	if (hr != S_OK)
	return _T("");
 CString strResult = oParser._GetValue(strSubNodeName, strParentNode);
	return strResult;
 }

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 CString CXmlParser::SttGetAttributeValue( CString strXmlFragment, CString strSubNodeName, CString strArtibuteName, INT_PTR nEncoding)
 {
 CXmlParser oParser;
 oParser.SetCashFragments(FALSE);
 HRESULT hr = oParser.Parse(strXmlFragment, nEncoding);
	if (hr != S_OK)
	return _T("");
 CString strResult = oParser._GetAttributeValue(strSubNodeName, strArtibuteName);
	return strResult;
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 CString CXmlParser::_GetValue( CString strChildNodeName, CString strParentNodeName)
 {
	CString strMapKey = GetStringSmall("%s_%s", strParentNodeName, strChildNodeName);
 CStringStack* pStack = (CStringStack*)m_mapStackNodeValues[strMapKey];
	if (!pStack)
	return _T("");
 CString strValue = pStack->PopFront();
	return strValue;
 }
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 CString CXmlParser::_GetAttributeValue( CString strNodeName, CString strAttributeName)
 {
	CString strMapKey = GetStringSmall("%s_%s", strNodeName, strAttributeName);
 CStringStack* pStack = (CStringStack*)m_mapStackNodeAttributeValues[strMapKey];
	if (!pStack)
	return _T("");
 CString strValue = pStack->PopFront();
	return strValue;
 }
　
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// CXmlBuilder2
 CXmlBuilder2::CXmlBuilder2()
	:m_strRootNodeName("empty")
	, m_strEncoding(_T("UTF-8"))
 , m_bHasBeenFinilized(FALSE)
 , m_lCountOfChilds(0)
 {
 HRESULT hr = S_OK;
	hr = m_pXMLDoc.CreateInstance(__uuidof(DOMDocument30));
	hr = m_pLoadXML.CreateInstance(__uuidof(DOMDocument30));
	hr = m_pXMLFormattedDoc.CreateInstance(__uuidof(DOMDocument30));
 }
 CXmlBuilder2::CXmlBuilder2(CString strRootNodeName, CString strEncoding)
 :m_strRootNodeName(strRootNodeName)
 , m_strEncoding(strEncoding)
 , m_bHasBeenFinilized(FALSE)
 , m_lCountOfChilds(0)
 {
 HRESULT hr = S_OK;
	hr = m_pXMLDoc.CreateInstance(__uuidof(DOMDocument30));
	hr = m_pLoadXML.CreateInstance(__uuidof(DOMDocument30));
	hr = m_pXMLFormattedDoc.CreateInstance(__uuidof(DOMDocument30));
 }

 CXmlBuilder2::~CXmlBuilder2()
 {
	// if (m_pXMLDoc)
	// m_pXMLDoc->Release(); 
	// if (m_pXMLRootElem)
	// m_pXMLRootElem->Release(); 
	// if (m_pXMLProcessingNode)
	// m_pXMLProcessingNode->Release();
	// if (m_pLoadXML)
	// m_pLoadXML->Release();
	// if (m_pXMLFormattedDoc)
	// m_pXMLFormattedDoc->Release();
	}
 BOOL CXmlBuilder2::InitXml( CString strXml, INT_PTR nEncoding, MSXML2::IXMLDOMDocument2Ptr& pXMLDom)
 {
	try
	{
 _bstr_t bstrxml;
 _bstr_t bstrBuff;
	if (nEncoding == ENCODING_UTF_16LE)
 {
 LPCWSTR pwstr = CT2W(strXml);
	int nLen = wcslen(pwstr);
	LPSTR lpsz = new char[nLen];
	auto_ptr<char> pAutoMem(lpsz);
 INT_PTR nResult = WideCharToMultiByte(0, 0, pwstr, nLen, lpsz, nLen, 0, 0);
	if (nResult <= 0)
 {
	return FALSE;
	}//if
 bstrBuff = (WCHAR*)lpsz;
	}//if
	else
	{
 bstrxml = strXml;
	}//else 
 	HRESULT hr = pXMLDom.CreateInstance(__uuidof(DOMDocument30));
	if (FAILED(hr))
 {
	return FALSE;
	}//if
 	pXMLDom->async = false;
 VARIANT_BOOL vbResult = pXMLDom->loadXML(bstrxml);
	if (vbResult == VARIANT_FALSE)
 {
	return FALSE;
	}//if
 CString strLoadedXml = pXMLDom->xml;
	if (pXMLDom->parseError->errorCode != 0)
 {
	if (pXMLDom)
 pXMLDom.Release();
	return FALSE;
	}//if
 	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
	return FALSE;
	}//catch
 	return TRUE;
 }
 BOOL CXmlBuilder2::LoadXml(CString strXml, INT_PTR nEncoding)
 {
	if( !InitXml(strXml, nEncoding, m_pXMLDoc ) )
	return FALSE;

 m_bHasBeenFinilized = TRUE;
	return TRUE;
 }
 HRESULT CXmlBuilder2::NewDocument( CString strRootAttr1, CString strRootAttrValue1, 
 CString strRootAttr2, CString strRootAttrValue2, 
 CString strRootAttr3, CString strRootAttrValue3 )
 {
 HRESULT hr = S_OK;
	try
	{
 hr = S_OK;
	if(!m_pXMLDoc)
	hr = m_pXMLDoc.CreateInstance(__uuidof(DOMDocument30));

	if (FAILED(hr))
	return S_FALSE;
	_bstr_t bstrEmptyXML = GetStringSmall("<%s></%s>", m_strRootNodeName, m_strRootNodeName);
	if (m_pXMLDoc->loadXML(bstrEmptyXML) == VARIANT_FALSE)
	return S_FALSE;
 m_pXMLRootElem = m_pXMLDoc->GetdocumentElement();
_bstr_t bstrAttrName;
 _bstr_t bstrAttrValue;
	if (!strRootAttr1.IsEmpty() && !strRootAttrValue1.IsEmpty())
 {
 bstrAttrName = strRootAttr1;
 bstrAttrValue = strRootAttrValue1;
 m_pXMLRootElem->setAttribute(bstrAttrName, _variant_t(bstrAttrValue));
	}//if
	if (!strRootAttr2.IsEmpty() && !strRootAttrValue2.IsEmpty())
 {
 bstrAttrName = strRootAttr2;
 bstrAttrValue = strRootAttrValue2;
 m_pXMLRootElem->setAttribute(bstrAttrName, _variant_t(bstrAttrValue));
	}//if
	if (!strRootAttr3.IsEmpty() && !strRootAttrValue3.IsEmpty())
 {
 bstrAttrName = strRootAttr3;
 bstrAttrValue = strRootAttrValue3;
 m_pXMLRootElem->setAttribute(bstrAttrName, _variant_t(bstrAttrValue));
	}//if
 CString strHeadVerEnc;
	strHeadVerEnc = GetStringSmall(" version='1.0' encoding='%s'", m_strEncoding);
 _bstr_t bstrHeadVerEnc = strHeadVerEnc;
	m_pXMLProcessingNode = m_pXMLDoc->createProcessingInstruction("xml", bstrHeadVerEnc);
 _variant_t vtObject;
 vtObject.vt = VT_DISPATCH;
 vtObject.pdispVal = m_pXMLRootElem;
 vtObject.pdispVal->AddRef();
 m_pXMLDoc->insertBefore(m_pXMLProcessingNode, vtObject);
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
 hr = S_FALSE;
	}//cacth
 	return hr;
 }
 HRESULT CXmlBuilder2::Finalize()
 { 
 HRESULT hr = S_OK;
	try
	{
	if (m_bHasBeenFinilized)
	return S_OK;
 hr = S_OK;
	if(!m_pLoadXML)
	hr = m_pLoadXML.CreateInstance(__uuidof(DOMDocument30));

	if (FAILED(hr))
	return S_FALSE;

	// прилагаме стил за да получим красиво форматиран XML
	if (m_pLoadXML->load(variant_t(_T("StyleSheet.xsl"))) == VARIANT_FALSE)
	return S_FALSE;

	//Създаваме финалния документ, който е форматиран
	hr = S_OK;
	if(!m_pXMLFormattedDoc)
	hr = m_pXMLFormattedDoc.CreateInstance(__uuidof(DOMDocument30));
 CComPtr<IDispatch> pDispatch;
	hr = m_pXMLFormattedDoc->QueryInterface(IID_IDispatch, (void**)&pDispatch);
	if (FAILED(hr)) 
	return S_FALSE;

 _variant_t vtOutObject;
 vtOutObject.vt = VT_DISPATCH;
 vtOutObject.pdispVal = pDispatch;
 vtOutObject.pdispVal->AddRef();
	// Прилагаме форматиране
	hr = m_pXMLDoc->transformNodeToObject(m_pLoadXML, vtOutObject);
	//By default - UTF-16. Променяме енкодинга на m_strEncoding 
	MSXML2::IXMLDOMNodePtr pXMLFirstChild = m_pXMLFormattedDoc->GetfirstChild(); // <?xml version="1.0" encoding="UTF-8"?>
	MSXML2::IXMLDOMNamedNodeMapPtr pXMLAttributeMap = pXMLFirstChild->Getattributes(); // Двойка (vesrsion, encoding) (1.0, UTF-8) 
	MSXML2::IXMLDOMNodePtr pXMLEncodNode = pXMLAttributeMap->getNamedItem(_T("encoding")); // вземаме елемента encoding
	_bstr_t bstrDefaultEncoding = m_strEncoding;
 pXMLEncodNode->PutnodeValue(bstrDefaultEncoding);

 m_bHasBeenFinilized = TRUE;
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
 hr = S_FALSE;
	}//cacth
	
	return hr;
 }
 HRESULT CXmlBuilder2::SaveToFile( CString strPath)
 {
 HRESULT hr = S_OK;
	try
	{
	if (PathFileExists(strPath))
	return S_FALSE;

 CString strLocation = strPath;
	if (strLocation.IsEmpty())
	return S_FALSE;
 hr = m_pXMLFormattedDoc->save(strLocation.AllocSysString());
	if (FAILED(hr))
	return S_FALSE;
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
 hr = S_FALSE;
	}//cacth
 	return hr;
 }
 MSXML2::IXMLDOMNodePtr CXmlBuilder2::GetRootElem()
{
	try
	{
 INT_PTR nDeep = 0;
 MSXML2::IXMLDOMNodePtr pChild;
	for (pChild = m_pXMLFormattedDoc->firstChild; pChild; pChild = pChild->nextSibling)
 {
	if (nDeep == 1)
	break;
 nDeep++;
	}//for
 	return pChild;
 }
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__); 
	}//cacth
 	return NULL;
 }
 CString CXmlBuilder2::ToSting()
 {
	if (!m_bHasBeenFinilized)
 Finalize();
 CString strXml = m_pXMLFormattedDoc->xml;
	return strXml;
 }
	long CXmlBuilder2::GetCountOfChilds()const
	{
	return m_lCountOfChilds;
 }
 MSXML2::IXMLDOMElementPtr CXmlBuilder2::CreateElement( CString strChildName, CString strValue,
 CString strAttribute1, CString strAttributeValue1,
 CString strAttribute2, CString strAttributeValue2,
 CString strAttribute3, CString strAttributeValue3 )
 {
	try
	{
	if (strChildName.IsEmpty() || strValue.IsEmpty())
	return NULL;
 _bstr_t bstrName = strChildName;
 _bstr_t bstrValue = strValue;
 _bstr_t bstrAttrName;
 _bstr_t bstrAttrValue;
 MSXML2::IXMLDOMElementPtr pXMLChild = m_pXMLDoc->createElement(bstrName);
 pXMLChild->Puttext(bstrValue);
	if (!strAttribute1.IsEmpty() && !strAttributeValue1.IsEmpty())
 {
 bstrAttrName = strAttribute1;
 bstrAttrValue = strAttributeValue1;
 pXMLChild->setAttribute(bstrAttrName, bstrAttrValue);
	}//if
 	if (!strAttribute2.IsEmpty() && !strAttributeValue2.IsEmpty())
 {
 bstrAttrName = strAttribute2;
 bstrAttrValue = strAttributeValue2;
 pXMLChild->setAttribute(bstrAttrName, bstrAttrValue);
	}//if
 	if (!strAttribute3.IsEmpty() && !strAttributeValue3.IsEmpty())
 {
 bstrAttrName = strAttribute3;
 bstrAttrValue = strAttributeValue3;
 pXMLChild->setAttribute(bstrAttrName, bstrAttrValue);
	}//if
 pXMLChild = m_pXMLRootElem->appendChild(pXMLChild);
 m_lCountOfChilds++;
	return pXMLChild;
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
	}//cacth
 	return NULL;
 }
	void CXmlBuilder2::AppendChild(MSXML2::IXMLDOMElementPtr pXMLChild)
 {
	try
	{
	if (!pXMLChild)
	return;
 pXMLChild = m_pXMLRootElem->appendChild(pXMLChild);
 m_lCountOfChilds++;
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
	}//cacth
	}
	MSXML2::IXMLDOMNodePtr CXmlBuilder2::FindElement(const CString& strNodeName, const CString& strParentNodeName )
 {
 _bstr_t bstrNode = m_pXMLDoc->documentElement->nodeName;
 MSXML2::IXMLDOMNodePtr pRootNode = m_pXMLDoc->selectSingleNode(bstrNode);
	if(!pRootNode)
	return NULL;
	try
	{
	//BOOL bDestroyed = FALSE;
	MSXML2::IXMLDOMNodePtr pChild = NULL;
 MSXML2::IXMLDOMNodePtr pMostWanted = NULL;

	for (pChild = pRootNode->firstChild; pChild; pChild = pRootNode->nextSibling)
 { 
	if(pChild)
 { 
 CString strXML = pChild->xml;
 CString strChildNodeName = pChild->nodeName;
 CString strChildParentName = pChild->parentNode->nodeName;

	// Ако има съвпадение
	if (!strNodeName.Compare(pChild->nodeName) && !strParentNodeName.Compare(pChild->parentNode->nodeName))
	return pChild;
	// Ако няма съвпадение претърсваме и под елементите на текущия, и неговите под под елементи и така до края.. рекурсивно обхождаме цялото дърво
	pMostWanted = CXmlBuilder2::Seek(pChild, strNodeName, strParentNodeName);
	if (pMostWanted)
	break;
	}//if
	}//for
 	return pMostWanted;
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
	}//catch
 	return NULL;
 }
	// WTF
	MSXML2::IXMLDOMNodePtr CXmlBuilder2::Seek( MSXML2::IXMLDOMNodePtr pNode, const CString& strNodeName, const CString& strParentNodeName )
 {
	try
	{ 
	//BOOL bFound = FALSE;
	if (!pNode)
	return NULL;
 MSXML2::IXMLDOMNodePtr pChild = NULL; 
 MSXML2::IXMLDOMNodePtr pMostWanted = NULL; 
	// Цикли безкрайно само 1 я child ?!
	for (pChild = pNode->firstChild; pChild; pChild = pNode->nextSibling)
 { 
	if (pChild)
 {
CString strXML = pChild->xml;
 CString strChildNodeName = pChild->nodeName;
 CString strChildParentName = pChild->parentNode->nodeName;
	// при съвпадение прекратяваме търсенето
	if (!strNodeName.Compare(strChildNodeName) && !strParentNodeName.Compare(strChildParentName)) 
	return pChild;

	// ако ням съвпадение , продължаваме рекрсивно
	pMostWanted = CXmlBuilder2::Seek(pChild, strNodeName, strParentNodeName );
	if (pMostWanted)
	return pMostWanted; 
	}//if 
	}//for 
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__); 
	}//cacth
 	return NULL;
 }
	//static 
	MSXML2::IXMLDOMElementPtr CXmlBuilder2::CreateIndependantElement( CString strChildName, CString strValue,
 CString strAttribute1, CString strAttributeValue1,
 CString strAttribute2, CString strAttributeValue2,
 CString strAttribute3, CString strAttributeValue3 )
 {
	try
	{
	CXmlBuilder2 oDummyBulder(_T("IndependantElement"));
	if (oDummyBulder.NewDocument() == S_OK)
 {
 MSXML2::IXMLDOMElementPtr p = oDummyBulder.CreateElement(strChildName, strValue, strAttribute1, strAttributeValue1, strAttribute2, strAttributeValue2, strAttribute3, strAttributeValue3);
	return p;
	}//if
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
	}//cacth
 	return NULL; 
 }
	//static 
	long CXmlBuilder2::GetCountOfChilds(MSXML2::IXMLDOMNodePtr pElem)
 {
	try
	{
	if (!pElem)
	return 0L;
	if (!pElem->firstChild)
	return 0L;
 LONG lCount = 0L;
 MSXML2::IXMLDOMNodePtr pChild;
	for (pChild = pElem->firstChild; pChild; pChild = pElem->nextSibling)
 {
 lCount++;
	}//for
 	return lCount;
	}//try
	catch (_com_error& ex)
 {
 _bstr_t bstrExMsg = ex.ErrorMessage();
 _bstr_t bstrDescr = ex.Description();

	LAPPERROR("[%s] Exception occured: %s Descr: %s", __FUNCTION__, bstrExMsg, bstrDescr ); 
 }
	catch (...)
 {
	LAPPERROR("[%s] Exception occured", __FUNCTION__);
	}//cacth
 	return 0L; 
 }
}//namespace

