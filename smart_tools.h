<!-- .hmmessage P { margin:0px; padding:0px } body.hmmessage { font-size: 12pt; font-family:Calibri } --> 
#pragma once 
#include <vector> 
#include <memory>
 #define ISODate 0 
#define ISODateTime 1
#define DB_DateTime 2
#define DT_Regular 3

typedef MSXML2::IXMLDOMElementPtr PXMLElement; 
typedef MSXML2::IXMLDOMNodePtr PXMLRootNode;
namespace smart_tools
 { 
	extern CString ToISODateTime(const TIMESTAMP_STRUCT dt); 
	extern CString GetCurDateTime(INT_PTR nFromat = DT_Regular);
	extern CString GetStringSmall(const char *pszFormat, ...); 
	extern CString ToCString(char* pszData);
	extern CString ToCString(char cData);
	extern CString ToCString(INT_PTR nData);
	extern CString ToCString(long lData);
	extern CString ToCString(ULONG lData);
	extern CString ToCString(double dData);
	extern CString EscapeSpecials(CString str); 
	extern std::vector<BYTE> ReadFileBinary(const char* pszFileName);
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//CAutoPtrCMapStringToOb
 	template<class T_VALUE>
	class CAutoPtrCMapStringToOb : public CMapStringToOb
 {
	public:
 CAutoPtrCMapStringToOb(BOOL bAutoRelease = TRUE, INT_PTR nBlockSize = 10)
 : CMapStringToOb(nBlockSize), m_bAutoRelease(bAutoRelease)
 {
 }
 ~CAutoPtrCMapStringToOb()
 {
	if (m_bAutoRelease)
 DeleteAll();
 }
	void DeleteAll()
 {
 CString strKey;
 CObject* pPtr = NULL;
	for (POSITION pos = GetStartPosition(); pos != NULL;)
 {
 GetNextAssoc(pos, strKey, pPtr);
	if (pPtr)
	delete pPtr;
	}//for 
	}
	private:
 BOOL m_bAutoRelease;
 };
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//CStringStack
 	class CStringStack : public CStringList
 {
	public:
 CStringStack(BOOL bPopFrontRemoveHeadAllowed = TRUE) :CStringList()
 {
 m_bPopFrontRemoveHeadAllowed = bPopFrontRemoveHeadAllowed;
 }
	// Add element to back of stack
	void PushBack(CString strData)
 {
 AddTail(strData);
 }
	// Peek at top element of stack
	CString Peek()
 {
	if( IsEmpty()) 
	return NULL;
 POSITION pos = GetHeadPosition();
 CString strData = GetNext(pos);

	return strData;
 }
	// Pop top element off stack
	CString PopFront()
 {
	if (IsEmpty())
	return NULL;
 POSITION pos = GetHeadPosition();
 CString strData = GetNext(pos);

	if (m_bPopFrontRemoveHeadAllowed)
 RemoveHead();
	return strData;
 }
 BOOL m_bPopFrontRemoveHeadAllowed;
 };
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//CAutoDestrPtrCMap
 	template<class TYPE_ARG1, class TYPE_ARG2>
	class CAutoPtrCMap : public CMap<TYPE_ARG1, TYPE_ARG1, TYPE_ARG2, TYPE_ARG2 >
 {
	public:
 CAutoPtrCMap(BOOL bAutoRelease = TRUE, INT_PTR nBlockSize = 10)
 : CMap(nBlockSize)
 m_bAutoRelease(bAutoRelease)
 {
 }
 ~CAutoPtrCMap()
 {
	if (m_bAutoRelease)
 DeleteAll();
 }
	void DeleteAll()
 {
	for (POSITION oPosition = GetStartPosition(); oPosition != NULL;)
 {
 KEY oKey;
 VALUE oValue;
 GetNextAssoc(oPosition, oKey, oValue);
	delete oValue;
	}//for
	}
	private:
 BOOL m_bAutoRelease;
 };
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//CStack
 	template<class TYPE_T>
	class CStack : public CTypedPtrList< CObList, TYPE_T >
 {
	public:
	void PushBack(TYPE_T objData)
 {
 AddTail(objData);
 }
	void PushFront(TYPE_T objData)
 {
 AddHead(objData);
 }
	// Peek at top element of stack
	TYPE_T* Peek()
 {
	return IsEmpty() ? NULL : GetHead();
 }
 TYPE_T PopBack()
 {
 TYPE_T objResult;
	if (GetCount() > 0)
 {
 objResult = RemoveTail();
	}//if
 	return objResult;
 }
 TYPE_T PopFront()
 {
 TYPE_T objResult;
	if (GetCount() > 0)
 {
 objResult = RemoveHead();
	}//if
 	return objResult;
 }
 };

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//CXmlParser
 	class CXmlParser
 {
	public:
 CXmlParser()
 {
m_pXMLDom = NULL;
	m_strFileName = _T("");
 m_bCashFragments = TRUE;
 m_nEncoding = ENCODING_UTF_8; 
 }
	virtual ~CXmlParser()
 { 
 }

	MSXML2::IXMLDOMDocumentPtr GetXmlDoc() const { return m_pXMLDom; }
	CString GetOriginalXml(){ return m_strXmlOriginal; }
	CString GetFileName(){ return m_strFileName; }
 BOOL DoesNodeExists( CString strNodeName);
	void SetCashFragments(BOOL bUseCash) { m_bCashFragments = bUseCash; } // Performance FLAG - ползва се за вътрешни нужди, ДА НЕ СЕ извиква НИКОГА НИКЪДЕ
	
	HRESULT Parse( CString strXml, INT_PTR nEncoding = ENCODING_UTF_8); // парсира XML string 
	HRESULT LoadAndParse(CString strPathToXmlFile, INT_PTR nEncoding = ENCODING_UTF_8); // зарежда xml'от файл
	
	CString GetXmlFragment( CString strParentNodeName, CString strChildNodeName ); // връща xml fragment
	INT_PTR GetCountOfFragment( CString strChildNodeName, CString strParentNodeName); // връща броя на раздробените фрагменти от xml'a (вариация)
	INT_PTR GetCountOfValues( CString strChildNodeName, CString strParentNodeName); // връща бройя на простите данни във фрагмент
	
	CString GetValue( CString strChildNodeName, CString strParentNodeName, CString strParentParentNodeName = _T("") );
 CString GetAttributeValue( CString strNodeName, CString strParentNodeName, CString strParentParentNodeName, CString strAttributeName);

	private: 
 CString _GetValue( CString strChildNodeName, CString strParentNodeName);
 CString _GetAttributeValue( CString strNodeName, CString strAttributeName);
	static CString SttGetValue( CString strXml, CString strSubNodeName, CString strParentNode, INT_PTR nEncoding = ENCODING_UTF_8);
	static CString SttGetAttributeValue( CString strXml, CString strSubNodeName, CString strArtibuteName, INT_PTR nEncoding = ENCODING_UTF_8);
 CString GetRootNodeName( CString strXml, INT_PTR nEncoding = ENCODING_UTF_8);
 BOOL InitXml( CString strXml, INT_PTR nEncoding, MSXML2::IXMLDOMDocumentPtr& pXMLDom); 

	HRESULT Parse( CString strXml, // XML фрагмент
	CString strNodeName, // име на root елемента
	CString strParentNodeName, // име на бащияния елемент(ако има такъв)
	INT_PTR& nRecursionDeepLevel, // дълбочина на рекурсията, при ексепшън се пише в лога
	INT_PTR nEncoding = ENCODING_UTF_8 );
	void DoStackMappingFragments(CString strKey, CString strValue); // пълни масива със фргменти
	void DoStackMappingValues(CString strKey, CString strValue); // пълни масива със стойности
	void DoStackMappingAttributeValues(CString strKey, CString strValue); // пълни масива със стойности на атрибутите
 	private: 
	MSXML2::IXMLDOMDocumentPtr m_pXMLDom; // указател към документа
	INT_PTR m_nEncoding;
 CString m_strFileName; 
 CString m_strXmlOriginal;
 BOOL m_bCashFragments; 
	std::vector<CString> m_vNodeNames; // списък с имената на всички нодове(позволява повторения)
	CAutoPtrCMapStringToOb<CStringStack> m_mapStackElements; // опашка сложни фрагменти 
	CAutoPtrCMapStringToOb<CStringStack> m_mapStackNodeValues; // опашка прости стойности 
	CAutoPtrCMapStringToOb<CStringStack> m_mapStackNodeAttributeValues; // опашка атрибути 
	};
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//CXmlBuilder2
 
	class CXmlBuilder2
 {
	public: 
 CXmlBuilder2();
	CXmlBuilder2(CString strRootNodeName, CString strEncoding = _T("UTF-8"));
 ~CXmlBuilder2();
 BOOL InitXml( CString strXml, INT_PTR nEncoding, MSXML2::IXMLDOMDocument2Ptr& pXMLDom);
 BOOL LoadXml(CString strXml, INT_PTR nEncoding = ENCODING_UTF_8);

	HRESULT NewDocument( CString strRootAttr1 = _T(""), CString strRootAttrValue1 = _T(""), 
	CString strRootAttr2 = _T(""), CString strRootAttrValue2 = _T(""),
	CString strRootAttr3 = _T(""), CString strRootAttrValue3 = _T("") );
 HRESULT Finalize(); 
 HRESULT SaveToFile(CString strPath);

 CString ToSting(); 
	long GetCountOfChilds()const;
	MSXML2::IXMLDOMNodePtr GetRootElem(); // Да се вика само след Finalize
	
	void AppendChild(MSXML2::IXMLDOMElementPtr pXMLChild); 
 MSXML2::IXMLDOMElementPtr CreateElement( CString strChildName, CString strValue,
	CString strAttribute1 = _T(""), CString strAttributeValue1 = _T(""),
	CString strAttribute2 = _T(""), CString strAttributeValue2 = _T(""),
	CString strAttribute3 = _T(""), CString strAttributeValue3 = _T("") );

	static MSXML2::IXMLDOMElementPtr CreateIndependantElement( CString strChildName, CString strValue,
	CString strAttribute1 = _T(""), CString strAttributeValue1 = _T(""),
	CString strAttribute2 = _T(""), CString strAttributeValue2 = _T(""),
	CString strAttribute3 = _T(""), CString strAttributeValue3 = _T("") ); 
	static long GetCountOfChilds(MSXML2::IXMLDOMNodePtr pElem); 
	MSXML2::IXMLDOMNodePtr FindElement(const CString& strNodeName, const CString& strParentNodeName );

	private: 
	// not working
	static MSXML2::IXMLDOMNodePtr Seek( MSXML2::IXMLDOMNodePtr pNode, const CString& strNodeName, const CString& strParentNodeName );

	long	m_lCountOfChilds;
 BOOL m_bHasBeenFinilized;
 CString m_strRootNodeName;
 CString m_strEncoding;
 MSXML2::IXMLDOMDocument2Ptr m_pXMLDoc; 
 MSXML2::IXMLDOMDocument2Ptr m_pLoadXML;
 MSXML2::IXMLDOMDocument2Ptr m_pXMLFormattedDoc; 
 MSXML2::IXMLDOMElementPtr m_pXMLRootElem;
 MSXML2::IXMLDOMProcessingInstructionPtr m_pXMLProcessingNode;

 };

}//namespace

