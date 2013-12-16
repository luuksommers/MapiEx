#ifndef __MAPIOBJECT_H__
#define __MAPIOBJECT_H__

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIObject.h
// Description: Base class for code common to MAPI Items (messages, contacts etc)
//
// Copyright (C) 2005-2010, Noel Dillabough
//
// This source code is free to use and modify provided this notice remains intact and that any enhancements
// or bug fixes are posted to the CodeProject page hosting this class for the community to benefit.
//
// Usage: see the CodeProject article at http://www.codeproject.com/internet/CMapiEx.asp
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CMAPIEx;
class CMAPIFolder;

/////////////////////////////////////////////////////////////
// CMAPIObject

class AFX_EXT_CLASS CMAPIObject : public CObject
{
public:
	CMAPIObject();
	~CMAPIObject();

// Attributes
protected:
	CMAPIEx* m_pMAPI;
	IMAPIProp* m_pItem;
	SBinary m_entryID;

// Operations
public:
	inline LPMESSAGE Message() { return (LPMESSAGE)m_pItem; }

	SBinary* GetEntryID() { return &m_entryID; }
	BOOL GetEntryIDString(CString& strEntryID);
	void SetEntryID(SBinary* pEntryID=NULL);
	int GetMessageFlags();
	BOOL SetMessageFlags(int nFlags);
	BOOL GetMessageClass(CString& strMessageClass);
	int GetMessageEditorFormat();
	BOOL SetMessageEditorFormat(int nFormat);

	virtual BOOL Open(CMAPIEx* pMAPI,SBinary entry);
	virtual void Close();
	virtual BOOL Save(BOOL bClose=TRUE);

	// Properties
	virtual BOOL GetPropertyString(ULONG ulProperty, CString& strProperty, BOOL bStream=FALSE);
	int GetPropertyValue(ULONG ulProperty, int nDefaultValue);
	BOOL GetNamedProperty(LPCTSTR szFieldName, LPSPropValue& pProp);
	BOOL GetNamedProperty(LPCTSTR szFieldName, CString& strField);
	BOOL GetOutlookProperty(ULONG ulData, ULONG ulProperty, LPSPropValue& pProp);
	BOOL GetOutlookPropertyString(ULONG ulData, ULONG ulProperty, CString& strProperty);
	virtual BOOL SetPropertyString(ULONG ulProperty, LPCTSTR szProperty, BOOL bStream=FALSE);
	BOOL SetNamedProperty(LPCTSTR szFieldName, LPCTSTR szField, BOOL bCreate=TRUE);
	BOOL SetNamedMVProperty(LPCTSTR szFieldName, LPCTSTR* arCategories, int nCount, LPSPropValue &pProp, BOOL bCreate=TRUE);
	BOOL SetOutlookProperty(ULONG ulData, ULONG ulProperty, LPCTSTR szField);
	BOOL SetOutlookProperty(ULONG ulData, ULONG ulProperty, int nField);
	BOOL SetOutlookProperty(ULONG ulData, ULONG ulProperty, FILETIME ftField);
	BOOL DeleteNamedProperty(LPCTSTR szFieldName);

	// Attachments
	int GetAttachmentCount();
	BOOL GetAttachmentCID(CString& strAttachmentCID, int nIndex);
	BOOL GetAttachmentName(CString& strAttachmentName, int nIndex);
	BOOL SaveAttachment(LPCTSTR szFolder, int nIndex=-1, LPCTSTR szFileName=NULL);
	BOOL DeleteAttachment(int nIndex=-1);
	BOOL AddAttachment(LPCTSTR szPath, LPCTSTR szName=NULL, LPCTSTR szCID=NULL);

	// Body
	BOOL GetBody(CString& strBody, BOOL bAutoDetect=TRUE);
	BOOL SetBody(LPCTSTR szBody);
	BOOL GetHTML(CString& strHTML);
	BOOL SetHTML(LPCTSTR szHTML);
	BOOL GetRTF(CString& strRTF);
	BOOL SetRTF(LPCTSTR szRTF);

protected:
	BOOL Create(CMAPIEx* pMAPI, CMAPIFolder* pFolder);
	HRESULT GetProperty(ULONG ulProperty, LPSPropValue &prop);
	BOOL GetPropTagArray(LPCTSTR szFieldName, LPSPropTagArray& lppPropTags, int& nFieldType, BOOL bCreate);
	BOOL GetOutlookPropTagArray(ULONG ulData, ULONG ulProperty, LPSPropTagArray& lppPropTags, int& nFieldType, BOOL bCreate);
	BOOL SaveAttachment(LPATTACH pAttachment, LPCTSTR szPath);
};

#endif
