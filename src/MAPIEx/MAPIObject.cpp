////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIObject.cpp
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

#include "MAPIExPCH.h"
#include "MAPIEx.h"

/////////////////////////////////////////////////////////////
// CMAPIObject

CMAPIObject::CMAPIObject()
{
	m_pMAPI=NULL;
	m_pItem=NULL;
	m_entryID.cb=0;
	SetEntryID(NULL);
}

CMAPIObject::~CMAPIObject()
{
}

BOOL CMAPIObject::GetEntryIDString(CString& strEntryID)
{
	if(m_entryID.cb) 
	{
		strEntryID.Empty();
		TCHAR szBuffer[3];
		for(ULONG i=0;i<m_entryID.cb;i++)
		{
			_stprintf_s(szBuffer, 3, _T("%02X"), m_entryID.lpb[i]);
			strEntryID+=szBuffer;
		}
		return TRUE;
	}
	return FALSE;
}

void CMAPIObject::SetEntryID(SBinary* pEntryID)
{
	if(m_entryID.cb) delete [] m_entryID.lpb;
	m_entryID.lpb=NULL;

	if(pEntryID) 
	{
		m_entryID.cb=pEntryID->cb;
		if(m_entryID.cb) 
		{
			m_entryID.lpb=new BYTE[m_entryID.cb];
			memcpy(m_entryID.lpb, pEntryID->lpb, m_entryID.cb);
		}
	} 
	else 
	{
		m_entryID.cb=0;
	}
}

BOOL CMAPIObject::Open(CMAPIEx* pMAPI,SBinary entryID)
{
	Close();
	m_pMAPI=pMAPI;
	ULONG ulObjType;
	if(m_pMAPI->GetSession()->OpenEntry(entryID.cb, (LPENTRYID)entryID.lpb, NULL, MAPI_BEST_ACCESS, &ulObjType, (LPUNKNOWN*)&m_pItem)!=S_OK) return FALSE;
	SetEntryID(&entryID);
	return TRUE;
}

void CMAPIObject::Close()
{
	SetEntryID(NULL);
	RELEASE(m_pItem);
	m_pMAPI=NULL;
}

BOOL CMAPIObject::Save(BOOL bClose)
{
	ULONG ulFlags=bClose ? 0 : KEEP_OPEN_READWRITE;
	if(m_pItem && m_pItem->SaveChanges(ulFlags)==S_OK) 
	{
		if(bClose) Close();
		return TRUE;
	}
	return FALSE;
}

int CMAPIObject::GetMessageFlags()
{
	return GetPropertyValue(PR_MESSAGE_FLAGS, 0);
}

BOOL CMAPIObject::SetMessageFlags(int nFlags)
{
	SPropValue prop;
	prop.ulPropTag=PR_MESSAGE_FLAGS;
	prop.Value.l=nFlags;
	return (m_pItem && m_pItem->SetProps(1, &prop, NULL)==S_OK);
}

BOOL CMAPIObject::GetMessageClass(CString& strMessageClass)
{
	return GetPropertyString(PR_MESSAGE_CLASS, strMessageClass);
}

int CMAPIObject::GetMessageEditorFormat()
{
	return GetPropertyValue(PR_MSG_EDITOR_FORMAT, EDITOR_FORMAT_DONTKNOW);
}

BOOL CMAPIObject::SetMessageEditorFormat(int nFormat)
{
	SPropValue prop;
	prop.ulPropTag=PR_MSG_EDITOR_FORMAT;
	prop.Value.l=nFormat;
	return (m_pItem && m_pItem->SetProps(1, &prop, NULL)==S_OK);
}

BOOL CMAPIObject::Create(CMAPIEx* pMAPI, CMAPIFolder* pFolder)
{
	if(!pMAPI) return FALSE;
	if(!pFolder) pFolder=pMAPI->GetFolder();
	if(!pFolder) return FALSE;
	Close();
	m_pMAPI=pMAPI;
	if(pFolder->Folder()->CreateMessage(NULL, 0, (LPMESSAGE*)&m_pItem)==S_OK) 
	{
		LPSPropValue pProp;
		if(GetProperty(PR_ENTRYID, pProp)==S_OK) 
		{
			SetEntryID(&pProp->Value.bin);
			MAPIFreeBuffer(pProp);
		} 
		return TRUE;
	}
	return FALSE;
}

HRESULT CMAPIObject::GetProperty(ULONG ulProperty, LPSPropValue& pProp)
{
	if(!m_pItem) return E_INVALIDARG;
	ULONG ulPropCount;
	ULONG p[2]={ 1, ulProperty };
	return m_pItem->GetProps((LPSPropTagArray)p, CMAPIEx::cm_nMAPICode, &ulPropCount, &pProp);
}

BOOL CMAPIObject::GetPropertyString(ULONG ulProperty, CString& strProperty, BOOL bStream)
{
	strProperty=_T("");
	if(bStream)
	{
		IStream* pStream;
		if(Message()->OpenProperty(ulProperty, &IID_IStream,STGM_READ, NULL, (LPUNKNOWN*)&pStream)==S_OK) 
		{
			const int BUF_SIZE=16384;
			TCHAR szBuf[BUF_SIZE+1];
			ULONG ulNumChars;

			do 
			{
				pStream->Read(szBuf, BUF_SIZE*sizeof(TCHAR), &ulNumChars);
				ulNumChars/=sizeof(TCHAR);
				szBuf[min(BUF_SIZE,ulNumChars)]=0;
				strProperty+=szBuf;
			} while(ulNumChars>=BUF_SIZE);

			RELEASE(pStream);
			return TRUE;
		}
	}
	else 
	{
		LPSPropValue pProp;
		if(GetProperty(ulProperty, pProp)==S_OK) 
		{
			strProperty=CMAPIEx::GetValidString(*pProp);
			MAPIFreeBuffer(pProp);
			return TRUE;
		} 
	}
	return FALSE;
}

int CMAPIObject::GetPropertyValue(ULONG ulProperty, int nDefaultValue)
{
	LPSPropValue pProp;
	if(GetProperty(ulProperty, pProp)==S_OK) 
	{
		nDefaultValue=pProp->Value.l;
		MAPIFreeBuffer(pProp);
	}
	return nDefaultValue;
}

const GUID GUIDPublicStrings={0x00020329, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

BOOL CMAPIObject::GetPropTagArray(LPCTSTR szFieldName, LPSPropTagArray& lppPropTags, int& nFieldType, BOOL bCreate)
{
	if(!m_pItem) return FALSE;

	MAPINAMEID nameID;
	nameID.lpguid=(GUID*)&GUIDPublicStrings;
	nameID.ulKind=MNID_STRING;
#ifdef UNICODE
	nFieldType=PT_UNICODE;
	nameID.Kind.lpwstrName=(LPWSTR)szFieldName;
#else
	nFieldType=PT_STRING8;
	WCHAR wszFieldName[256];
	MultiByteToWideChar(CP_ACP, 0, szFieldName,-1,wszFieldName,255);
	nameID.Kind.lpwstrName=wszFieldName;
#endif

	LPMAPINAMEID lpNameID[1]={ &nameID };

	HRESULT hr=m_pItem->GetIDsFromNames(1, lpNameID, bCreate ? MAPI_CREATE : 0, &lppPropTags);
	return (hr==S_OK);
}

BOOL CMAPIObject::GetNamedProperty(LPCTSTR szFieldName, LPSPropValue &pProp)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetPropTagArray(szFieldName, lppPropTags, nFieldType, FALSE)) return FALSE;

	ULONG ulPropCount;
	HRESULT hr=m_pItem->GetProps(lppPropTags, CMAPIEx::cm_nMAPICode, &ulPropCount, &pProp);
	MAPIFreeBuffer(lppPropTags);
	return (hr==S_OK);
}

BOOL CMAPIObject::GetNamedProperty(LPCTSTR szFieldName, CString& strField)
{
	LPSPropValue pProp;
	if(GetNamedProperty(szFieldName, pProp)) 
	{
		strField=CMAPIEx::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		return TRUE;
	}
	return FALSE;
}

BOOL CMAPIObject::GetOutlookPropTagArray(ULONG ulData, ULONG ulProperty, LPSPropTagArray& lppPropTags, int& nFieldType, BOOL bCreate)
{
	if(!m_pItem) return FALSE;

	const GUID guidOutlookEmail1={ulData, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 };

	MAPINAMEID nameID;
	nameID.lpguid=(GUID*)&guidOutlookEmail1;
	nameID.ulKind=MNID_ID;
	nameID.Kind.lID=ulProperty;

#ifdef UNICODE
	nFieldType=PT_UNICODE;
#else
	nFieldType=PT_STRING8;
#endif

	LPMAPINAMEID lpNameID[1]={ &nameID };

	HRESULT hr=m_pItem->GetIDsFromNames(1, lpNameID, bCreate ? MAPI_CREATE : 0, &lppPropTags);
	return (hr==S_OK);
}

// gets a custom outlook property (ie EmailAddress1 of a contact)
BOOL CMAPIObject::GetOutlookProperty(ULONG ulData, ULONG ulProperty, LPSPropValue& pProp)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetOutlookPropTagArray(ulData,ulProperty, lppPropTags, nFieldType, FALSE)) return FALSE;

	ULONG ulPropCount;
	HRESULT hr=m_pItem->GetProps(lppPropTags, CMAPIEx::cm_nMAPICode, &ulPropCount, &pProp);
	MAPIFreeBuffer(lppPropTags);
	return (hr==S_OK);
}

// gets a custom outlook property string (ie EmailAddress1 of a contact)
BOOL CMAPIObject::GetOutlookPropertyString(ULONG ulData, ULONG ulProperty, CString& strProperty)
{
	LPSPropValue pProp;
	if(GetOutlookProperty(ulData, ulProperty, pProp)) 
	{
		strProperty=CMAPIEx::GetValidString(*pProp);
		MAPIFreeBuffer(pProp);
		return TRUE;
	} 
	return FALSE;
}

BOOL CMAPIObject::SetPropertyString(ULONG ulProperty, LPCTSTR szProperty, BOOL bStream)
{
	if(m_pItem && szProperty) 
	{
		if(bStream)
		{
			LPSTREAM pStream=NULL;
			if(m_pItem->OpenProperty(ulProperty, &IID_IStream, 0, MAPI_MODIFY | MAPI_CREATE, (LPUNKNOWN*)&pStream)==S_OK) 
			{
				pStream->Write(szProperty, (ULONG)(_tcslen(szProperty)+1)*sizeof(TCHAR), NULL);
				RELEASE(pStream);
				return TRUE;
			}
		}
		else
		{
			SPropValue prop;
			prop.ulPropTag=ulProperty;
			prop.Value.LPSZ=(LPTSTR)szProperty;
			return (m_pItem->SetProps(1, &prop, NULL)==S_OK);
		}
	}
	return FALSE;
}

// if bCreate is true, the property will be created if necessary otherwise if not present will return FALSE
BOOL CMAPIObject::SetNamedProperty(LPCTSTR szFieldName, LPCTSTR szField, BOOL bCreate)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetPropTagArray(szFieldName, lppPropTags, nFieldType, bCreate)) return FALSE;

	SPropValue prop;
	prop.ulPropTag=(lppPropTags->aulPropTag[0]|nFieldType);
	prop.Value.LPSZ=(LPTSTR)szField;
	HRESULT hr=m_pItem->SetProps(1, &prop, NULL);
	MAPIFreeBuffer(lppPropTags);
	return (hr==S_OK);
}

BOOL CMAPIObject::SetNamedMVProperty(LPCTSTR szFieldName, LPCTSTR* arCategories, int nCount, LPSPropValue &pProp, BOOL bCreate)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetPropTagArray(szFieldName, lppPropTags, nFieldType, bCreate)) return FALSE;

	HRESULT hr=MAPIAllocateBuffer(sizeof(SPropValue), (LPVOID*)&pProp);
	if(hr==S_OK) 
	{
		pProp->ulPropTag=(lppPropTags->aulPropTag[0]|nFieldType|MV_FLAG);
		pProp->Value.MVi.cValues=nCount;
		if(PROP_TYPE(pProp->ulPropTag)==PT_MV_STRING8) 
		{
			hr=MAPIAllocateMore(sizeof(LPSTR)*nCount, pProp, (LPVOID*)&pProp->Value.MVszA.lppszA);
		} 
		else 
		{ 
			// PT_MV_UNICODE
			hr=MAPIAllocateMore(sizeof(LPSTR)*nCount, pProp, (LPVOID*)&pProp->Value.MVszW.lppszW);
		}
		if(hr==S_OK) 
		{
			for(int i=0;i<nCount;i++) pProp->Value.MVSZ.LPPSZ[i]=(LPTSTR)arCategories[i];
			hr=m_pItem->SetProps(1, pProp, NULL);
		}
		if(hr!=S_OK) MAPIFreeBuffer(pProp);
	}
	return (hr==S_OK);
}

BOOL CMAPIObject::SetOutlookProperty(ULONG ulData, ULONG ulProperty, LPCTSTR szField)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetOutlookPropTagArray(ulData,ulProperty, lppPropTags, nFieldType, FALSE)) return FALSE;

	SPropValue prop;
	prop.ulPropTag=(lppPropTags->aulPropTag[0]|nFieldType);
	prop.Value.LPSZ=(LPTSTR)szField;
	HRESULT hr=m_pItem->SetProps(1, &prop, NULL);
	MAPIFreeBuffer(lppPropTags);
	return (hr==S_OK);
}

BOOL CMAPIObject::SetOutlookProperty(ULONG ulData, ULONG ulProperty, int nField)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetOutlookPropTagArray(ulData,ulProperty, lppPropTags, nFieldType, FALSE)) return FALSE;
	nFieldType=PT_LONG;

	SPropValue prop;
	prop.ulPropTag=(lppPropTags->aulPropTag[0]|nFieldType);
	prop.Value.l=nField;
	HRESULT hr=m_pItem->SetProps(1, &prop, NULL);
	MAPIFreeBuffer(lppPropTags);
	return (hr==S_OK);
}

BOOL CMAPIObject::SetOutlookProperty(ULONG ulData, ULONG ulProperty, FILETIME ftField)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetOutlookPropTagArray(ulData,ulProperty, lppPropTags, nFieldType, FALSE)) return FALSE;
	nFieldType=PT_SYSTIME;

	SPropValue prop;
	prop.ulPropTag=(lppPropTags->aulPropTag[0]|nFieldType);
	prop.Value.ft=ftField;
	HRESULT hr=m_pItem->SetProps(1, &prop, NULL);
	MAPIFreeBuffer(lppPropTags);
	return (hr==S_OK);
}

BOOL CMAPIObject::DeleteNamedProperty(LPCTSTR szFieldName)
{
	LPSPropTagArray lppPropTags;
	int nFieldType;
	if(!GetPropTagArray(szFieldName, lppPropTags, nFieldType, FALSE)) return FALSE;

	HRESULT hr=m_pItem->DeleteProps(lppPropTags, NULL);
	MAPIFreeBuffer(lppPropTags);
	return (hr==S_OK);
}

int CMAPIObject::GetAttachmentCount()
{
	ULONG ulCount=0;
	LPSPropValue pProp;
	if(GetProperty(PR_HASATTACH, pProp)==S_OK) 
	{
		if(pProp->Value.b) 
		{
			LPMAPITABLE pAttachTable=NULL;
			if(Message()->GetAttachmentTable(0, &pAttachTable)==S_OK) 
			{
#ifndef _WIN32_WCE
				if(pAttachTable->GetRowCount(0, &ulCount)!=S_OK) ulCount=0;
#else
				enum { PROP_ATTACH_CONTENT_ID, ATTACH_COLS };
				static SizedSPropTagArray(ATTACH_COLS, Columns)={ATTACH_COLS, PR_ATTACH_CONTENT_ID };
				if(pAttachTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK)
				{
					ULONG cRows;
					do {
						LPSRowSet pRows=NULL;
						if(pAttachTable->QueryRows(1, 0, &pRows)!=S_OK) 
						{
							MAPIFreeBuffer(pRows);
							break;
						}
						cRows=pRows->cRows;
						ulCount+=cRows;
						FreeProws(pRows);
					} while(cRows);
				}
#endif
				RELEASE(pAttachTable);
			}
		}
		MAPIFreeBuffer(pProp);
	}
	return ulCount;
}

BOOL CMAPIObject::GetAttachmentCID(CString& strAttachmentCID, int nIndex)
{
	strAttachmentCID="";

	LPMAPITABLE pAttachTable=NULL;
	if(Message()->GetAttachmentTable(0, &pAttachTable)==S_OK)
	{
		ULONG ulCount=GetAttachmentCount();
		if(nIndex<(int)ulCount) 
		{
			enum { PROP_ATTACH_CONTENT_ID, ATTACH_COLS };
			static SizedSPropTagArray(ATTACH_COLS, Columns)={ATTACH_COLS, PR_ATTACH_CONTENT_ID };
			if(pAttachTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK)
			{
				LPSRowSet pRows=NULL;
				if(pAttachTable->QueryRows(ulCount, 0, &pRows)==S_OK) 
				{
					if(nIndex < (int)pRows->cRows)
					{
						if(!CMAPIEx::GetValidString(pRows->aRow[nIndex].lpProps[PROP_ATTACH_CONTENT_ID])) strAttachmentCID=pRows->aRow[nIndex].lpProps[PROP_ATTACH_CONTENT_ID].Value.LPSZ;
					}
					FreeProws(pRows);
					MAPIFreeBuffer(pRows);
				}
			}
		}
		RELEASE(pAttachTable);
	}
	return (!strAttachmentCID.IsEmpty());
}

BOOL CMAPIObject::GetAttachmentName(CString& strAttachmentName, int nIndex)
{
	strAttachmentName=_T("");

	LPMAPITABLE pAttachTable=NULL;
	if(Message()->GetAttachmentTable(0, &pAttachTable)==S_OK)
	{
		ULONG ulCount=GetAttachmentCount();
		if(nIndex<(int)ulCount) 
		{
			enum { PROP_ATTACH_LONG_FILENAME, PROP_ATTACH_FILENAME, ATTACH_COLS };
			static SizedSPropTagArray(ATTACH_COLS, Columns)={ATTACH_COLS, PR_ATTACH_LONG_FILENAME, PR_ATTACH_FILENAME };
			if(pAttachTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK)
			{
				LPSRowSet pRows=NULL;
				if(pAttachTable->QueryRows(ulCount, 0, &pRows)==S_OK) 
				{
					if(nIndex < (int)pRows->cRows)
					{
						if(CMAPIEx::GetValidString(pRows->aRow[nIndex].lpProps[PROP_ATTACH_LONG_FILENAME])) strAttachmentName=pRows->aRow[nIndex].lpProps[PROP_ATTACH_LONG_FILENAME].Value.LPSZ;
						else if(CMAPIEx::GetValidString(pRows->aRow[nIndex].lpProps[PROP_ATTACH_FILENAME])) strAttachmentName=pRows->aRow[nIndex].lpProps[PROP_ATTACH_FILENAME].Value.LPSZ;
					}
					FreeProws(pRows);
					MAPIFreeBuffer(pRows);
				}
			}
		}
		RELEASE(pAttachTable);
	}
	return (!strAttachmentName.IsEmpty());
}

BOOL CMAPIObject::SaveAttachment(LPATTACH pAttachment, LPCTSTR szPath)
{
	CFile file;
	if(!file.Open(szPath, CFile::modeCreate | CFile::modeWrite)) return FALSE;

	IStream* pStream;
	if(pAttachment->OpenProperty(PR_ATTACH_DATA_BIN, &IID_IStream,STGM_READ, NULL, (LPUNKNOWN*)&pStream)!=S_OK) 
	{
		file.Close();
		return FALSE;
	}

	const int BUF_SIZE=4096;
	BYTE b[BUF_SIZE];
	ULONG ulRead;

	do {
		pStream->Read(&b, BUF_SIZE, &ulRead);
		if(ulRead) file.Write(b,ulRead);
	} while(ulRead>=BUF_SIZE);

	file.Close();
	RELEASE(pStream);
	return TRUE;
}

// use nIndex of -1 to save all attachments to szFolder
BOOL CMAPIObject::SaveAttachment(LPCTSTR szFolder, int nIndex, LPCTSTR szFileName)
{
	LPMAPITABLE pAttachTable=NULL;
	if(Message()->GetAttachmentTable(0, &pAttachTable)!=S_OK) return FALSE;

	CString strPath;
	BOOL bResult=FALSE;
	enum { PROP_ATTACH_NUM, PROP_ATTACH_LONG_FILENAME, PROP_ATTACH_FILENAME, ATTACH_COLS };
	static SizedSPropTagArray(ATTACH_COLS, Columns)={ATTACH_COLS, PR_ATTACH_NUM, PR_ATTACH_LONG_FILENAME, PR_ATTACH_FILENAME };
	if(pAttachTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK)
	{
		int i=0;
		LPSRowSet pRows=NULL;
		while(TRUE)
		{
			if(pAttachTable->QueryRows(1, 0, &pRows)==S_OK)
			{
				if(!pRows->cRows) FreeProws(pRows);
				else if(i<nIndex)
				{
					i++;
					continue;
				}
				else
				{
					LPATTACH pAttachment;
					if(Message()->OpenAttach(pRows->aRow[0].lpProps[PROP_ATTACH_NUM].Value.bin.cb, NULL, 0, &pAttachment)==S_OK)
					{
						if (szFileName != NULL)
						{
							strPath.Format(_T("%s\\%s"), szFolder,szFileName);
						}
						else
						{
							if(CMAPIEx::GetValidString(pRows->aRow[0].lpProps[PROP_ATTACH_LONG_FILENAME])) strPath.Format(_T("%s\\%s"), szFolder, pRows->aRow[0].lpProps[PROP_ATTACH_LONG_FILENAME].Value.LPSZ);
							else if(CMAPIEx::GetValidString(pRows->aRow[0].lpProps[PROP_ATTACH_FILENAME])) strPath.Format(_T("%s\\%s"), szFolder, pRows->aRow[0].lpProps[PROP_ATTACH_FILENAME].Value.LPSZ);
							else strPath.Format(_T("%s\\Attachment.dat"), szFolder);
						}
						if(!SaveAttachment(pAttachment, strPath))
						{
							pAttachment->Release();
							FreeProws(pRows);
							RELEASE(pAttachTable);
							return FALSE;
						}
						bResult=TRUE;
						pAttachment->Release();
					}

					FreeProws(pRows);
					MAPIFreeBuffer(pRows);
					if(nIndex==-1) continue;
				}
			}
			break;
		}
	}
	RELEASE(pAttachTable);
	return bResult;
}

// use nIndex of -1 to delete all attachments
BOOL CMAPIObject::DeleteAttachment(int nIndex)
{
	LPMAPITABLE pAttachTable=NULL;
	if(Message()->GetAttachmentTable(0, &pAttachTable)!=S_OK) return FALSE;

	BOOL bResult=FALSE;
	enum { PROP_ATTACH_NUM, ATTACH_COLS };
	static SizedSPropTagArray(ATTACH_COLS, Columns)={ATTACH_COLS, PR_ATTACH_NUM };
	if(pAttachTable->SetColumns((LPSPropTagArray)&Columns, 0)==S_OK) 
	{
		int i=0;
		LPSRowSet pRows=NULL;
		while(TRUE) 
		{
			if(pAttachTable->QueryRows(1, 0, &pRows)==S_OK) 
			{
				if(!pRows->cRows) FreeProws(pRows);
				else if(i<nIndex) 
				{
					i++;
					continue;
				} 
				else 
				{
					if(Message()->DeleteAttach(pRows->aRow[0].lpProps[PROP_ATTACH_NUM].Value.bin.cb, 0, NULL, 0)!=S_OK) 
					{
						FreeProws(pRows);
						RELEASE(pAttachTable);
						return FALSE;
					}
					bResult=TRUE;
					FreeProws(pRows);
					if(nIndex==-1) continue;
				}
				MAPIFreeBuffer(pRows);
			}
			break;
		}
	}
	RELEASE(pAttachTable);
	return bResult;
}

BOOL CMAPIObject::AddAttachment(LPCTSTR szPath, LPCTSTR szName, LPCTSTR szCID)
{
	if(!Message()) return FALSE;

	IAttach* pAttachment=NULL;
	ULONG ulAttachmentNum=0;

	CFile file;
	if(!file.Open(szPath, CFile::modeRead)) return FALSE;

	LPTSTR szFileName=(LPTSTR)szName;
	if(!szFileName) 
	{
		szFileName=(LPTSTR)szPath;
		for(int i=(int)_tcsclen(szPath)-1;i>=0;i--) if(szPath[i]=='\\' || szPath[i]=='/') 
		{
			szFileName=(LPTSTR)&szPath[i+1];
			break;
		}
	}

	if(Message()->CreateAttach(NULL, 0, &ulAttachmentNum, &pAttachment)!=S_OK) 
	{
		file.Close();
		return FALSE;
	}

	const int nProperties=5;
	SPropValue prop[nProperties];
	memset(prop, 0,sizeof(SPropValue)*nProperties);
	prop[0].ulPropTag=PR_ATTACH_METHOD;
	prop[0].Value.ul=ATTACH_BY_VALUE;
	prop[1].ulPropTag=PR_RENDERING_POSITION;
	prop[1].Value.l=-1;

	if(!szCID || _tcscmp(szCID, CONTACT_PICTURE)) 
	{
		prop[2].ulPropTag=PR_ATTACH_LONG_FILENAME;
		prop[2].Value.LPSZ=(TCHAR*)szFileName;
		prop[3].ulPropTag=PR_ATTACH_FILENAME;
		prop[3].Value.LPSZ=(TCHAR*)szFileName;
		if(!szCID)
		{
			prop[4].ulPropTag=PR_NULL;
		}
		else
		{
			prop[4].ulPropTag=PR_ATTACH_CONTENT_ID;
			prop[4].Value.LPSZ=(TCHAR*)szCID;
		}
	} 
	else 
	{
		prop[2].ulPropTag=PR_ATTACH_LONG_FILENAME;
		prop[2].Value.LPSZ=(TCHAR*)szCID;
		prop[3].ulPropTag=PR_ATTACH_FILENAME;
		prop[3].Value.LPSZ=(TCHAR*)szCID;
		prop[4].ulPropTag=0x7FFF000B;
		prop[4].Value.b=TRUE;
	}

	if(pAttachment->SetProps(nProperties, prop, NULL)==S_OK) 
	{
		LPSTREAM pStream=NULL;
		if(pAttachment->OpenProperty(PR_ATTACH_DATA_BIN, &IID_IStream, 0, MAPI_MODIFY | MAPI_CREATE, (LPUNKNOWN*)&pStream)==S_OK) 
		{
			const int BUF_SIZE=4096;
			BYTE pData[BUF_SIZE];
			ULONG ulSize=0,ulRead,ulWritten;

			ulRead=file.Read(pData, BUF_SIZE);
			while(ulRead) 
			{
				pStream->Write(pData,ulRead, &ulWritten);
				ulSize+=ulRead;
				ulRead=file.Read(pData, BUF_SIZE);
			}

			pStream->Commit(STGC_DEFAULT);
			RELEASE(pStream);
			file.Close();

			prop[0].ulPropTag=PR_ATTACH_SIZE;
			prop[0].Value.ul=ulSize;
			pAttachment->SetProps(1, prop, NULL);

			pAttachment->SaveChanges(KEEP_OPEN_READONLY);
			RELEASE(pAttachment);
			return TRUE;
		}
	}

	file.Close();
	RELEASE(pAttachment);
	return FALSE;
}

// Gets the body of the item, if bAutoDetect is set, uses the MessageEditorFormat to try to determine which property to return 
BOOL CMAPIObject::GetBody(CString& strBody, BOOL bAutoDetect)
{
	if(bAutoDetect)
	{
		int nFormat=GetMessageEditorFormat();
		if(nFormat==EDITOR_FORMAT_DONTKNOW || nFormat==EDITOR_FORMAT_RTF) return GetRTF(strBody);
		else if(nFormat==EDITOR_FORMAT_HTML) return GetHTML(strBody);
	}
	return GetPropertyString(PR_BODY, strBody, TRUE);
}

BOOL CMAPIObject::SetBody(LPCTSTR szBody)
{
	if(SetPropertyString(PR_BODY, szBody, TRUE))
	{
		SetMessageEditorFormat(EDITOR_FORMAT_PLAINTEXT);
		return TRUE;
	}
	return FALSE;
}

BOOL CMAPIObject::GetHTML(CString& strHTML)
{
	return GetPropertyString(PR_BODY_HTML, strHTML, TRUE);
}

BOOL CMAPIObject::SetHTML(LPCTSTR szHTML)
{
	if(szHTML)
	{
		// does this Message Store support HTML directly?
		ULONG ulSupport=m_pMAPI ? m_pMAPI->GetMessageStoreSupport() : 0;
		if(ulSupport&STORE_HTML_OK) 
		{
			if(SetPropertyString(PR_BODY_HTML, szHTML, TRUE))
			{
				SetMessageEditorFormat(EDITOR_FORMAT_HTML);
				return TRUE;
			}
		} 
		else 
		{
			// otherwise lets encode it into RTF 
			TCHAR szCodePage[6]=_T("1252"); // default codepage is ANSI - Latin I
			GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, LOCALE_IDEFAULTANSICODEPAGE, szCodePage, sizeof(szCodePage));

			CString strRTF;
			strRTF.Format(_T("{\\rtf1\\ansi\\ansicpg%s\\fromhtml1 {\\*\\htmltag1 "), szCodePage);
			strRTF+=szHTML;
			strRTF+=_T(" }}");

			return SetRTF(strRTF);
		}
	}
	return FALSE;
}

BOOL CMAPIObject::GetRTF(CString& strRTF)
{
	strRTF=_T("");
	IStream* pStream;

	const int BUF_SIZE=16384;
	char szBuf[BUF_SIZE+1];
	ULONG ulNumChars;

#ifdef _WIN32_WCE
	int nMessageStatus=GetMessageStatus();
	if(nMessageStatus & MSGSTATUS_PARTIAL) return;

	if(nMessageStatus & MSGSTATUS_HAS_PR_BODY_HTML) 
	{
		if(Message()->OpenProperty(PR_BODY_HTML_A, &IID_IStream, STGM_READ, NULL, (LPUNKNOWN*)&pStream)!=S_OK) return FALSE;

		do 
		{
			pStream->Read(szBuf, BUF_SIZE, &ulNumChars);
			szBuf[min(BUF_SIZE,ulNumChars)]=0;
			strRTF+=szBuf;
		} while(ulNumChars>=BUF_SIZE);

		RELEASE(pStream);
	} 
	else if(nMessageStatus & MSGSTATUS_HAS_PR_BODY) 
	{
		if(Message()->OpenProperty(PR_BODY_A, &IID_IStream, STGM_READ, NULL, (LPUNKNOWN*)&pStream)!=S_OK) return FALSE;

		do 
		{
			pStream->Read(szBuf, BUF_SIZE, &ulNumChars);
			szBuf[min(BUF_SIZE,ulNumChars)]=0;
			strRTF+=szBuf;
		} while(ulNumChars>=BUF_SIZE);

		RELEASE(pStream);
	}
#else
	if(Message()->OpenProperty(PR_RTF_COMPRESSED, &IID_IStream,STGM_READ, 0, (LPUNKNOWN*)&pStream)!=S_OK) return FALSE;

	IStream *pUncompressed;
	if(WrapCompressedRTFStream(pStream, 0, &pUncompressed)==S_OK) 
	{
		do 
		{
			pUncompressed->Read(szBuf, BUF_SIZE, &ulNumChars);
			szBuf[min(BUF_SIZE,ulNumChars)]=0;
			strRTF+=szBuf;
		} while(ulNumChars>=BUF_SIZE);
		RELEASE(pUncompressed);
	}

	RELEASE(pStream);
#endif

	LPCTSTR s;
	CString strText;

	// does this RTF contain encoded HTML? If so decode it
	// code taken from Lucian Wischik's example at http://www.wischik.com/lu/programmer/mapi_utils.html
	if(strRTF.Find(_T("\\fromhtml"))!=-1) 
	{
		s=strRTF;

		// scan to <html tag
		// Ignore { and }. These are part of RTF markup.
		// Ignore \htmlrtf...\htmlrtf0. This is how RTF keeps its equivalent markup separate from the html.
		// Ignore \r and \n. The real carriage returns are stored in \par tags.
		// Ignore \pntext{..} and \liN and \fi-N. These are RTF junk.
		// Convert \par and \tab into \r\n and \t
		// Convert \'XX into the ascii character indicated by the hex number XX
		// Convert \{ and \} into { and }. This is how RTF escapes its curly braces.
		// When we get \*\mhtmltagN, keep the tag, but ignore the subsequent \*\htmltagN
		// When we get \*\htmltagN, keep the tag as long as it isn't subsequent to a \*\mhtmltagN
		// All other text should be kept as it is.

		while(*s) 
		{
			if(!_tcsnccmp(s, _T("<html"),5) || !_tcsnccmp(s, _T("\\*\\htmltag"),10)) break;
			s++;
		}

		int nTag=-1, nIgnoreTag=-1;
		while(*s) 
		{
			if(*s==(TCHAR)'{') s++;
			else if(*s==(TCHAR)'}') s++;
			else if(*s==(TCHAR)'\r' || *s==(TCHAR)'\n') s++;
			else if(!_tcsnccmp(s, _T("\\*\\htmltag"),10)) 
			{
				s+=10;
				nTag=0;
				while(*s>=(TCHAR)'0' && *s<=(TCHAR)'9') 
				{
					nTag=nTag*10+*s-(TCHAR)'0';
					s++;
				}
				if(*s==(TCHAR)' ') s++;
				if(nTag==nIgnoreTag) 
				{
					while(*s) 
					{
						if(*s==(TCHAR)'}') break;
						s++;
					}
					nIgnoreTag=-1;
				}
			} 
			else if(_tcsnccmp(s, _T("\\*\\mhtmltag"),11)==0) 
			{ 
				s+=11; 
				while(*s>=(TCHAR)'0' && *s<=(TCHAR)'9') 
				{
					nTag=nTag*10+*s-(TCHAR)'0';
					s++;
				}
				if(*s==(TCHAR)' ') s++;
				nIgnoreTag=nTag;
			} 
			else if(_tcsnccmp(s, _T("\\par"),4)==0) 
			{
				strText+=_T("\r\n");
				s+=4;
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\tab"),4)==0) 
			{
				strText+=_T("\t");
				s+=4;
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\li"),3)==0) 
			{ 
				s+=3; 
				while(*s>=(TCHAR)'0' && *s<=(TCHAR)'9') s++; 
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\fi-"),4)==0) 
			{ 
				s+=4; 
				while(*s>=(TCHAR)'0' && *s<=(TCHAR)'9') s++; 
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\'"),2)==0) 
			{ 
				TCHAR hi=s[2], lo=s[3];
				if(hi>='0' && hi<='9') hi-='0'; else if(hi>='A' && hi<='Z') hi=hi-'A'+10; else if(hi>='a' && hi<='z') hi=hi-'a'+10;
				if(lo>='0' && lo<='9') lo-='0'; else if(lo>='A' && lo<='Z') lo=lo-'A'+10; else if(lo>='a' && lo<='z') lo=lo-'a'+10;
				strText+=(TCHAR)(hi*16+lo);
				s+=4;
			} 
			else if(_tcsnccmp(s, _T("\\pntext"),7)==0) 
			{
				s+=7; 
				while(*s) 
				{
					if(*s==(TCHAR)'}') break;
					s++;
				}
			} 
			else if(_tcsnccmp(s, _T("\\htmlrtf"),8)==0) 
			{
				s+=8;
				while(*s) 
				{
					if(_tcsnccmp(s, _T("\\htmlrtf0"),9)==0) 
					{
						s+=9;
						if(*s==(TCHAR)' ') s++;
						break;
					}
					s++;
				}
			} 
			else if(_tcsnccmp(s, _T("\\{"),2)==0) 
			{ 
				strText+='{';
				s+=2;
			} 
			else if(_tcsnccmp(s, _T("\\}"),2)==0) 
			{ 
				strText+='}';
				s+=2;
			} 
			else 
			{
				strText+=*s;
				s++;
			}
		}
		strRTF=strText;
	} 
	else if(strRTF.Find(_T("\\fromtext"))!=-1) 
	{
		s=strRTF;

		// similar to above we strip out the RTF from the text message that happens to be wrapped in RTF
		while(*s) 
		{
			if(!_tcsnccmp(s, _T("\\fs20"),5)) 
			{
				s+=5;
				if(*s==(TCHAR)' ') s++;
				break;
			}
			s++;
		}

		while(*s) 
		{
			if(*s==(TCHAR)'{') s++;
			else if(*s==(TCHAR)'}') s++;
			else if(*s==(TCHAR)'\r' || *s==(TCHAR)'\n') s++;
			else if(_tcsnccmp(s, _T("\\par"),4)==0) 
			{
				strText+=_T("\r\n");
				s+=4;
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\tab"),4)==0) 
			{
				strText+=_T("\t");
				s+=4;
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\li"),3)==0) 
			{ 
				s+=3; 
				while(*s>=(TCHAR)'0' && *s<=(TCHAR)'9') s++; 
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\fi-"),4)==0) 
			{ 
				s+=4; 
				while(*s>=(TCHAR)'0' && *s<=(TCHAR)'9') s++; 
				if(*s==(TCHAR)' ') s++;
			} 
			else if(_tcsnccmp(s, _T("\\'"),2)==0) 
			{ 
				TCHAR hi=s[2], lo=s[3];
				if(hi>='0' && hi<='9') hi-='0'; else if(hi>='A' && hi<='Z') hi=hi-'A'+10; else if(hi>='a' && hi<='z') hi=hi-'a'+10;
				if(lo>='0' && lo<='9') lo-='0'; else if(lo>='A' && lo<='Z') lo=lo-'A'+10; else if(lo>='a' && lo<='z') lo=lo-'a'+10;
				strText+=(TCHAR)(hi*16+lo);
				s+=4;
			} 
			else if(_tcsnccmp(s, _T("\\pntext"),7)==0) 
			{
				s+=7; 
				while(*s) 
				{
					if(*s==(TCHAR)'}') break;
					s++;
				}
			} 
			else if(_tcsnccmp(s, _T("\\{"),2)==0) 
			{ 
				strText+='{';
				s+=2;
			} 
			else if(_tcsnccmp(s, _T("\\}"),2)==0) 
			{ 
				strText+='}';
				s+=2;
			} 
			else 
			{
				strText+=*s;
				s++;
			}
		}
		strRTF=strText;
	}

	return TRUE;
}

BOOL CMAPIObject::SetRTF(LPCTSTR szRTF)
{
	LPSTREAM pStream=NULL;
	if(Message()->OpenProperty(PR_RTF_COMPRESSED, &IID_IStream, STGM_CREATE | STGM_WRITE, MAPI_MODIFY | MAPI_CREATE, (LPUNKNOWN*)&pStream)==S_OK) 
	{
#ifndef _WIN32_WCE
		IStream *pUncompressed;
		if(WrapCompressedRTFStream(pStream,MAPI_MODIFY, &pUncompressed)==S_OK) 
		{
			pUncompressed->Write(szRTF, (ULONG)_tcslen(szRTF)*sizeof(TCHAR), NULL);
			if(pUncompressed->Commit(STGC_DEFAULT)==S_OK) pStream->Commit(STGC_DEFAULT);
			RELEASE(pUncompressed);
		}
#endif
		RELEASE(pStream);
		SetMessageEditorFormat(EDITOR_FORMAT_RTF);
		return TRUE;
	}
	return FALSE;
}
