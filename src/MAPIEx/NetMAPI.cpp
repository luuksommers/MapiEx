////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: NetMAPI.cpp
// Description: Exported functions for Extended MAPI, meant to be called by .NET
//
// Copyright (C) 2005-2010, Noel Dillabough
//
// This source code is free to use and modify provided this notice remains intact and that any enhancements
// or bug fixes are posted to the CodeProject page hosting this class for all to benefit.
//
// Usage: see the CodeProject article at http://www.codeproject.com/internet/CMapiEx.asp
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

#include "MAPIExPCH.h"
#include "NetMAPI.h"

// Initialize and Terminate

BOOL MAPIInit(BOOL bMultiThreadedNotifications, BOOL bInitAsService)
{
	return CMAPIEx::Init(bMultiThreadedNotifications, bInitAsService);
}

void MAPITerm()
{
	CMAPIEx::Term();
}

void CopyString(LPTSTR szDest, LPCTSTR szSrc, int nMaxLength)
{
	int nSrcLen=(int)_tcslen(szSrc)+1;
	int nLen=min(nSrcLen, nMaxLength);
	memcpy(szDest, szSrc,sizeof(TCHAR)*nLen);
	if(nLen==nMaxLength) szDest[nMaxLength-1]=0;
}

// Profiles, Message Store

CMAPIEx* MAPILogin(LPCTSTR szProfile, BOOL bInitAsService)
{
	if(szProfile && !_tcslen(szProfile)) szProfile=NULL;
	CMAPIEx* pMAPI=new CMAPIEx();
	if(pMAPI->Login(szProfile, bInitAsService)) return pMAPI;
	delete pMAPI;
	return NULL;
}

void MAPILogout(CMAPIEx* pMAPI)
{
	pMAPI->Logout();
	delete pMAPI;
}

BOOL MAPIOpenMessageStore(CMAPIEx* pMAPI, LPCTSTR szStore)
{
	if(szStore && !_tcslen(szStore)) szStore=NULL;
	return pMAPI->OpenMessageStore(szStore);
}

CMAPIFolder* MAPIGetFolder(CMAPIEx* pMAPI)
{
	return pMAPI->GetFolder();
}

void MAPISetFolder(CMAPIEx* pMAPI, CMAPIFolder* pFolder)
{
	pMAPI->SetFolder(pFolder);
}

BOOL MAPIGetProfileName(CMAPIEx* pMAPI, LPTSTR szProfileName, int nMaxLength)
{
	CString strProfileName;
	if(pMAPI->GetProfileName(strProfileName)) 
	{
		CopyString(szProfileName, strProfileName, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL MAPIGetProfileEmail(CMAPIEx* pMAPI, LPTSTR szProfileEmail, int nMaxLength)
{
	CString strProfileEmail;
	if(pMAPI->GetProfileEmail(strProfileEmail)) 
	{
		CopyString(szProfileEmail, strProfileEmail, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

#ifdef _WIN32_WCE

CPOOM* MAPIGetPOOM(CMAPIEx* pMAPI)
{
	return pMAPI->GetPOOM();
}

BOOL POOMGetContents(CPOOM* pPOOM)
{
	return pPOOM->GetContents();
}

BOOL POOMSortContents(CPOOM* pPOOM, BOOL bDescending, LPCTSTR szSortField)
{
	return pPOOM->SortContents(bDescending, szSortField);
}

int POOMGetRowCount(CPOOM* pPOOM)
{
	return pPOOM->GetRowCount();
}

#endif

// Folders

CMAPIFolder* MAPIOpenFolder(CMAPIEx* pMAPI, LPCTSTR szFolderName, bool bInternal)
{
	return pMAPI->OpenFolder(szFolderName, bInternal);
}

CMAPIFolder* MAPIOpenRootFolder(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenRootFolder(bInternal);
}

CMAPIFolder* MAPIOpenInbox(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenInbox(bInternal);
}

CMAPIFolder* MAPIOpenOutbox(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenOutbox(bInternal);
}

CMAPIFolder* MAPIOpenSentItems(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenSentItems(bInternal);
}

CMAPIFolder* MAPIOpenDeletedItems(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenDeletedItems(bInternal);
}

CMAPIFolder* MAPIOpenContacts(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenContacts(bInternal);
}

CMAPIFolder* MAPIOpenDrafts(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenDrafts(bInternal);
}

CMAPIFolder* MAPIOpenCalendar(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenCalendar(bInternal);
}

CMAPIFolder* MAPIOpenJunkFolder(CMAPIEx* pMAPI, bool bInternal)
{
	return pMAPI->OpenJunkFolder(bInternal);
}

BOOL MAPIGetHierarchy(CMAPIEx* pMAPI)
{
	return (pMAPI->GetHierarchy()!=NULL);
}

// Messages, Contacts

BOOL MAPIGetContents(CMAPIEx* pMAPI)
{
	return (pMAPI->GetContents()!=NULL);
}

int MAPIGetRowCount(CMAPIEx* pMAPI)
{
	return pMAPI->GetRowCount();
}

enum SortFields { SORT_RECEIVED_TIME, SORT_SUBJECT };

BOOL MAPISortContents(CMAPIEx* pMAPI, BOOL bAscending, int nSortField)
{
	unsigned long ulSortField;
	switch(nSortField) 
	{
		default:
		case SORT_RECEIVED_TIME: ulSortField=PR_MESSAGE_DELIVERY_TIME;break;
		case SORT_SUBJECT: ulSortField=PR_SUBJECT;break;
	}

	return pMAPI->SortContents(bAscending ? TABLE_SORT_ASCEND : TABLE_SORT_DESCEND,ulSortField);
}

BOOL MAPISetUnreadOnly(CMAPIEx* pMAPI, BOOL bUnreadOnly)
{
	return pMAPI->SetUnreadOnly(bUnreadOnly);
}

BOOL MAPIGetNextMessage(CMAPIEx* pMAPI, CMAPIMessage*& pMessage)
{
	pMessage=new CMAPIMessage();
	if(!pMAPI->GetNextMessage(*pMessage)) 
	{
		delete pMessage;
		pMessage=NULL;
		return FALSE;
	}
	return TRUE;
}

BOOL MAPIGetNextContact(CMAPIEx* pMAPI, CMAPIContact*& pContact)
{
	pContact=new CMAPIContact();
	if(!pMAPI->GetNextContact(*pContact)) 
	{
		delete pContact;
		pContact=NULL;
		return FALSE;
	}
	return TRUE;
}

BOOL MAPIGetNextAppointment(CMAPIEx* pMAPI, CMAPIAppointment*& pAppointment)
{
	pAppointment=new CMAPIAppointment();
	if(!pMAPI->GetNextAppointment(*pAppointment)) 
	{
		delete pAppointment;
		pAppointment=NULL;
		return FALSE;
	}
	return TRUE;
}

BOOL MAPIGetNextSubFolder(CMAPIEx* pMAPI, CMAPIFolder*& pFolder, LPTSTR szFolder, int nMaxLength)
{
	CString strFolder;
	pFolder=new CMAPIFolder();
	if(!pMAPI->GetNextSubFolder(*pFolder, strFolder)) 
	{
		delete pFolder;
		pFolder=NULL;
		return FALSE;
	}
	CopyString(szFolder, strFolder, nMaxLength);
	return TRUE;
}

// Common Object functions

void ObjectClose(CMAPIObject* pObject)
{
	delete pObject;
}

BOOL ObjectSave(CMAPIObject* pObject, BOOL bClose)
{
	return pObject->Save(bClose);
}

int ObjectGetMessageFlags(CMAPIObject* pObject)
{
	return pObject->GetMessageFlags();
}

BOOL ObjectGetStringEntryID(CMAPIObject* pObject, LPTSTR szEntryID, int nMaxLength)
{
	CString strEntryID;
	if(pObject->GetEntryIDString(strEntryID)) 
	{
		CopyString(szEntryID, strEntryID, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ObjectGetMessageClass(CMAPIObject* pObject, LPTSTR szMessageClass, int nMaxLength)
{
	CString strMessageClass;
	if(pObject->GetMessageClass(strMessageClass)) 
	{
		CopyString(szMessageClass, strMessageClass, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

int ObjectGetMessageEditorFormat(CMAPIObject* pObject)
{
	return pObject->GetMessageEditorFormat();
}

BOOL ObjectGetPropertyString(CMAPIObject* pObject, ULONG ulProperty, LPTSTR szProperty, int nMaxLength, BOOL bStream)
{
	CString strProperty;
	if(pObject->GetPropertyString(ulProperty, strProperty, bStream)) 
	{
		CopyString(szProperty, strProperty, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ObjectGetNamedProperty(CMAPIObject* pObject, LPCTSTR szFieldName, LPTSTR szField, int nMaxLength)
{
	CString strField;
	if(pObject->GetNamedProperty(szFieldName, strField)) 
	{
		CopyString(szField, strField, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

CString ObjectBody;

BOOL ObjectGetBody(CMAPIObject* pObject, LPCTSTR& szBody, BOOL bAutoDetect)
{
	if(pObject->GetBody(ObjectBody, bAutoDetect))
	{
		szBody=ObjectBody;
		return TRUE;
	}
	return FALSE;
}

BOOL ObjectGetHTML(CMAPIObject* pObject, LPCTSTR& szHTML)
{
	if(pObject->GetHTML(ObjectBody))
	{
		szHTML=ObjectBody;
		return TRUE;
	}
	return FALSE;
}

BOOL ObjectGetRTF(CMAPIObject* pObject, LPCTSTR& szRTF)
{
	if(pObject->GetRTF(ObjectBody))
	{
		szRTF=ObjectBody;
		return TRUE;
	}
	return FALSE;
}

void ObjectFreeBody()
{
	ObjectBody.Empty();
}

BOOL ObjectSetMessageFlags(CMAPIObject* pObject, int nFlags)
{
	return pObject->SetMessageFlags(nFlags);
}

BOOL ObjectSetMessageEditorFormat(CMAPIObject* pObject, int nFormat)
{
	return pObject->SetMessageEditorFormat(nFormat);
}

BOOL ObjectSetPropertyString(CMAPIObject* pObject, ULONG ulProperty, LPCTSTR szProperty, BOOL bStream)
{
	return pObject->SetPropertyString(ulProperty, szProperty, bStream);
}

BOOL ObjectSetNamedProperty(CMAPIObject* pObject, LPCTSTR szFieldName, LPCTSTR szField, BOOL bCreate)
{
	return pObject->SetNamedProperty(szFieldName, szField, bCreate);
}

BOOL ObjectSetBody(CMAPIObject* pObject, LPCTSTR szBody)
{
	return pObject->SetBody(szBody);
}

BOOL ObjectSetHTML(CMAPIObject* pObject, LPCTSTR szHTML)
{
	return pObject->SetHTML(szHTML);
}

BOOL ObjectSetRTF(CMAPIObject* pObject, LPCTSTR szRTF)
{
	return pObject->SetRTF(szRTF);
}

// Folder functions

BOOL FolderGetHierarchy(CMAPIFolder* pFolder)
{
	return (pFolder->GetHierarchy()!=NULL);
}

CMAPIFolder* FolderOpenSubFolder(CMAPIFolder* pFolder, LPCTSTR szSubFolder)
{
	return pFolder->OpenSubFolder(szSubFolder);
}

CMAPIFolder* FolderCreateSubFolder(CMAPIFolder* pFolder, LPCTSTR szSubFolder)
{
	return pFolder->CreateSubFolder(szSubFolder);
}

BOOL FolderDeleteSubFolderByName(CMAPIFolder* pFolder, LPCTSTR szSubFolder)
{
	return pFolder->DeleteSubFolder(szSubFolder);
}

BOOL FolderDeleteSubFolder(CMAPIFolder* pFolder, CMAPIFolder* pSubFolder)
{
	return pFolder->DeleteSubFolder(pSubFolder);
}

BOOL FolderGetContents(CMAPIFolder* pFolder)
{
	return (pFolder->GetContents()!=NULL);
}

int FolderGetRowCount(CMAPIFolder* pFolder)
{
	return pFolder ? pFolder->GetRowCount() : 0;
}

BOOL FolderSortContents(CMAPIFolder* pFolder, BOOL bAscending, int nSortField)
{
	unsigned long ulSortField;
	switch(nSortField) 
	{
		default:
		case SORT_RECEIVED_TIME: ulSortField=PR_MESSAGE_DELIVERY_TIME;break;
		case SORT_SUBJECT: ulSortField=PR_SUBJECT;break;
	}

	return pFolder->SortContents(bAscending ? TABLE_SORT_ASCEND : TABLE_SORT_DESCEND,ulSortField);
}

BOOL FolderSetUnreadOnly(CMAPIFolder* pFolder, BOOL bUnreadOnly)
{
	return pFolder->SetUnreadOnly(bUnreadOnly);
}

BOOL FolderGetNextMessage(CMAPIFolder* pFolder, CMAPIMessage*& pMessage)
{
	pMessage=new CMAPIMessage();
	if(!pFolder->GetNextMessage(*pMessage)) 
	{
		delete pMessage;
		pMessage=NULL;
		return FALSE;
	}
	return TRUE;
}

BOOL FolderGetNextContact(CMAPIFolder* pFolder, CMAPIContact*& pContact)
{
	pContact=new CMAPIContact();
	if(!pFolder->GetNextContact(*pContact)) 
	{
		delete pContact;
		pContact=NULL;
		return FALSE;
	}
	return TRUE;
}

BOOL FolderGetNextAppointment(CMAPIFolder* pFolder, CMAPIAppointment*& pAppointment)
{
	pAppointment=new CMAPIAppointment();
	if(!pFolder->GetNextAppointment(*pAppointment)) 
	{
		delete pAppointment;
		pAppointment=NULL;
		return FALSE;
	}
	return TRUE;
}

BOOL FolderGetNextSubFolder(CMAPIFolder* pFolder, CMAPIFolder*& pSubFolder, LPTSTR szFolder, int nMaxLength)
{
	CString strFolder;
	pSubFolder=new CMAPIFolder();
	if(!pFolder->GetNextSubFolder(*pSubFolder, strFolder)) 
	{
		delete pSubFolder;
		pSubFolder=NULL;
		return FALSE;
	} 
	CopyString(szFolder, strFolder, nMaxLength);
	return TRUE;
}

BOOL FolderDeleteMessage(CMAPIFolder* pFolder, CMAPIMessage* pMessage)
{
	return pFolder->DeleteMessage(*pMessage);
}

BOOL FolderCopyMessage(CMAPIFolder* pFolder, CMAPIMessage* pMessage, CMAPIFolder* pFolderDest)
{
	return pFolder->CopyMessage(*pMessage, pFolderDest);
}

BOOL FolderMoveMessage(CMAPIFolder* pFolder, CMAPIMessage* pMessage, CMAPIFolder* pFolderDest)
{
	return pFolder->MoveMessage(*pMessage, pFolderDest);
}

BOOL FolderDeleteContact(CMAPIFolder* pFolder, CMAPIContact* pContact)
{
	return pFolder->DeleteContact(*pContact);
}

BOOL FolderDeleteAppointment(CMAPIFolder* pFolder, CMAPIAppointment* pAppointment)
{
	return pFolder->DeleteAppointment(*pAppointment);
}

// Message functions

BOOL MessageCreate(CMAPIEx* pMAPI, CMAPIMessage*& pMessage, int nImportance, BOOL bSaveToSentFolder, CMAPIFolder* pFolder)
{
	pMessage=new CMAPIMessage();
	if(!pMessage->Create(pMAPI, nImportance, bSaveToSentFolder, pFolder)) 
	{
		delete pMessage;
		pMessage=NULL;
		return FALSE;
	}
	return TRUE;
}

int MessageShowForm(CMAPIEx* pMAPI, CMAPIMessage* pMessage)
{
	return pMessage->ShowForm(pMAPI);
}

BOOL MessageSend(CMAPIMessage* pMessage)
{
	return pMessage->Send();
}

BOOL MessageIsUnread(CMAPIMessage* pMessage)
{
	return pMessage->IsUnread();
}

BOOL MessageMarkAsRead(CMAPIMessage* pMessage, BOOL bRead)
{
	return pMessage->MarkAsRead(bRead);
}

BOOL MessageGetHeader(CMAPIMessage* pMessage, LPTSTR szHeader, int nMaxLength)
{
	CString strHeader;
	if(pMessage->GetHeader(strHeader)) 
	{
		CopyString(szHeader, strHeader, nMaxLength);
		return true;
	}
	return false;
}

void MessageGetSenderName(CMAPIMessage* pMessage, LPTSTR szSenderName, int nMaxLength)
{
	CopyString(szSenderName, pMessage->GetSenderName(), nMaxLength);
}

void MessageGetSenderEmail(CMAPIMessage* pMessage, LPTSTR szSenderEmail, int nMaxLength)
{
	CopyString(szSenderEmail, pMessage->GetSenderEmail(), nMaxLength);
}

void MessageGetSubject(CMAPIMessage* pMessage, LPTSTR szSubject, int nMaxLength)
{
	CopyString(szSubject, pMessage->GetSubject(), nMaxLength);
}

BOOL MessageGetReceivedTime(CMAPIMessage* pMessage, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond)
{
	SYSTEMTIME tm;
	if(pMessage->GetReceivedTime(tm)) 
	{
		nYear=tm.wYear;
		nMonth=tm.wMonth;
		nDay=tm.wDay;
		nHour=tm.wHour;
		nMinute=tm.wMinute;
		nSecond=tm.wSecond;
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetReceivedTimeString(CMAPIMessage* pMessage, LPTSTR szReceivedTime, int nMaxLength, LPCTSTR szFormat)
{
	CString strReceivedTime;
	if(szFormat && !_tcslen(szFormat)) szFormat=NULL;
	if(pMessage->GetReceivedTime(strReceivedTime, szFormat)) 
	{
		CopyString(szReceivedTime, strReceivedTime, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetSubmitTime(CMAPIMessage* pMessage, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond)
{
	SYSTEMTIME tm;
	if(pMessage->GetSubmitTime(tm)) 
	{
		nYear=tm.wYear;
		nMonth=tm.wMonth;
		nDay=tm.wDay;
		nHour=tm.wHour;
		nMinute=tm.wMinute;
		nSecond=tm.wSecond;
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetSubmitTimeString(CMAPIMessage* pMessage, LPTSTR szSubmitTime, int nMaxLength, LPCTSTR szFormat)
{
	CString strSubmitTime;
	if(szFormat && !_tcslen(szFormat)) szFormat=NULL;
	if(pMessage->GetSubmitTime(strSubmitTime, szFormat)) 
	{
		CopyString(szSubmitTime, strSubmitTime, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetTo(CMAPIMessage* pMessage, LPTSTR szTo, int nMaxLength)
{
	CString strTo;
	if(pMessage->GetTo(strTo)) 
	{
		CopyString(szTo, strTo, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetCC(CMAPIMessage* pMessage, LPTSTR szCC, int nMaxLength)
{
	CString strCC;
	if(pMessage->GetCC(strCC)) 
	{
		CopyString(szCC, strCC, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetBCC(CMAPIMessage* pMessage, LPTSTR szBCC, int nMaxLength)
{
	CString strBCC;
	if(pMessage->GetBCC(strBCC)) 
	{
		CopyString(szBCC, strBCC, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

int MessageGetSensitivity(CMAPIMessage* pMessage)
{
	return pMessage->GetSensitivity();
}

int MessageGetPriority(CMAPIMessage* pMessage)
{
	return pMessage->GetPriority();
}

int MessageGetImportance(CMAPIMessage* pMessage)
{
	return pMessage->GetImportance();
}

BOOL MessageGetRecipients(CMAPIMessage* pMessage)
{
	return pMessage->GetRecipients();
}

BOOL MessageGetNextRecipient(CMAPIMessage* pMessage, LPTSTR szName, int nMaxLenName, LPTSTR szEmail, int nMaxLenEmail, int& nType)
{
	CString strName, strEmail;
	if(pMessage->GetNextRecipient(strName, strEmail, nType)) 
	{
		CopyString(szName, strName, nMaxLenName);
		CopyString(szEmail, strEmail, nMaxLenEmail);
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetReplyTo(CMAPIMessage* pMessage, LPTSTR szEmail, int nMaxLength)
{
	CString strEmail;
	if(pMessage->GetReplyTo(strEmail)) 
	{
		CopyString(szEmail, strEmail, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

int MessageGetAttachmentCount(CMAPIMessage* pMessage)
{
	return pMessage->GetAttachmentCount();
}

BOOL MessageGetAttachmentCID(CMAPIMessage* pMessage, LPTSTR szAttachmentCID, int nMaxLength, int nIndex)
{
	CString strAttachmentCID;
	if(pMessage->GetAttachmentCID(strAttachmentCID, nIndex)) 
	{
		CopyString(szAttachmentCID, strAttachmentCID, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL MessageGetAttachmentName(CMAPIMessage* pMessage, LPTSTR szAttachmentName, int nMaxLength, int nIndex)
{
	CString strAttachmentName;
	if(pMessage->GetAttachmentName(strAttachmentName, nIndex)) 
	{
		CopyString(szAttachmentName, strAttachmentName, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL MessageSaveAttachment(CMAPIMessage* pMessage, LPCTSTR szFolder, int nIndex)
{
	return pMessage->SaveAttachment(szFolder, nIndex);
}

BOOL MessageDeleteAttachment(CMAPIMessage* pMessage, int nIndex)
{
	return pMessage->DeleteAttachment(nIndex);
}

BOOL MessageSetMessageStatus(CMAPIMessage* pMessage, int nMessageStatus)
{
	return pMessage->SetMessageStatus(nMessageStatus);
}

BOOL MessageAddRecipient(CMAPIMessage* pMessage, LPCTSTR szEmail, int nType, LPCTSTR szAddrType)
{
	return pMessage->AddRecipient(szEmail, nType, szAddrType);
}

void MessageSetSubject(CMAPIMessage* pMessage, LPCTSTR szSubject)
{
	pMessage->SetSubject(szSubject);
}

void MessageSetSender(CMAPIMessage* pMessage, LPCTSTR szSenderName, LPCTSTR szSenderEmail)
{
	pMessage->SetSender(szSenderName, szSenderEmail);
}

BOOL MessageSetReceivedTime(CMAPIMessage* pMessage, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond, BOOL bLocal)
{
	SYSTEMTIME tm;
	CMAPIEx::GetSystemTime(tm, nYear, nMonth, nDay, nHour, nMinute, nSecond);
	return pMessage->SetReceivedTime(tm, bLocal);
}

BOOL MessageSetSubmitTime(CMAPIMessage* pMessage, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond, BOOL bLocal)
{
	SYSTEMTIME tm;
	CMAPIEx::GetSystemTime(tm, nYear, nMonth, nDay, nHour, nMinute, nSecond);
	return pMessage->SetSubmitTime(tm, bLocal);
}

BOOL MessageAddAttachment(CMAPIMessage* pMessage, LPCTSTR szPath, LPCTSTR szName, LPCTSTR szCID)
{
	if(szName && !_tcslen(szName)) szName=NULL;
	if(szCID && !_tcslen(szCID)) szCID=NULL;
	return pMessage->AddAttachment(szPath, szName, szCID);
}

BOOL MessageSetReadReceipt(CMAPIMessage* pMessage, BOOL bSet, LPCTSTR szReceiverEmail)
{
	if(szReceiverEmail && !_tcslen(szReceiverEmail)) szReceiverEmail=NULL;
	return pMessage->SetReadReceipt(bSet, szReceiverEmail);
}

BOOL MessageSetDeliveryReceipt(CMAPIMessage* pMessage, BOOL bSet)
{
	return pMessage->SetDeliveryReceipt(bSet);
}

BOOL MessageMarkAsPrivate(CMAPIMessage* pMessage)
{
	return pMessage->MarkAsPrivate();
}

BOOL MessageSetSensitivity(CMAPIMessage* pMessage, int nSensitivity)
{
	return pMessage->SetSensitivity(nSensitivity);
}

// Contact functions

BOOL ContactCreate(CMAPIEx* pMAPI, CMAPIContact*& pContact, CMAPIFolder* pFolder)
{
	pContact=new CMAPIContact();
	if(!pContact->Create(pMAPI, pFolder)) 
	{
		delete pContact;
		pContact=NULL;
		return FALSE;
	}
	return TRUE;
}

enum NameType { DISPLAY_NAME, GIVEN_NAME, MIDDLE_NAME, SURNAME };

ULONG GetNameID(int nNameType)
{
	ULONG ulNameID;
	switch(nNameType) 
	{
		default: ulNameID=0;break;
		case DISPLAY_NAME: ulNameID=PR_DISPLAY_NAME;break;
		case GIVEN_NAME: ulNameID=PR_GIVEN_NAME;break;
		case MIDDLE_NAME: ulNameID=PR_MIDDLE_NAME;break;
		case SURNAME: ulNameID=PR_SURNAME;break;
	}
	return ulNameID;
}

BOOL ContactGetName(CMAPIContact* pContact, LPTSTR szName, int nMaxLength, int nType)
{
	ULONG ulNameID=GetNameID(nType);
	if(!ulNameID) return FALSE;

	CString strName;
	if(pContact->GetName(strName,ulNameID)) 
	{
		CopyString(szName, strName, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetEmail(CMAPIContact* pContact, LPTSTR szEmail, int nMaxLength, int nIndex)
{
	CString strEmail;
	if(pContact->GetEmail(strEmail, nIndex)) 
	{
		CopyString(szEmail, strEmail, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetEmailDisplayAs(CMAPIContact* pContact, LPTSTR szDisplayAs, int nMaxLength, int nIndex)
{
	CString strEmail;
	if(pContact->GetEmailDisplayAs(strEmail, nIndex)) 
	{
		CopyString(szDisplayAs, strEmail, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetHomePage(CMAPIContact* pContact, LPTSTR szHomePage, int nMaxLength)
{
	CString strHomePage;
	if(pContact->GetHomePage(strHomePage)) 
	{
		CopyString(szHomePage, strHomePage, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

enum PhoneType {
	PRIMARY_TELEPHONE_NUMBER, BUSINESS_TELEPHONE_NUMBER, HOME_TELEPHONE_NUMBER, 
	CALLBACK_TELEPHONE_NUMBER, BUSINESS2_TELEPHONE_NUMBER, MOBILE_TELEPHONE_NUMBER,
	RADIO_TELEPHONE_NUMBER, CAR_TELEPHONE_NUMBER, OTHER_TELEPHONE_NUMBER,
	PAGER_TELEPHONE_NUMBER, PRIMARY_FAX_NUMBER, BUSINESS_FAX_NUMBER,
	HOME_FAX_NUMBER, TELEX_NUMBER, ISDN_NUMBER, ASSISTANT_TELEPHONE_NUMBER,
	HOME2_TELEPHONE_NUMBER, TTYTDD_PHONE_NUMBER, COMPANY_MAIN_PHONE_NUMBER
};

ULONG GetPhoneNumberID(int nPhoneType)
{
	ULONG ulPhoneNumberID;
	switch(nPhoneType) 
	{
		default: ulPhoneNumberID=0;break;
		case PRIMARY_TELEPHONE_NUMBER: ulPhoneNumberID=PR_PRIMARY_TELEPHONE_NUMBER;break;
		case BUSINESS_TELEPHONE_NUMBER: ulPhoneNumberID=PR_BUSINESS_TELEPHONE_NUMBER;break;
		case HOME_TELEPHONE_NUMBER: ulPhoneNumberID=PR_HOME_TELEPHONE_NUMBER;break;
		case CALLBACK_TELEPHONE_NUMBER: ulPhoneNumberID=PR_CALLBACK_TELEPHONE_NUMBER;break;
		case BUSINESS2_TELEPHONE_NUMBER: ulPhoneNumberID=PR_BUSINESS2_TELEPHONE_NUMBER;break;
		case MOBILE_TELEPHONE_NUMBER: ulPhoneNumberID=PR_MOBILE_TELEPHONE_NUMBER;break;
		case RADIO_TELEPHONE_NUMBER: ulPhoneNumberID=PR_RADIO_TELEPHONE_NUMBER;break;
		case CAR_TELEPHONE_NUMBER: ulPhoneNumberID=PR_CAR_TELEPHONE_NUMBER;break;
		case OTHER_TELEPHONE_NUMBER: ulPhoneNumberID=PR_OTHER_TELEPHONE_NUMBER;break;
		case PAGER_TELEPHONE_NUMBER: ulPhoneNumberID=PR_PAGER_TELEPHONE_NUMBER;break;
		case PRIMARY_FAX_NUMBER: ulPhoneNumberID=PR_PRIMARY_FAX_NUMBER;break;
		case BUSINESS_FAX_NUMBER: ulPhoneNumberID=PR_BUSINESS_FAX_NUMBER;break;
		case HOME_FAX_NUMBER: ulPhoneNumberID=PR_HOME_FAX_NUMBER;break;
		case TELEX_NUMBER: ulPhoneNumberID=PR_TELEX_NUMBER;break;
		case ISDN_NUMBER: ulPhoneNumberID=PR_ISDN_NUMBER;break;
		case ASSISTANT_TELEPHONE_NUMBER: ulPhoneNumberID=PR_ASSISTANT_TELEPHONE_NUMBER;break;
		case HOME2_TELEPHONE_NUMBER: ulPhoneNumberID=PR_HOME2_TELEPHONE_NUMBER;break;
		case TTYTDD_PHONE_NUMBER: ulPhoneNumberID=PR_TTYTDD_PHONE_NUMBER;break;
		case COMPANY_MAIN_PHONE_NUMBER: ulPhoneNumberID=PR_COMPANY_MAIN_PHONE_NUMBER;break;
	}
	return ulPhoneNumberID;
}

BOOL ContactGetPhoneNumber(CMAPIContact* pContact, LPTSTR szPhoneNumber, int nMaxLength, int nType)
{
	ULONG ulPhoneNumberID=GetPhoneNumberID(nType);
	if(!ulPhoneNumberID) return FALSE;

	CString	strPhoneNumber;
	if(pContact->GetPhoneNumber(strPhoneNumber,ulPhoneNumberID)) 
	{
		CopyString(szPhoneNumber, strPhoneNumber, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetAddress(CMAPIContact* pContact, CContactAddress*& pAddress, int nType)
{
	pAddress=new CContactAddress();
	if(!pContact->GetAddress(*pAddress, (CContactAddress::AddressType)nType)) 
	{
		delete pAddress;
		pAddress=NULL;
		return FALSE;
	}
	return TRUE;
}

BOOL ContactGetPostalAddress(CMAPIContact* pContact, LPTSTR szAddress, int nMaxLength)
{
	CString strAddress;
	if(pContact->GetPostalAddress(strAddress)) 
	{
		CopyString(szAddress, strAddress, nMaxLength);
		return TRUE;
	}
	return FALSE;	
}

int ContactGetNotesSize(CMAPIContact* pContact, BOOL bRTF)
{
	CString strNotes;
	if(pContact->GetNotes(strNotes, bRTF)) return strNotes.GetLength();
	return 0;	
}

BOOL ContactGetNotes(CMAPIContact* pContact, LPTSTR szNotes, int nMaxLength, BOOL bRTF)
{
	CString strNotes;
	if(pContact->GetNotes(strNotes, bRTF)) 
	{
		CopyString(szNotes, strNotes, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

int ContactGetSensitivity(CMAPIContact* pContact)
{
	return pContact->GetSensitivity();
}

BOOL ContactGetIMAddress(CMAPIContact* pContact, LPTSTR szIMAddress, int nMaxLength)
{
	CString strIMAddress;
	if(pContact->GetIMAddress(strIMAddress)) 
	{
		CopyString(szIMAddress, strIMAddress, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetTitle(CMAPIContact* pContact, LPTSTR szTitle, int nMaxLength)
{
	CString strTitle;
	if(pContact->GetTitle(strTitle)) 
	{
		CopyString(szTitle, strTitle, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetCompany(CMAPIContact* pContact, LPTSTR szCompany, int nMaxLength)
{
	CString strCompany;
	if(pContact->GetCompany(strCompany)) 
	{
		CopyString(szCompany, strCompany, nMaxLength);
		return TRUE;
	}
	return FALSE;
}


BOOL ContactGetProfession(CMAPIContact* pContact, LPTSTR szProfession, int nMaxLength)
{
	CString strProfession;
	if(pContact->GetProfession(strProfession)) 
	{
		CopyString(szProfession, strProfession, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetDisplayNamePrefix(CMAPIContact* pContact, LPTSTR szDisplayNamePrefix, int nMaxLength)
{
	CString strDisplayNamePrefix;
	if(pContact->GetDisplayNamePrefix(strDisplayNamePrefix)) 
	{
		CopyString(szDisplayNamePrefix, strDisplayNamePrefix, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetGeneration(CMAPIContact* pContact, LPTSTR szGeneration, int nMaxLength)
{
	CString strGeneration;
	if(pContact->GetGeneration(strGeneration)) 
	{
		CopyString(szGeneration, strGeneration, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetDepartment(CMAPIContact* pContact, LPTSTR szDepartment, int nMaxLength)
{
	CString strDepartment;
	if(pContact->GetDepartment(strDepartment)) 
	{
		CopyString(szDepartment, strDepartment, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetOffice(CMAPIContact* pContact, LPTSTR szOffice, int nMaxLength)
{
	CString strOffice;
	if(pContact->GetOffice(strOffice)) 
	{
		CopyString(szOffice, strOffice, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetManagerName(CMAPIContact* pContact, LPTSTR szManagerName, int nMaxLength)
{
	CString strManagerName;
	if(pContact->GetManagerName(strManagerName)) 
	{
		CopyString(szManagerName, strManagerName, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetAssistantName(CMAPIContact* pContact, LPTSTR szAssistantName, int nMaxLength)
{
	CString strAssistantName;
	if(pContact->GetAssistantName(strAssistantName)) 
	{
		CopyString(szAssistantName, strAssistantName, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetNickName(CMAPIContact* pContact, LPTSTR szNickName, int nMaxLength)
{
	CString strNickName;
	if(pContact->GetNickName(strNickName)) 
	{
		CopyString(szNickName, strNickName, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetSpouseName(CMAPIContact* pContact, LPTSTR szSpouseName, int nMaxLength)
{
	CString strSpouseName;
	if(pContact->GetSpouseName(strSpouseName)) 
	{
		CopyString(szSpouseName, strSpouseName, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetBirthday(CMAPIContact* pContact, int& nYear, int& nMonth, int& nDay)
{
	SYSTEMTIME tm;
	if(pContact->GetBirthday(tm)) 
	{
		nYear=tm.wYear;
		nMonth=tm.wMonth;
		nDay=tm.wDay;
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetBirthdayString(CMAPIContact* pContact, LPTSTR szBirthday, int nMaxLength, LPCTSTR szFormat)
{
	CString strBirthday;
	if(szFormat && !_tcslen(szFormat)) szFormat=NULL;
	if(pContact->GetBirthday(strBirthday, szFormat)) 
	{
		CopyString(szBirthday, strBirthday, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetAnniversary(CMAPIContact* pContact, int& nYear, int& nMonth, int& nDay)
{
	SYSTEMTIME tm;
	if(pContact->GetAnniversary(tm)) 
	{
		nYear=tm.wYear;
		nMonth=tm.wMonth;
		nDay=tm.wDay;
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetAnniversaryString(CMAPIContact* pContact, LPTSTR szAnniversary, int nMaxLength, LPCTSTR szFormat)
{
	CString strAnniversary;
	if(szFormat && !_tcslen(szFormat)) szFormat=NULL;
	if(pContact->GetAnniversary(strAnniversary, szFormat)) 
	{
		CopyString(szAnniversary, strAnniversary, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactGetCategories(CMAPIContact* pContact, LPTSTR szField, int nMaxLength)
{
	CString strField;
	if(pContact->GetCategories(strField)) 
	{
		CopyString(szField, strField, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL ContactSetName(CMAPIContact* pContact, LPCTSTR szName, int nType)
{
	ULONG ulNameID=GetNameID(nType);
	if(!ulNameID) return FALSE;
	return pContact->SetName(szName,ulNameID);
}

BOOL ContactSetEmail(CMAPIContact* pContact, LPCTSTR szEmail, int nIndex)
{
	return pContact->SetEmail(szEmail, nIndex);
}

BOOL ContactSetEmailDisplayAs(CMAPIContact* pContact, LPCTSTR szDisplayAs, int nIndex)
{
	return pContact->SetEmailDisplayAs(szDisplayAs, nIndex);
}

BOOL ContactSetHomePage(CMAPIContact* pContact, LPCTSTR szHomePage)
{
	return pContact->SetHomePage(szHomePage);
}

BOOL ContactSetPhoneNumber(CMAPIContact* pContact, LPCTSTR szPhoneNumber, int nType)
{
	ULONG ulPhoneNumberID=GetPhoneNumberID(nType);
	if(!ulPhoneNumberID) return FALSE;
	return pContact->SetPhoneNumber(szPhoneNumber,ulPhoneNumberID);
}

BOOL ContactSetAddress(CMAPIContact* pContact, CContactAddress* pAddress, CContactAddress::AddressType nType)
{
	return pContact->SetAddress(*pAddress, nType);
}

BOOL ContactSetPostalAddress(CMAPIContact* pContact, CContactAddress::AddressType nType)
{
	return pContact->SetPostalAddress(nType);
}

BOOL ContactUpdateDisplayAddress(CMAPIContact* pContact, CContactAddress::AddressType nType)
{
	return pContact->UpdateDisplayAddress(nType);
}

BOOL ContactSetNotes(CMAPIContact* pContact, LPCTSTR szNotes, BOOL bRTF)
{
	return pContact->SetNotes(szNotes, bRTF);
}

BOOL ContactSetSensitivity(CMAPIContact* pContact, int nSensitivity)
{
	return pContact->SetSensitivity(nSensitivity);
}

BOOL ContactSetIMAddress(CMAPIContact* pContact, LPCTSTR szIMAddress)
{
	return pContact->SetIMAddress(szIMAddress);
}

BOOL ContactSetFileAs(CMAPIContact* pContact, LPCTSTR szFileAs)
{
	return pContact->SetFileAs(szFileAs);
}

BOOL ContactSetTitle(CMAPIContact* pContact, LPCTSTR szTitle)
{
	return pContact->SetTitle(szTitle);
}

BOOL ContactSetCompany(CMAPIContact* pContact, LPCTSTR szCompany)
{
	return pContact->SetCompany(szCompany);
}

BOOL ContactSetProfession(CMAPIContact* pContact, LPCTSTR szProfession)
{
	return pContact->SetProfession(szProfession);
}

BOOL ContactSetDisplayNamePrefix(CMAPIContact* pContact, LPCTSTR szPrefix)
{
	return pContact->SetDisplayNamePrefix(szPrefix);
}

BOOL ContactSetGeneration(CMAPIContact* pContact, LPCTSTR szGeneration)
{
	return pContact->SetGeneration(szGeneration);
}

BOOL ContactUpdateDisplayName(CMAPIContact* pContact)
{
	return pContact->UpdateDisplayName();
}

BOOL ContactSetDepartment(CMAPIContact* pContact, LPCTSTR szDepartment)
{
	return pContact->SetDepartment(szDepartment);
}

BOOL ContactSetOffice(CMAPIContact* pContact, LPCTSTR szOffice)
{
	return pContact->SetOffice(szOffice);	
}

BOOL ContactSetManagerName(CMAPIContact* pContact, LPCTSTR szManagerName)
{
	return pContact->SetManagerName(szManagerName);
}

BOOL ContactSetAssistantName(CMAPIContact* pContact, LPCTSTR szAssistantName)
{
	return pContact->SetAssistantName(szAssistantName);
}

BOOL ContactSetNickName(CMAPIContact* pContact, LPCTSTR szNickName)
{
	return pContact->SetNickName(szNickName);
}

BOOL ContactSetSpouseName(CMAPIContact* pContact, LPCTSTR szSpouseName)
{
	return pContact->SetSpouseName(szSpouseName);
}

BOOL ContactSetBirthday(CMAPIContact* pContact, int nYear, int nMonth, int nDay)
{
	SYSTEMTIME tm;
	CMAPIEx::GetSystemTime(tm, nYear, nMonth, nDay);
	return pContact->SetBirthday(tm);
}

BOOL ContactSetAnniversary(CMAPIContact* pContact, int nYear, int nMonth, int nDay)
{
	SYSTEMTIME tm;
	CMAPIEx::GetSystemTime(tm, nYear, nMonth, nDay);
	return pContact->SetAnniversary(tm);
}

BOOL ContactSetCategories(CMAPIContact* pContact, LPCTSTR szCategories)
{
	return pContact->SetCategories(szCategories);
}

BOOL ContactSetPicture(CMAPIContact* pContact, LPCTSTR szPath)
{
	return pContact->SetPicture(szPath);
}

// Address functions 

void AddressClose(CContactAddress* pAddress)
{
	delete pAddress;
}

void AddressGetStreet(CContactAddress* pAddress, LPTSTR szStreet, int nMaxLength)
{
	CopyString(szStreet, pAddress->m_strStreet, nMaxLength);
}

void AddressGetCity(CContactAddress* pAddress, LPTSTR szCity, int nMaxLength)
{
	CopyString(szCity, pAddress->m_strCity, nMaxLength);
}

void AddressGetStateOrProvince(CContactAddress* pAddress, LPTSTR szStateOrProvince, int nMaxLength)
{
	CopyString(szStateOrProvince, pAddress->m_strStateOrProvince, nMaxLength);
}

void AddressGetPostalCode(CContactAddress* pAddress, LPTSTR szPostalCode, int nMaxLength)
{
	CopyString(szPostalCode, pAddress->m_strPostalCode, nMaxLength);
}

void AddressGetCountry(CContactAddress* pAddress, LPTSTR szCountry, int nMaxLength)
{
	CopyString(szCountry, pAddress->m_strCountry, nMaxLength);
}

void AddressSetStreet(CContactAddress* pAddress, LPCTSTR szStreet)
{
	pAddress->m_strStreet=szStreet;
}

void AddressSetCity(CContactAddress* pAddress, LPCTSTR szCity)
{
	pAddress->m_strCity=szCity;
}

void AddressSetStateOrProvince(CContactAddress* pAddress, LPCTSTR szStateOrProvince)
{
	pAddress->m_strStateOrProvince=szStateOrProvince;
}

void AddressSetPostalCode(CContactAddress* pAddress, LPCTSTR szPostalCode)
{
	pAddress->m_strPostalCode=szPostalCode;
}

void AddressSetCountry(CContactAddress* pAddress, LPCTSTR szCountry)
{
	pAddress->m_strCountry=szCountry;
}

// Appointment functions

BOOL AppointmentGetSubject(CMAPIAppointment* pAppointment, LPTSTR szSubject, int nMaxLength)
{
	CString strSubject;
	if(pAppointment->GetSubject(strSubject)) 
	{
		CopyString(szSubject, strSubject, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL AppointmentGetLocation(CMAPIAppointment* pAppointment, LPTSTR szLocation, int nMaxLength)
{
	CString strLocation;
	if(pAppointment->GetLocation(strLocation)) 
	{
		CopyString(szLocation, strLocation, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL AppointmentGetStartTime(CMAPIAppointment* pAppointment, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond)
{
	SYSTEMTIME tm;
	if(pAppointment->GetStartTime(tm)) 
	{
		nYear=tm.wYear;
		nMonth=tm.wMonth;
		nDay=tm.wDay;
		nHour=tm.wHour;
		nMinute=tm.wMinute;
		nSecond=tm.wSecond;
		return TRUE;
	}
	return FALSE;
}

BOOL AppointmentGetStartTimeString(CMAPIAppointment* pAppointment, LPTSTR szStartTime, int nMaxLength, LPCTSTR szFormat)
{
	CString strStartTime;
	if(szFormat && !_tcslen(szFormat)) szFormat=NULL;
	if(pAppointment->GetStartTime(strStartTime, szFormat)) 
	{
		CopyString(szStartTime, strStartTime, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL AppointmentGetEndTime(CMAPIAppointment* pAppointment, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond)
{
	SYSTEMTIME tm;
	if(pAppointment->GetEndTime(tm)) 
	{
		nYear=tm.wYear;
		nMonth=tm.wMonth;
		nDay=tm.wDay;
		nHour=tm.wHour;
		nMinute=tm.wMinute;
		nSecond=tm.wSecond;
		return TRUE;
	}
	return FALSE;
}

BOOL AppointmentGetEndTimeString(CMAPIAppointment* pAppointment, LPTSTR szEndTime, int nMaxLength, LPCTSTR szFormat)
{
	CString strEndTime;
	if(szFormat && !_tcslen(szFormat)) szFormat=NULL;
	if(pAppointment->GetEndTime(strEndTime, szFormat)) 
	{
		CopyString(szEndTime, strEndTime, nMaxLength);
		return TRUE;
	}
	return FALSE;
}

BOOL AppointmentSetSubject(CMAPIAppointment* pAppointment, LPCTSTR szSubject)
{
	return pAppointment->SetSubject(szSubject);
}

BOOL AppointmentSetLocation(CMAPIAppointment* pAppointment, LPCTSTR szLocation)
{
	return pAppointment->SetLocation(szLocation);
}

BOOL AppointmentSetStartTime(CMAPIAppointment* pAppointment, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond)
{
	SYSTEMTIME tm;
	CMAPIEx::GetSystemTime(tm, nYear, nMonth, nDay, nHour, nMinute, nSecond);
	return pAppointment->SetStartTime(tm);
}

BOOL AppointmentSetEndTime(CMAPIAppointment* pAppointment, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond)
{
	SYSTEMTIME tm;
	CMAPIEx::GetSystemTime(tm, nYear, nMonth, nDay, nHour, nMinute, nSecond);
	return pAppointment->SetEndTime(tm);
}
