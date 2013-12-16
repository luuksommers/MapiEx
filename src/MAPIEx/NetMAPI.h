#ifndef __NETMAPI_H__
#define __NETMAPI_H__

////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: NetMAPI.h
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

#include "MAPIEx.h"

extern "C" {

// Initialize and Terminate
AFX_EXT_CLASS BOOL MAPIInit(BOOL bMultiThreadedNotifications, BOOL bInitAsService);
AFX_EXT_CLASS void MAPITerm();

// Profiles, Message Store
AFX_EXT_CLASS CMAPIEx* MAPILogin(LPCTSTR szProfile, BOOL bInitAsService);
AFX_EXT_CLASS void MAPILogout(CMAPIEx* pMAPI);
AFX_EXT_CLASS BOOL MAPIOpenMessageStore(CMAPIEx* pMAPI, LPCTSTR szStore);
AFX_EXT_CLASS CMAPIFolder* MAPIGetFolder(CMAPIEx* pMAPI);
AFX_EXT_CLASS void MAPISetFolder(CMAPIEx* pMAPI, CMAPIFolder* pFolder);
AFX_EXT_CLASS BOOL MAPIGetProfileName(CMAPIEx* pMAPI, LPTSTR szProfileName, int nMaxLength);
AFX_EXT_CLASS BOOL MAPIGetProfileEmail(CMAPIEx* pMAPI, LPTSTR szProfileEmail, int nMaxLength);

// POOM related functions
#ifdef _WIN32_WCE
AFX_EXT_CLASS CPOOM* MAPIGetPOOM(CMAPIEx* pMAPI);
AFX_EXT_CLASS BOOL POOMGetContents(CPOOM* pPOOM);
AFX_EXT_CLASS BOOL POOMSortContents(CPOOM* pPOOM, BOOL bDescending, LPCTSTR szSortField);
AFX_EXT_CLASS int POOMGetRowCount(CPOOM* pPOOM);
#endif

// Folders
AFX_EXT_CLASS CMAPIFolder* MAPIOpenFolder(CMAPIEx* pMAPI, LPCTSTR szFolderName, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenRootFolder(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenInbox(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenOutbox(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenSentItems(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenDeletedItems(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenContacts(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenDrafts(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenCalendar(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS CMAPIFolder* MAPIOpenJunkFolder(CMAPIEx* pMAPI, bool bInternal);
AFX_EXT_CLASS BOOL MAPIGetHierarchy(CMAPIEx* pMAPI);

// Messages, Contacts
AFX_EXT_CLASS BOOL MAPIGetContents(CMAPIEx* pMAPI);
AFX_EXT_CLASS int MAPIGetRowCount(CMAPIEx* pMAPI);
AFX_EXT_CLASS BOOL MAPISortContents(CMAPIEx* pMAPI, BOOL bAscending, int nSortField);
AFX_EXT_CLASS BOOL MAPISetUnreadOnly(CMAPIEx* pMAPI, BOOL bUnreadOnly);
AFX_EXT_CLASS BOOL MAPIGetNextMessage(CMAPIEx* pMAPI, CMAPIMessage*& pMessage);
AFX_EXT_CLASS BOOL MAPIGetNextContact(CMAPIEx* pMAPI, CMAPIContact*& pContact);
AFX_EXT_CLASS BOOL MAPIGetNextAppointment(CMAPIEx* pMAPI, CMAPIAppointment*& pAppointment);
AFX_EXT_CLASS BOOL MAPIGetNextSubFolder(CMAPIEx* pMAPI, CMAPIFolder*& pFolder, LPTSTR szFolder, int nMaxLength);

// Common Object functions
AFX_EXT_CLASS void ObjectClose(CMAPIObject* pObject);
AFX_EXT_CLASS BOOL ObjectSave(CMAPIObject* pObject, BOOL bClose);
AFX_EXT_CLASS BOOL ObjectGetEntryID(CMAPIObject* pObject, LPTSTR szEntryID, int nMaxLength);
AFX_EXT_CLASS int ObjectGetMessageFlags(CMAPIObject* pObject);
AFX_EXT_CLASS BOOL ObjectGetMessageClass(CMAPIObject* pObject, LPTSTR szMessageClass, int nMaxLength);
AFX_EXT_CLASS int ObjectGetMessageEditorFormat(CMAPIObject* pObject);
AFX_EXT_CLASS BOOL ObjectGetPropertyString(CMAPIObject* pObject, ULONG ulProperty, LPTSTR szProperty, int nMaxLength, BOOL bStream);
AFX_EXT_CLASS BOOL ObjectGetNamedProperty(CMAPIObject* pObject, LPCTSTR szFieldName, LPTSTR szField, int nMaxLength);
AFX_EXT_CLASS BOOL ObjectGetBody(CMAPIObject* pObject, LPCTSTR& szBody, BOOL bAutoDetect);
AFX_EXT_CLASS BOOL ObjectGetHTML(CMAPIObject* pObject, LPCTSTR& szHTML);
AFX_EXT_CLASS BOOL ObjectGetRTF(CMAPIObject* pObject, LPCTSTR& szRTF);
AFX_EXT_CLASS void ObjectFreeBody();
AFX_EXT_CLASS BOOL ObjectSetMessageFlags(CMAPIObject* pObject, int nFlags);
AFX_EXT_CLASS BOOL ObjectSetMessageEditorFormat(CMAPIObject* pObject, int nFormat);
AFX_EXT_CLASS BOOL ObjectSetPropertyString(CMAPIObject* pObject, ULONG ulProperty, LPCTSTR szProperty, BOOL bStream);
AFX_EXT_CLASS BOOL ObjectSetNamedProperty(CMAPIObject* pObject, LPCTSTR szFieldName, LPCTSTR szField, BOOL bCreate);
AFX_EXT_CLASS BOOL ObjectSetBody(CMAPIObject* pObject, LPCTSTR szBody);
AFX_EXT_CLASS BOOL ObjectSetHTML(CMAPIObject* pObject, LPCTSTR szHTML);
AFX_EXT_CLASS BOOL ObjectSetRTF(CMAPIObject* pObject, LPCTSTR szRTF);

// Folder functions
AFX_EXT_CLASS BOOL FolderGetHierarchy(CMAPIFolder* pFolder);
AFX_EXT_CLASS CMAPIFolder* FolderOpenSubFolder(CMAPIFolder* pFolder, LPCTSTR szSubFolder);
AFX_EXT_CLASS CMAPIFolder* FolderCreateSubFolder(CMAPIFolder* pFolder, LPCTSTR szSubFolder);
AFX_EXT_CLASS BOOL FolderDeleteSubFolderByName(CMAPIFolder* pFolder, LPCTSTR szSubFolder);
AFX_EXT_CLASS BOOL FolderDeleteSubFolder(CMAPIFolder* pFolder, CMAPIFolder* pSubFolder);
AFX_EXT_CLASS BOOL FolderGetContents(CMAPIFolder* pFolder);
AFX_EXT_CLASS int FolderGetRowCount(CMAPIFolder* pFolder);
AFX_EXT_CLASS BOOL FolderSortContents(CMAPIFolder* pFolder, BOOL bAscending, int nSortField);
AFX_EXT_CLASS BOOL FolderSetUnreadOnly(CMAPIFolder* pFolder, BOOL bUnreadOnly);
AFX_EXT_CLASS BOOL FolderGetNextMessage(CMAPIFolder* pFolder, CMAPIMessage*& pMessage);
AFX_EXT_CLASS BOOL FolderGetNextContact(CMAPIFolder* pFolder, CMAPIContact*& pContact);
AFX_EXT_CLASS BOOL FolderGetNextAppointment(CMAPIFolder* pFolder, CMAPIAppointment*& pAppointment);
AFX_EXT_CLASS BOOL FolderGetNextSubFolder(CMAPIFolder* pFolder, CMAPIFolder*& pSubFolder, LPTSTR szFolder, int nMaxLength);
AFX_EXT_CLASS BOOL FolderDeleteMessage(CMAPIFolder* pFolder, CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL FolderCopyMessage(CMAPIFolder* pFolder, CMAPIMessage* pMessage, CMAPIFolder* pFolderDest);
AFX_EXT_CLASS BOOL FolderMoveMessage(CMAPIFolder* pFolder, CMAPIMessage* pMessage, CMAPIFolder* pFolderDest);
AFX_EXT_CLASS BOOL FolderDeleteContact(CMAPIFolder* pFolder, CMAPIContact* pContact);
AFX_EXT_CLASS BOOL FolderDeleteAppointment(CMAPIFolder* pFolder, CMAPIAppointment* pAppointment);

// Message functions
AFX_EXT_CLASS BOOL MessageCreate(CMAPIEx* pMAPI, CMAPIMessage*& pMessage, int nImportance, BOOL bSaveToSentFolder, CMAPIFolder* pFolder);
AFX_EXT_CLASS int MessageShowForm(CMAPIEx* pMAPI, CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL MessageSend(CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL MessageIsUnread(CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL MessageMarkAsRead(CMAPIMessage* pMessage, BOOL bRead);
AFX_EXT_CLASS BOOL MessageGetHeader(CMAPIMessage* pMessage, LPTSTR szHeader, int nMaxLength);
AFX_EXT_CLASS void MessageGetSenderName(CMAPIMessage* pMessage, LPTSTR szSenderName, int nMaxLength);
AFX_EXT_CLASS void MessageGetSenderEmail(CMAPIMessage* pMessage, LPTSTR szSenderEmail, int nMaxLength);
AFX_EXT_CLASS void MessageGetSubject(CMAPIMessage* pMessage, LPTSTR szSubject, int nMaxLength);
AFX_EXT_CLASS BOOL MessageGetReceivedTime(CMAPIMessage* pMessage, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond);
AFX_EXT_CLASS BOOL MessageGetReceivedTimeString(CMAPIMessage* pMessage, LPTSTR szReceivedTime, int nMaxLength, LPCTSTR szFormat);
AFX_EXT_CLASS BOOL MessageGetSubmitTime(CMAPIMessage* pMessage, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond);
AFX_EXT_CLASS BOOL MessageGetSubmitTimeString(CMAPIMessage* pMessage, LPTSTR szSubmitTime, int nMaxLength, LPCTSTR szFormat);
AFX_EXT_CLASS BOOL MessageGetTo(CMAPIMessage* pMessage, LPTSTR szTo, int nMaxLength);
AFX_EXT_CLASS BOOL MessageGetCC(CMAPIMessage* pMessage, LPTSTR szCC, int nMaxLength);
AFX_EXT_CLASS BOOL MessageGetBCC(CMAPIMessage* pMessage, LPTSTR szBCC, int nMaxLength);
AFX_EXT_CLASS int MessageGetSensitivity(CMAPIMessage* pMessage);
AFX_EXT_CLASS int MessageGetPriority(CMAPIMessage* pMessage);
AFX_EXT_CLASS int MessageGetImportance(CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL MessageGetRecipients(CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL MessageGetNextRecipient(CMAPIMessage* pMessage, LPTSTR szName, int nMaxLenName, LPTSTR szEmail, int nMaxLenEmail, int& nType);
AFX_EXT_CLASS BOOL MessageGetReplyTo(CMAPIMessage* pMessage, LPTSTR szEmail, int nMaxLength);
AFX_EXT_CLASS int MessageGetAttachmentCount(CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL MessageGetAttachmentName(CMAPIMessage* pMessage, LPTSTR szAttachmentName, int nMaxLength, int nIndex);
AFX_EXT_CLASS BOOL MessageSaveAttachment(CMAPIMessage* pMessage, LPCTSTR szFolder, int nIndex);
AFX_EXT_CLASS BOOL MessageDeleteAttachment(CMAPIMessage* pMessage, int nIndex);
AFX_EXT_CLASS BOOL MessageSetMessageStatus(CMAPIMessage* pMessage, int nMessageStatus);
AFX_EXT_CLASS BOOL MessageAddRecipient(CMAPIMessage* pMessage, LPCTSTR szEmail, int nType, LPCTSTR szAddrType);
AFX_EXT_CLASS void MessageSetSubject(CMAPIMessage* pMessage, LPCTSTR szSubject);
AFX_EXT_CLASS void MessageSetSender(CMAPIMessage* pMessage, LPCTSTR szSenderName, LPCTSTR szSenderEmail);
AFX_EXT_CLASS BOOL MessageSetReceivedTime(CMAPIMessage* pMessage, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond, BOOL bLocal);
AFX_EXT_CLASS BOOL MessageSetSubmitTime(CMAPIMessage* pMessage, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond, BOOL bLocal);
AFX_EXT_CLASS BOOL MessageAddAttachment(CMAPIMessage* pMessage, LPCTSTR szPath, LPCTSTR szName, LPCTSTR szCID);
AFX_EXT_CLASS BOOL MessageSetReadReceipt(CMAPIMessage* pMessage, BOOL bSet, LPCTSTR szReceiverEmail);
AFX_EXT_CLASS BOOL MessageSetDeliveryReceipt(CMAPIMessage* pMessage, BOOL bSet);
AFX_EXT_CLASS BOOL MessageMarkAsPrivate(CMAPIMessage* pMessage);
AFX_EXT_CLASS BOOL MessageSetSensitivity(CMAPIMessage* pMessage, int nSensitivity);

// Contact functions
AFX_EXT_CLASS BOOL ContactCreate(CMAPIEx* pMAPI, CMAPIContact*& pContact, CMAPIFolder* pFolder);
AFX_EXT_CLASS BOOL ContactGetName(CMAPIContact* pContact, LPTSTR szName, int nMaxLength, int nType);
AFX_EXT_CLASS BOOL ContactGetEmail(CMAPIContact* pContact, LPTSTR szEmail, int nMaxLength, int nIndex);
AFX_EXT_CLASS BOOL ContactGetEmailDisplayAs(CMAPIContact* pContact, LPTSTR szDisplayAs, int nMaxLength, int nIndex);
AFX_EXT_CLASS BOOL ContactGetHomePage(CMAPIContact* pContact, LPTSTR szHomePage, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetPhoneNumber(CMAPIContact* pContact, LPTSTR szPhoneNumber, int nMaxLength, int nType);
AFX_EXT_CLASS BOOL ContactGetAddress(CMAPIContact* pContact, CContactAddress*& pAddress, int nType);
AFX_EXT_CLASS BOOL ContactGetPostalAddress(CMAPIContact* pContact, LPTSTR szAddress, int nMaxLength);
AFX_EXT_CLASS int ContactGetNotesSize(CMAPIContact* pContact, BOOL bRTF);
AFX_EXT_CLASS BOOL ContactGetNotes(CMAPIContact* pContact, LPTSTR szNotes, int nMaxLength, BOOL bRTF);
AFX_EXT_CLASS int ContactGetSensitivity(CMAPIContact* pContact);
AFX_EXT_CLASS BOOL ContactGetIMAddress(CMAPIContact* pContact, LPTSTR szIMAddress, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetTitle(CMAPIContact* pContact, LPTSTR szTitle, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetCompany(CMAPIContact* pContact, LPTSTR szCompany, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetProfession(CMAPIContact* pContact, LPTSTR szProfession, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetDisplayNamePrefix(CMAPIContact* pContact, LPTSTR szDisplayNamePrefix, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetGeneration(CMAPIContact* pContact, LPTSTR szGeneration, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetDepartment(CMAPIContact* pContact, LPTSTR szDepartment, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetOffice(CMAPIContact* pContact, LPTSTR szOffice, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetManagerName(CMAPIContact* pContact, LPTSTR szManagerName, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetAssistantName(CMAPIContact* pContact, LPTSTR szAssistantName, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetNickName(CMAPIContact* pContact, LPTSTR szNickName, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetSpouseName(CMAPIContact* pContact, LPTSTR szSpouseName, int nMaxLength);
AFX_EXT_CLASS BOOL ContactGetBirthday(CMAPIContact* pContact, int& nYear, int& nMonth, int& nDay);
AFX_EXT_CLASS BOOL ContactGetBirthdayString(CMAPIContact* pContact, LPTSTR szBirthday, int nMaxLength, LPCTSTR szFormat);
AFX_EXT_CLASS BOOL ContactGetAnniversary(CMAPIContact* pContact, int& nYear, int& nMonth, int& nDay);
AFX_EXT_CLASS BOOL ContactGetAnniversaryString(CMAPIContact* pContact, LPTSTR szAnniversary, int nMaxLength, LPCTSTR szFormat);
AFX_EXT_CLASS BOOL ContactGetCategories(CMAPIContact* pContact, LPTSTR szField, int nMaxLength);
AFX_EXT_CLASS BOOL ContactSetName(CMAPIContact* pContact, LPCTSTR szName, int nType);
AFX_EXT_CLASS BOOL ContactSetEmail(CMAPIContact* pContact, LPCTSTR szEmail, int nIndex);
AFX_EXT_CLASS BOOL ContactSetEmailDisplayAs(CMAPIContact* pContact, LPCTSTR szDisplayAs, int nIndex);
AFX_EXT_CLASS BOOL ContactSetHomePage(CMAPIContact* pContact, LPCTSTR szHomePage);
AFX_EXT_CLASS BOOL ContactSetPhoneNumber(CMAPIContact* pContact, LPCTSTR szPhoneNumber, int nType);
AFX_EXT_CLASS BOOL ContactSetAddress(CMAPIContact* pContact, CContactAddress* pAddress, CContactAddress::AddressType nType);
AFX_EXT_CLASS BOOL ContactSetPostalAddress(CMAPIContact* pContact, CContactAddress::AddressType nType);
AFX_EXT_CLASS BOOL ContactUpdateDisplayAddress(CMAPIContact* pContact, CContactAddress::AddressType nType);
AFX_EXT_CLASS BOOL ContactSetNotes(CMAPIContact* pContact, LPCTSTR szNotes, BOOL bRTF);
AFX_EXT_CLASS BOOL ContactSetSensitivity(CMAPIContact* pContact, int nSensitivity);
AFX_EXT_CLASS BOOL ContactSetIMAddress(CMAPIContact* pContact, LPTSTR szIMAddress);
AFX_EXT_CLASS BOOL ContactSetFileAs(CMAPIContact* pContact, LPCTSTR szFileAs);
AFX_EXT_CLASS BOOL ContactSetTitle(CMAPIContact* pContact, LPCTSTR szTitle);
AFX_EXT_CLASS BOOL ContactSetCompany(CMAPIContact* pContact, LPCTSTR szCompany);
AFX_EXT_CLASS BOOL ContactSetProfession(CMAPIContact* pContact, LPTSTR szProfession);
AFX_EXT_CLASS BOOL ContactSetDisplayNamePrefix(CMAPIContact* pContact, LPCTSTR szPrefix);
AFX_EXT_CLASS BOOL ContactSetGeneration(CMAPIContact* pContact, LPCTSTR szGeneration);
AFX_EXT_CLASS BOOL ContactUpdateDisplayName(CMAPIContact* pContact);
AFX_EXT_CLASS BOOL ContactSetDepartment(CMAPIContact* pContact, LPCTSTR szDepartment);
AFX_EXT_CLASS BOOL ContactSetOffice(CMAPIContact* pContact, LPCTSTR szOffice);
AFX_EXT_CLASS BOOL ContactSetManagerName(CMAPIContact* pContact, LPCTSTR szManagerName);
AFX_EXT_CLASS BOOL ContactSetAssistantName(CMAPIContact* pContact, LPCTSTR szAssistantName);
AFX_EXT_CLASS BOOL ContactSetNickName(CMAPIContact* pContact, LPCTSTR szNickName);
AFX_EXT_CLASS BOOL ContactSetSpouseName(CMAPIContact* pContact, LPCTSTR szSpouseName);
AFX_EXT_CLASS BOOL ContactSetBirthday(CMAPIContact* pContact, int nYear, int nMonth, int nDay);
AFX_EXT_CLASS BOOL ContactSetAnniversary(CMAPIContact* pContact, int nYear, int nMonth, int nDay);
AFX_EXT_CLASS BOOL ContactSetCategories(CMAPIContact* pContact, LPCTSTR szCategories);
AFX_EXT_CLASS BOOL ContactSetPicture(CMAPIContact* pContact, LPCTSTR szPath);

// Address functions 
AFX_EXT_CLASS void AddressClose(CContactAddress* pAddress);
AFX_EXT_CLASS void AddressGetStreet(CContactAddress* pAddress, LPTSTR szStreet, int nMaxLength);
AFX_EXT_CLASS void AddressGetCity(CContactAddress* pAddress, LPTSTR szCity, int nMaxLength);
AFX_EXT_CLASS void AddressGetStateOrProvince(CContactAddress* pAddress, LPTSTR szStateOrProvince, int nMaxLength);
AFX_EXT_CLASS void AddressGetPostalCode(CContactAddress* pAddress, LPTSTR szPostalCode, int nMaxLength);
AFX_EXT_CLASS void AddressGetCountry(CContactAddress* pAddress, LPTSTR szCountry, int nMaxLength);
AFX_EXT_CLASS void AddressSetStreet(CContactAddress* pAddress, LPCTSTR szStreet);
AFX_EXT_CLASS void AddressSetCity(CContactAddress* pAddress, LPCTSTR szCity);
AFX_EXT_CLASS void AddressSetStateOrProvince(CContactAddress* pAddress, LPCTSTR szStateOrProvince);
AFX_EXT_CLASS void AddressSetPostalCode(CContactAddress* pAddress, LPCTSTR szPostalCode);
AFX_EXT_CLASS void AddressSetCountry(CContactAddress* pAddress, LPCTSTR szCountry);

// Appointment functions 
AFX_EXT_CLASS BOOL AppointmentGetSubject(CMAPIAppointment* pAppointment, LPTSTR szSubject, int nMaxLength);
AFX_EXT_CLASS BOOL AppointmentGetLocation(CMAPIAppointment* pAppointment, LPTSTR szLocation, int nMaxLength);
AFX_EXT_CLASS BOOL AppointmentGetStartTime(CMAPIAppointment* pAppointment, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond);
AFX_EXT_CLASS BOOL AppointmentGetStartTimeString(CMAPIAppointment* pAppointment, LPTSTR szStartTime, int nMaxLength, LPCTSTR szFormat);
AFX_EXT_CLASS BOOL AppointmentGetEndTime(CMAPIAppointment* pAppointment, int& nYear, int& nMonth, int& nDay, int& nHour, int& nMinute, int& nSecond);
AFX_EXT_CLASS BOOL AppointmentGetEndTimeString(CMAPIAppointment* pAppointment, LPTSTR szEndTime, int nMaxLength, LPCTSTR szFormat);
AFX_EXT_CLASS BOOL AppointmentSetSubject(CMAPIAppointment* pAppointment, LPCTSTR szSubject);
AFX_EXT_CLASS BOOL AppointmentSetLocation(CMAPIAppointment* pAppointment, LPCTSTR szLocation);
AFX_EXT_CLASS BOOL AppointmentSetStartTime(CMAPIAppointment* pAppointment, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond);
AFX_EXT_CLASS BOOL AppointmentSetEndTime(CMAPIAppointment* pAppointment, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond);

}

#endif
