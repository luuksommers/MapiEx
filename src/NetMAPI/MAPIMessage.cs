////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIMessage.cs
// Description: .NET Extended MAPI wrapper for Messages
//
// Copyright (C) 2005-2010, Noel Dillabough
//
// This source code is free to use and modify provided this notice remains intact and that any enhancements
// or bug fixes are posted to the CodeProject page hosting this class for all to benefit.
//
// Usage: see the CodeProject article at http://www.codeproject.com
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Runtime.InteropServices;
using System.Text;

namespace MAPIEx
{
    /// <summary>
    /// Messages
    /// </summary>
    public class MAPIMessage : MAPIObject
    {
        public enum RecipientType { UNKNOWN, TO, CC, BCC };
        public enum Importance { IMPORTANCE_LOW, IMPORTANCE_NORMAL, IMPORTANCE_HIGH };

        public MAPIMessage()
        {
        }

        public MAPIMessage(IntPtr pMessage) : base(pMessage)
        {
            
        }

        #region Message Functions

        /// <summary>
        /// Create a new message in the current folder
        /// </summary>
        /// <param name="mapi">NetMAPI session with an open folder</param>
        /// <param name="nImportance">priority of the message</param>
        /// <returns>true on success</returns>
        public bool Create(NetMAPI mapi, Importance nImportance)
        {
            return MessageCreate(mapi.MAPI, out pObject, (int)nImportance, true, IntPtr.Zero);
        }

        /// <summary>
        /// Create a new message in the current folder
        /// </summary>
        /// <param name="mapi">NetMAPI session with an open folder</param>
        /// <param name="nImportance">priority of the message</param>
        /// <param name="bSaveToSentFolder">save in sent or delete after sending</param>
        /// <returns>true on success</returns>
        public bool Create(NetMAPI mapi, Importance nImportance, bool bSaveToSentFolder)
        {
            return MessageCreate(mapi.MAPI, out pObject, (int)nImportance, bSaveToSentFolder, IntPtr.Zero);
        }

        /// <summary>
        /// Create a new message in the current folder
        /// </summary>
        /// <param name="mapi">NetMAPI session with an open folder</param>
        /// <param name="nImportance">priority of the message</param>
        /// <param name="bSaveToSentFolder">save in sent or delete after sending</param>
        /// <param name="pFolder">folder to create in</param>
        /// <returns>true on success</returns>
        public bool Create(NetMAPI mapi, Importance nImportance, bool bSaveToSentFolder, IntPtr pFolder)
        {
            return MessageCreate(mapi.MAPI, out pObject, (int)nImportance, bSaveToSentFolder, pFolder);
        }

        /// <summary>
        /// Shows the default IMessage form for this message
        /// </summary>
        /// <param name="mapi">NetMAPI session with an open folder</param>
        /// <returns>0 on failure, IDOK (1) on close or send and IDCANCEL (2) on close for new messages</returns>
        public int ShowForm(NetMAPI mapi)
        {
            return MessageShowForm(mapi.MAPI, pObject);
        }

        /// <summary>
        /// Sends the message
        /// </summary>
        /// <returns>true on success</returns>
        public bool Send()
        {
            return MessageSend(pObject);
        }

        /// <summary>
        /// Is this an unread message?
        /// </summary>
        /// <returns>true if unread</returns>
        public bool IsUnread()
        {
            return MessageIsUnread(pObject);
        }

        /// <summary>
        /// Mark this message as read or unread
        /// </summary>
        /// <param name="bRead">true for read</param>
        public bool MarkAsRead(bool bRead)
        {
            return MessageMarkAsRead(pObject, bRead);
        }

        /// <summary>
        /// Get the message header
        /// </summary>
        /// <param name="strHeader">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetHeader(StringBuilder strHeader)
        {
            return MessageGetHeader(pObject, strHeader, strHeader.Capacity);
        }

        /// <summary>
        /// Get the sender name
        /// </summary>
        /// <param name="strSenderName">buffer to receive</param>
        public void GetSenderName(StringBuilder strSenderName)
        {
            MessageGetSenderName(pObject, strSenderName, strSenderName.Capacity);
        }

        /// <summary>
        /// Get the sender's email address
        /// </summary>
        /// <param name="strSenderEmail">buffer to receive</param>
        public void GetSenderEmail(StringBuilder strSenderEmail)
        {
            MessageGetSenderEmail(pObject, strSenderEmail, strSenderEmail.Capacity);
        }

        /// <summary>
        /// Get the message subject
        /// </summary>
        /// <param name="strSubject">buffer to receive</param>
        public void GetSubject(StringBuilder strSubject)
        {
            MessageGetSubject(pObject, strSubject, strSubject.Capacity);
        }

        /// <summary>
        /// Gets the received time
        /// </summary>
        /// <param name="dt">DateTime received</param>
        /// <returns>true on success</returns>
        public bool GetReceivedTime(out DateTime dt)
        {
            int nYear, nMonth, nDay, nHour, nMinute, nSecond;
            bool bResult = MessageGetReceivedTime(pObject, out nYear, out nMonth, out nDay, out nHour, out nMinute, out nSecond);
            dt = new DateTime(nYear, nMonth, nDay, nHour, nMinute, nSecond);
            return bResult;
        }

        /// <summary>
        /// Gets the received time using the default format (MM/dd/yyyy hh:mm:ss tt)
        /// </summary>
        /// <param name="strReceivedTime">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetReceivedTime(StringBuilder strReceivedTime)
        {
            return MessageGetReceivedTimeString(pObject, strReceivedTime, strReceivedTime.Capacity, "");
        }

        /// <summary>
        /// Gets the received time
        /// </summary>
        /// <param name="strReceivedTime">buffer to receive</param>
        /// <param name="strFormat">format string for date (empty for default)</param>
        /// <returns>true on success</returns>
        public bool GetReceivedTime(StringBuilder strReceivedTime, string strFormat)
        {
            return MessageGetReceivedTimeString(pObject, strReceivedTime, strReceivedTime.Capacity, strFormat);
        }

        /// <summary>
        /// Gets the submit time
        /// </summary>
        /// <param name="dt">DateTime submitted</param>
        /// <returns>true on success</returns>
        public bool GetSubmitTime(out DateTime dt)
        {
            int nYear, nMonth, nDay, nHour, nMinute, nSecond;
            bool bResult = MessageGetSubmitTime(pObject, out nYear, out nMonth, out nDay, out nHour, out nMinute, out nSecond);
            dt = new DateTime(nYear, nMonth, nDay, nHour, nMinute, nSecond);
            return bResult;
        }

        /// <summary>
        /// Gets the submit time using the default format (MM/dd/yyyy hh:mm:ss tt)
        /// </summary>
        /// <param name="strSubmitTime">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetSubmitTime(StringBuilder strSubmitTime)
        {
            return MessageGetSubmitTimeString(pObject, strSubmitTime, strSubmitTime.Capacity, "");
        }

        /// <summary>
        /// Gets the submit time
        /// </summary>
        /// <param name="strSubmitTime">buffer to receive</param>
        /// <param name="strFormat">format string for date (empty for default)</param>
        /// <returns>true on success</returns>
        public bool GetSubmitTime(StringBuilder strSubmitTime, string strFormat)
        {
            return MessageGetSubmitTimeString(pObject, strSubmitTime, strSubmitTime.Capacity, strFormat);
        }

        /// <summary>
        /// Gets the TO field
        /// </summary>
        /// <param name="strTo">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetTo(StringBuilder strTo)
        {
            return MessageGetTo(pObject, strTo, strTo.Capacity);
        }

        /// <summary>
        /// Gets the CC field
        /// </summary>
        /// <param name="strCC">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetCC(StringBuilder strCC)
        {
            return MessageGetCC(pObject, strCC, strCC.Capacity);
        }

        /// <summary>
        /// Gets the BCC field
        /// </summary>
        /// <param name="strBCC">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetBCC(StringBuilder strBCC)
        {
            return MessageGetBCC(pObject, strBCC, strBCC.Capacity);
        }

        /// <summary>
        /// Gets the sensitivity of the message
        /// </summary>
        /// <returns>see NetMAPI.cs for Sensitivity enums</returns>
        public Sensitivity GetSensitivity()
        {
            return (Sensitivity)MessageGetSensitivity(pObject);
        }

        /// <summary>
        /// Gets the priority of the message
        /// </summary>
        /// <returns>see MAPI docs for PR_PRIORITY values</returns>
        public int GetPriority()
        {
            return MessageGetPriority(pObject);
        }

        /// <summary>
        /// Gets the importance of the message
        /// </summary>
        /// <returns>see MAPIMessage.cs for Importance enums</returns>
        public Importance GetImportance()
        {
            return (Importance)MessageGetImportance(pObject);
        }

        /// <summary>
        /// Get the recipients table, call this before calling GetNextRecipient
        /// </summary>
        /// <returns>true on success</returns>
        public bool GetRecipients()
        {
            return MessageGetRecipients(pObject);
        }

        /// <summary>
        /// Gets the next recipient
        /// </summary>
        /// <param name="strName">Name of recipient</param>
        /// <param name="strEmail">Email of recipient</param>
        /// <param name="nType">RecipientType (TO, CC, BCC)</param>
        /// <returns>true on success</returns>
        public bool GetNextRecipient(StringBuilder strName, StringBuilder strEmail, out RecipientType nType)
        {
            int nRecipientType;
            nType = RecipientType.UNKNOWN;
            if (MessageGetNextRecipient(pObject, strName, strName.Capacity, strEmail, strEmail.Capacity, out nRecipientType))
            {
                nType = (RecipientType)nRecipientType;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Gets the Reply To email address, if it is set
        /// </summary>
        /// <param name="strEmail">Email to reply to</param>
        /// <returns>true on success</returns>
        public bool GetReplyTo(StringBuilder strEmail)
        {
            return MessageGetReplyTo(pObject, strEmail, strEmail.Capacity);
        }

        /// <summary>
        /// Get the number of attachments
        /// </summary>
        /// <returns>the number of attachments</returns>
        public int GetAttachmentCount()
        {
            return MessageGetAttachmentCount(pObject);
        }

        /// <summary>
        /// Get the CID (Content ID) of attachment n
        /// </summary>
        /// <param name="strAttachmentCID">buffer to receive</param>
        /// <param name="nIndex">index of the attachment</param>
        /// <returns>true on success</returns>
        public bool GetAttachmentCID(StringBuilder strAttachmentCID, int nIndex)
        {
            return MessageGetAttachmentCID(pObject, strAttachmentCID, strAttachmentCID.Capacity, nIndex);
        }

        /// <summary>
        /// Get the name of attachment n
        /// </summary>
        /// <param name="strAttachmentName">buffer to receive</param>
        /// <param name="nIndex">index of the attachment</param>
        /// <returns>true on success</returns>
        public bool GetAttachmentName(StringBuilder strAttachmentName, int nIndex)
        {
            return MessageGetAttachmentName(pObject, strAttachmentName, strAttachmentName.Capacity, nIndex);
        }

        /// <summary>
        /// Saves one or all attachments
        /// </summary>
        /// <param name="strFolder">path to the folder to save to</param>
        /// <param name="nIndex">index of attachment to save (-1 for all)</param>
        /// <returns>true on success</returns>
        public bool SaveAttachment(string strFolder, int nIndex)
        {
            return MessageSaveAttachment(pObject, strFolder, nIndex);
        }

        /// <summary>
        /// Deletes one or all attachments
        /// </summary>
        /// <param name="nIndex">index of attachment to save (-1 for all)</param>
        /// <returns>true on success</returns>
        public bool DeleteAttachment(int nIndex)
        {
            return MessageDeleteAttachment(pObject, nIndex);
        }

        /// <summary>
        /// Sets the message status
        /// </summary>
        /// <param name="nMessageStatus">used only by WinCE, pass in MSGSTATUS_RECTYPE_SMS to send an SMS</param>
        public bool SetMessageStatus(int nMessageStatus)
        {
            return MessageSetMessageStatus(pObject, nMessageStatus);
        }

        /// <summary>
        /// Adds a recipient to a message to the TO list
        /// </summary>
        /// <param name="strEmail">email address of recipient</param>
        /// <returns>true on success</returns>
        public bool AddRecipient(string strEmail)
        {
            return AddRecipient(strEmail, MAPIMessage.RecipientType.TO, "SMTP");
        }

        /// <summary>
        /// Adds a recipient to a message
        /// </summary>
        /// <param name="strEmail">email address of recipient</param>
        /// <param name="nType">type of recipient (TO, CC and BCC)</param>
        /// <returns>true on success</returns>
        public bool AddRecipient(string strEmail, RecipientType nType)
        {
            return MessageAddRecipient(pObject, strEmail, (int)nType, "SMTP");
        }

        /// <summary>
        /// Adds a recipient to a message
        /// </summary>
        /// <param name="strEmail">email address of recipient</param>
        /// <param name="nType">type of recipient (TO, CC and BCC)</param>
        /// <param name="strAddrType">Address type of address (SMTP, SMS etc)</param>
        /// <returns>true on success</returns>
        public bool AddRecipient(string strEmail, RecipientType nType, string strAddrType)
        {
            return MessageAddRecipient(pObject, strEmail, (int)nType, strAddrType);
        }

        /// <summary>
        /// Sets the subject field
        /// </summary>
        /// <param name="strSubject">value to set</param>
        public void SetSubject(string strSubject)
        {
            MessageSetSubject(pObject, strSubject);
        }

        /// <summary>
        /// Sets the sender name and email
        /// </summary>
        /// <param name="strSenderName">Sender's name</param>
        /// <param name="strSenderEmail">Sender's SMTP email address</param>
        public void SetSender(string strSenderName, string strSenderEmail)
        {
            MessageSetSender(pObject, strSenderName, strSenderEmail);
        }

        /// <summary>
        /// Sets the message Received time
        /// </summary>
        /// <param name="dt">time (in local time) to set</param>
        /// <returns>true on success</returns>
        public bool SetReceivedTime(DateTime dt)
        {
            return SetReceivedTime(dt, true);
        }

        /// <summary>
        /// Sets the message Received time
        /// </summary>
        /// <param name="dt">time to set</param>
        /// <param name="bLocal">true if this time is local time</param>
        /// <returns>true on success</returns>
        public bool SetReceivedTime(DateTime dt, bool bLocal)
        {
            return MessageSetReceivedTime(pObject, dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second, bLocal);
        }

        /// <summary>
        /// Sets the message Submit time
        /// </summary>
        /// <param name="dt">time (in local time) to set</param>
        /// <returns>true on success</returns>
        public bool SetSubmitTime(DateTime dt)
        {
            return SetSubmitTime(dt, true);
        }

        /// <summary>
        /// Sets the message Submit time
        /// </summary>
        /// <param name="dt">time to set</param>
        /// <param name="bLocal">true if this time is local time</param>
        /// <returns>true on success</returns>
        public bool SetSubmitTime(DateTime dt, bool bLocal)
        {
            return MessageSetSubmitTime(pObject, dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second, bLocal);
        }

        /// <summary>
        /// Adds an attachment to the message
        /// </summary>
        /// <param name="strPath">path of the attachment</param>
        /// <returns>true on success</returns>
        public bool AddAttachment(string strPath)
        {
            return AddAttachment(strPath, "", "");
        }

        /// <summary>
        /// Adds an attachment to the message
        /// </summary>
        /// <param name="strPath">path of the attachment</param>
        /// <param name="strName">name of the attachment, "" for name of file</param>
        /// <returns>true on success</returns>
        public bool AddAttachment(string strPath, string strName)
        {
            return AddAttachment(strPath, strName, "");
        }

        /// <summary>
        /// Adds an attachment to the message
        /// </summary>
        /// <param name="strPath">path of the attachment</param>
        /// <param name="strName">name of the attachment, "" for name of file</param>
        /// <param name="strCID">content ID of the attachment</param>
        /// <returns>true on success</returns>
        public bool AddAttachment(string strPath, string strName, string strCID)
        {
            return MessageAddAttachment(pObject, strPath, strName, strCID);
        }

        /// <summary>
        /// Sets the read receipt flag
        /// </summary>
        /// <param name="bSet">true to set</param>
        public bool SetReadReceipt(bool bSet)
        {
            return MessageSetReadReceipt(pObject, bSet, "");
        }

        /// <summary>
        /// Sets the read receipt flag
        /// </summary>
        /// <param name="bSet">true to set</param>
        /// <param name="strReceiverEmail">email to key off</param>
        public bool SetReadReceipt(bool bSet, string strReceiverEmail)
        {
            return MessageSetReadReceipt(pObject, bSet, strReceiverEmail);
        }

        /// <summary>
        /// Sets the delivery receipt flag
        /// </summary>
        /// <param name="bSet">true to set</param>
        public bool SetDeliveryReceipt(bool bSet)
        {
            return MessageSetDeliveryReceipt(pObject, bSet);
        }

        /// <summary>
        /// Mark as private
        /// </summary>
        public bool MarkAsPrivate()
        {
            return MessageMarkAsPrivate(pObject);
        }

        /// <summary>
        /// Sets the sensitivity of the message
        /// </summary>
        /// <param name="nSensitivity">see NetMAPI.cs for Sensitivity enums</param>
        public bool SetSensitivity(Sensitivity nSensitivity)
        {
            return MessageSetSensitivity(pObject, (int)nSensitivity);
        }

        #endregion

        #region DLLCalls

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageCreate(IntPtr pMAPI, out IntPtr pMessage, int nImportance, bool bSaveToSentFolder, IntPtr pFolder);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSend(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int MessageShowForm(IntPtr pMAPI, IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageIsUnread(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageMarkAsRead(IntPtr pMessage, bool bRead);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetHeader(IntPtr pMessage, StringBuilder strHeader, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void MessageGetSenderName(IntPtr pMessage, StringBuilder strSenderName, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void MessageGetSenderEmail(IntPtr pMessage, StringBuilder strSenderEmail, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void MessageGetSubject(IntPtr pMessage, StringBuilder strSubject, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetReceivedTime(IntPtr pMessage, out int nYear, out int nMonth, out int nDay, out int nHour, out int nMinute, out int nSecond);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetReceivedTimeString(IntPtr pMessage, StringBuilder strReceivedTime, int nMaxLength, string szFormat);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetSubmitTime(IntPtr pMessage, out int nYear, out int nMonth, out int nDay, out int nHour, out int nMinute, out int nSecond);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetSubmitTimeString(IntPtr pMessage, StringBuilder strSubmitTime, int nMaxLength, string szFormat);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetTo(IntPtr pMessage, StringBuilder strTo, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetCC(IntPtr pMessage, StringBuilder strCC, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetBCC(IntPtr pMessage, StringBuilder strBCC, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int MessageGetSensitivity(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int MessageGetPriority(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int MessageGetImportance(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetRecipients(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetNextRecipient(IntPtr pMessage, StringBuilder strName, int nMaxLenName, StringBuilder strEmail, int nMaxLenEmail, out int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetReplyTo(IntPtr pMessage, StringBuilder strEmail, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int MessageGetAttachmentCount(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetAttachmentCID(IntPtr pMessage, StringBuilder strAttachmentCID, int nMaxLength, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageGetAttachmentName(IntPtr pMessage, StringBuilder strAttachmentName, int nMaxLength, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSaveAttachment(IntPtr pMessage, string strFolder, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageDeleteAttachment(IntPtr pMessage, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSetMessageStatus(IntPtr pMessage, int nMessageStatus);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageAddRecipient(IntPtr pMessage, string strEmail, int nType, string strAddrType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void MessageSetSubject(IntPtr pMessage, string strSubject);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void MessageSetSender(IntPtr pMessage, string strSenderName, string strSenderEmail);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSetReceivedTime(IntPtr pMessage, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond, bool bLocal);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSetSubmitTime(IntPtr pMessage, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond, bool bLocal);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageAddAttachment(IntPtr pMessage, string strPath, string strName, string strCID);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSetReadReceipt(IntPtr pMessage, bool bSet, string strReceiverEmail);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSetDeliveryReceipt(IntPtr pMessage, bool bSet);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageMarkAsPrivate(IntPtr pMessage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool MessageSetSensitivity(IntPtr pMessage, int nSensitivity);

        #endregion
    }
}

