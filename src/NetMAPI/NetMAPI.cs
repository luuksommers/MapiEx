////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: NetMAPI.cs
// Description: .NET Extended MAPI wrapper
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
    public enum SortFields { SORT_RECEIVED_TIME, SORT_SUBJECT };
    public enum Sensitivity { SENSITIVITY_NONE, SENSITIVITY_PERSONAL, SENSITIVITY_PRIVATE, SENSITIVITY_COMPANY_CONFIDENTIAL };

    /// <summary>
    /// Extended MAPI Wrapper for .NET, to use a Unicode version change DefaultCharSet and compile MAPIEx for Unicode.
    /// </summary>
    public partial class NetMAPI : IDisposable
    {
        protected IntPtr pMAPI;

        public NetMAPI()
        {
            pMAPI = IntPtr.Zero;
        }

        ~NetMAPI()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            Logout();
        }

        /// <summary>
        /// returns the internal MAPI pointer used to call the DLL interface
        /// </summary>
        public IntPtr MAPI
        {
            get { return pMAPI; }
        }

        /// <summary>
        /// Attempts to Initialize MAPIEx, if the MAPIEx DLL is not present or loadable, a silent false is returned
        /// </summary>
        /// <returns>true on success</returns>
        public static bool Init()
        {
            return Init(false);
        }

        /// <summary>
        /// Attempts to Initialize MAPIEx, if the MAPIEx DLL is not present or loadable, a silent false is returned
        /// </summary>
        /// <param name="bInitAsService">Use this setting if you are creating a service.</param>
        /// <returns>true on success</returns>
        public static bool Init(bool bInitAsService)
        {
            try
            {
                return MAPIInit(false, bInitAsService);
            }
            catch (DllNotFoundException)
            {
                return false;
            }
        }

        /// <summary>
        /// Terminates MAPIEx, call this before exiting your application
        /// </summary>
        public static void Term()
        {
            MAPITerm();
        }

        #pragma warning disable 0162
        public static string MarshalString(IntPtr szString)
        {
            if (szString != IntPtr.Zero) 
            {
                if (DefaultCharSet == CharSet.Ansi) return Marshal.PtrToStringAnsi(szString);
                else return Marshal.PtrToStringUni(szString);
            }
            return "";
        }

        #region Profiles, Message Store

        /// <summary>
        /// Login with the default profile
        /// </summary>
        /// <returns>true on success</returns>
        public bool Login()
        {
            return Login("", false);
        }

        /// <summary>
        /// Login with a specific profile
        /// </summary>
        /// <param name="strProfile">Profile to open for example "Outlook"</param>
        /// <returns>true on success</returns>
        public bool Login(string strProfile)
        {
            return Login(strProfile, false);
        }

        /// <summary>
        /// Login with a specific profile
        /// </summary>
        /// <param name="strProfile">Profile to open for example "Outlook"</param>
        /// <param name="bInitAsService">Use this setting if you are creating a service.</param>
        /// <returns>true on success</returns>
        public bool Login(string strProfile, bool bInitAsService)
        {
            if (pMAPI == IntPtr.Zero) pMAPI = MAPILogin(strProfile, bInitAsService);
            return (pMAPI != IntPtr.Zero);
        }

        /// <summary>
        /// Logout of the current session, if you login successfully be sure to log out so the memory used by 
        /// MAPIEx is released.
        /// </summary>
        public void Logout()
        {
            if (pMAPI != IntPtr.Zero)
            {
                MAPILogout(pMAPI);
                pMAPI = IntPtr.Zero;
            }
        }

        /// <summary>
        /// Open the default Message Store
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenMessageStore()
        {
            return OpenMessageStore("");
        }

        /// <summary>
        /// Opens a specific Message Store
        /// </summary>
        /// <param name="strStore">Name of the message store to open</param>
        /// <returns>true on success</returns>
        public bool OpenMessageStore(string strStore)
        {
            return MAPIOpenMessageStore(pMAPI, strStore);
        }

        /// <summary>
        /// Sometimes you need to know the current profile's name, this retrieves it
        /// </summary>
        /// <param name="strProfileName">Name of the currently opened profile</param>
        /// <returns>true on success</returns>
        public bool GetProfileName(StringBuilder strProfileName)
        {
            return MAPIGetProfileName(pMAPI, strProfileName, strProfileName.Capacity);
        }

        /// <summary>
        /// Sometimes you need to know the current profile's email, this retrieves it
        /// </summary>
        /// <param name="strProfileName">Email of the currently opened profile</param>
        /// <returns>true on success</returns>
        public bool GetProfileEmail(StringBuilder strProfileEmail)
        {
            return MAPIGetProfileEmail(pMAPI, strProfileEmail, strProfileEmail.Capacity);
        }
        #endregion

        #region POOM Functions

        /// <summary>
        /// Get the contents list of the currently opened folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool GetPOOMContents()
        {
            return POOMGetContents(MAPIGetPOOM(pMAPI));
        }

        /// <summary>
        /// Sort the current contents
        /// </summary>
        /// <param name="bDescending">true if descending order</param>
        /// <param name="strSortField">sort field surrounded by [] ie [FileAs]</param>
        /// <returns></returns>
        public bool SortPOOMContents(bool bDescending, string strSortField)
        {
            return POOMSortContents(MAPIGetPOOM(pMAPI), bDescending, strSortField);
        }

        /// <summary>
        /// Get the row count of the current contents
        /// </summary>
        /// <returns>number of rows</returns>
        public int GetPOOMRowCount()
        {
            return POOMGetRowCount(MAPIGetPOOM(pMAPI));
        }

        #endregion

        #region Folders

        /// <summary>
        /// returns the currently opened folder, used in advanced folder operations
        /// </summary>
        public MAPIFolder Folder
        {
            get
            {
                IntPtr pFolder = MAPIGetFolder(pMAPI);
                if (pFolder!=IntPtr.Zero) return new MAPIFolder(pFolder, false);
                return null;
            }
            set 
            {
                MAPIFolder folder = value;
                MAPISetFolder(pMAPI, folder.Pointer); 
            }
        }

        /// <summary>
        /// Opens a folder via a string
        /// </summary>
        /// <param name="folderName">name to match</param>
        /// <returns>true on success</returns>
        public bool OpenFolder(string folderName)
        {
            return (MAPIOpenFolder(pMAPI, folderName, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens a folder via a string
        /// </summary>
        /// <param name="folderName">name to match</param>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenFolder(string folderName, bool bInternal)
        {
            IntPtr pFolder = MAPIOpenFolder(pMAPI, folderName, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the root folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenRootFolder()
        {
            return (MAPIOpenRootFolder(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the root folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenRootFolder(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenRootFolder(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// OPens the Inbox folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenInbox()
        {
            return (MAPIOpenInbox(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// OPens the Inbox folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenInbox(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenInbox(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Outbox folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenOutbox()
        {
            return (MAPIOpenOutbox(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the Outbox folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenOutbox(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenOutbox(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Sent Items folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenSentItems()
        {
            return (MAPIOpenSentItems(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the Sent Items folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenSentItems(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenSentItems(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Deleted Items folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenDeletedItems()
        {
            return (MAPIOpenDeletedItems(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the Deleted Items folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenDeletedItems(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenDeletedItems(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Contacts folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenContacts()
        {
            return (MAPIOpenContacts(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the Contacts folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenContacts(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenContacts(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Drafts folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenDrafts()
        {
            return (MAPIOpenDrafts(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the Drafts folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenDrafts(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenDrafts(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Calendar folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenCalendar()
        {
            return (MAPIOpenCalendar(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the Calendar folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenCalendar(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenCalendar(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Junk E-mail folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool OpenJunkFolder()
        {
            return (MAPIOpenJunkFolder(pMAPI, true) != IntPtr.Zero);
        }

        /// <summary>
        /// Opens the Junk E-mail folder
        /// </summary>
        /// <param name="bInternal">true to let MAPIEx handle disposal of this folder</param>
        /// <returns>true on success</returns>
        public MAPIFolder OpenJunkFolder(bool bInternal)
        {
            IntPtr pFolder = MAPIOpenJunkFolder(pMAPI, bInternal);
            if (pFolder != IntPtr.Zero) return new MAPIFolder(pFolder);
            return null;
        }

        /// <summary>
        /// Opens the Hierarchy of the currently opened folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool GetHierarchy()
        {
            return MAPIGetHierarchy(pMAPI);
        }

        #endregion

        #region Messages, Contacts

        /// <summary>
        /// Opens the contents of the currently opened folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool GetContents()
        {
            return MAPIGetContents(pMAPI);
        }

        /// <summary>
        /// Returns the row count of the currently opened contents
        /// </summary>
        public int RowCount
        {
            get { return MAPIGetRowCount(pMAPI); }
        }

        /// <summary>
        /// Sorts the contents based on receive time
        /// </summary>
        /// <param name="bAscending">ascending or descending</param>
        /// <returns>true on success</returns>
        public bool SortContents(bool bAscending)
        {
            return MAPISortContents(pMAPI, bAscending, (int)SortFields.SORT_RECEIVED_TIME);
        }

        /// <summary>
        /// Sorts the contents based on the sort field specified
        /// </summary>
        /// <param name="bAscending">ascending or descending</param>
        /// <param name="sortField">one of the sort fields above</param>
        /// <returns>true on success</returns>
        public bool SortContents(bool bAscending, SortFields sortField)
        {
            return MAPISortContents(pMAPI, bAscending, (int)sortField);
        }

        /// <summary>
        /// Places a restriction on the current contents to pass unread items only
        /// </summary>
        /// <param name="bUnreadOnly">true to filter unread items</param>
        /// <returns>true on success</returns>
        public bool SetUnreadOnly(bool bUnreadOnly)
        {
            return MAPISetUnreadOnly(pMAPI, bUnreadOnly);
        }

        /// <summary>
        /// Gets the next message in the currently opened folder
        /// </summary>
        /// <param name="pMessage">MAPIMessage object</param>
        /// <param name="bUnreadOnly">only get unread messages</param>
        /// <returns>true on sucess</returns>
        public bool GetNextMessage(out MAPIMessage message)
        {
            IntPtr pMessage;
            message = null;
            if (MAPIGetNextMessage(pMAPI, out pMessage))
            {
                message = new MAPIMessage(pMessage);
            }
            return (message != null);
        }

        /// <summary>
        /// Gets the next contact (assumed you have the contacts folder open)
        /// </summary>
        /// <param name="contact">MAPIContact object</param>
        /// <returns>true on success</returns>
        public bool GetNextContact(out MAPIContact contact)
        {
            IntPtr pContact;
            contact = null;
            if (MAPIGetNextContact(pMAPI, out pContact))
            {
                contact = new MAPIContact(pContact);
            }
            return (contact != null);
        }

        /// <summary>
        /// Gets the next appointment (assumed you have the Calendar folder open)
        /// </summary>
        /// <param name="appointment">MAPIAppointment object</param>
        /// <returns>true on success</returns>
        public bool GetNextAppointment(out MAPIAppointment appointment)
        {
            IntPtr pAppointment;
            appointment = null;
            if (MAPIGetNextAppointment(pMAPI, out pAppointment))
            {
                appointment = new MAPIAppointment(pAppointment);
            }
            return (appointment != null);
        }

        /// <summary>
        /// Retrieves the next sub-folder in the currently opened hierarchy
        /// </summary>
        /// <param name="folder">MAPIFolder object</param>
        /// <param name="strFolderName">name of the folder</param>
        /// <returns>true on success</returns>
        public bool GetNextSubFolder(out MAPIFolder folder, StringBuilder strFolderName)
        {
            IntPtr pFolder;
            folder = null;
            if (MAPIGetNextSubFolder(pMAPI, out pFolder, strFolderName, strFolderName.Capacity))
            {
                folder = new MAPIFolder(pFolder);
            }
            return (folder != null);
        }

        #endregion

        #region DLLCalls
        // These shouldn't be called directly by the client, use the interface in NetMAPI for this.

        // Initialize and Terminate

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIInit(bool bMultiThreadedNotifications, bool bInitAsService);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPITerm();

        // Profiles, Message Store

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPILogin(string strProfile, bool bInitAsService);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern void MAPILogout(IntPtr pMAPI);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIOpenMessageStore(IntPtr pMAPI, string strStore);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIGetFolder(IntPtr pMAPI);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPISetFolder(IntPtr pMAPI, IntPtr pFolder);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetProfileName(IntPtr pMAPI, StringBuilder strProfileName, int nMaxLength);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetProfileEmail(IntPtr pMAPI, StringBuilder strProfileEmail, int nMaxLength);

        // POOM functions
        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIGetPOOM(IntPtr pMAPI);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool POOMGetContents(IntPtr pPOOM);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool POOMSortContents(IntPtr pPOOM, bool bDescending, string strSortField);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern int POOMGetRowCount(IntPtr pPOOM);

        // Folders
        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenFolder(IntPtr pMAPI, string folderName, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenRootFolder(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenInbox(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenOutbox(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenSentItems(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenDeletedItems(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenContacts(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenDrafts(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenCalendar(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern IntPtr MAPIOpenJunkFolder(IntPtr pMAPI, bool bInternal);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetHierarchy(IntPtr pMAPI);

        // Messages

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetContents(IntPtr pMAPI);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern int MAPIGetRowCount(IntPtr pMAPI);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPISortContents(IntPtr pMAPI, bool bAscending, int nSortField);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPISetUnreadOnly(IntPtr pMAPI, bool bUnreadOnly);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetNextMessage(IntPtr pMAPI, out IntPtr pMessage);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetNextContact(IntPtr pMAPI, out IntPtr pContact);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetNextAppointment(IntPtr pMAPI, out IntPtr pAppointment);

        [DllImport(MAPIExDLL, CharSet = DefaultCharSet, CallingConvention = DefaultCallingConvention)]
        protected static extern bool MAPIGetNextSubFolder(IntPtr pMAPI, out IntPtr pFolder, StringBuilder strFolder, int nMaxLength);

        #endregion
    }
}
