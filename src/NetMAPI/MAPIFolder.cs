////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIFolder.cs
// Description: .NET Extended MAPI wrapper for Folders
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
    /// Folders
    /// </summary>
    public class MAPIFolder : MAPIObject
    {
        protected bool bDisposeObject;

        public MAPIFolder()
        {
            bDisposeObject = true;
        }

        public MAPIFolder(IntPtr pFolder) : this(pFolder,true)
        {
        }

        public MAPIFolder(IntPtr pFolder, bool bDisposeObject) : base(pFolder)
        {
            this.bDisposeObject = bDisposeObject;
        }

        protected override void Dispose(bool disposing)
        {
            if(bDisposeObject) base.Dispose(disposing);
            pObject = IntPtr.Zero;
        }

        #region Folder Functions

        /// <summary>
        /// Opens the Hierarchy of this folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool GetHierarchy()
        {
            return FolderGetHierarchy(pObject);
        }
        /// <summary>
        /// High Level function to open a sub-folder by iterating recursively (DFS) over all folders 
        /// (use instead of manually calling GetHierarchy and GetNextSubFolder)
        /// </summary>
        /// <param name="strSubFolder">name of the folder</param>
        /// <param name="folder">MAPIFolder object</param>
        /// <returns>true on success</returns>
        public bool OpenSubFolder(string strSubFolder,out MAPIFolder folder)
        {
            folder = null;
            IntPtr pSubFolder = FolderOpenSubFolder(pObject, strSubFolder);
            if (pSubFolder != IntPtr.Zero)
            {
                folder = new MAPIFolder(pSubFolder);
            }
            return (folder != null);
        }
        
        /// <summary>
        /// Creates a sub-folder or opens it if it already exists
        /// </summary>
        /// <param name="strSubFolder">name of the folder</param>
        /// <param name="folder">MAPIFolder object</param>
        /// <returns>true on success</returns>
        public bool CreateSubFolder(string strSubFolder,out MAPIFolder folder)
        {
            folder = null;
            IntPtr pSubFolder = FolderCreateSubFolder(pObject, strSubFolder);
            if (pSubFolder != IntPtr.Zero)
            {
                folder = new MAPIFolder(pSubFolder);
            }
            return (folder != null);
        }

        /// <summary>
        /// Deletes a sub-folder and ALL sub-folders/messages
        /// </summary>
        /// <param name="strSubFolder">name of folder</param>
        /// <returns>true on success</returns>
        public bool DeleteSubFolderByName(string strSubFolder)
        {
            return FolderDeleteSubFolderByName(pObject, strSubFolder);
        }

        /// <summary>
        /// Deletes a sub-folder and ALL sub-folders/messages
        /// </summary>
        /// <param name="folder">folder object to delete</param>
        /// <returns>true on success</returns>
        public bool DeleteSubFolder(MAPIFolder folder)
        {
            return FolderDeleteSubFolder(pObject, folder.pObject);
        }

        /// <summary>
        /// Gets the contents of this folder
        /// </summary>
        /// <returns>true on success</returns>
        public bool GetContents()
        {
            return FolderGetContents(pObject);
        }

        /// <summary>
        /// Returns the row count of the currently opened contents
        /// </summary>
        public int RowCount
        {
            get { return FolderGetRowCount(pObject); }
        }

        /// <summary>
        /// Sorts the contents based on receive time
        /// </summary>
        /// <param name="bAscending">ascending or descending</param>
        /// <returns>true on success</returns>
        public bool SortContents(bool bAscending)
        {
            return SortContents(bAscending, SortFields.SORT_RECEIVED_TIME);
        }

        /// <summary>
        /// Sorts the contents based on the sort field specified
        /// </summary>
        /// <param name="bAscending">ascending or descending</param>
        /// <param name="sortField">one of the sort fields defined in NetMAPI</param>
        /// <returns>true on success</returns>
        public bool SortContents(bool bAscending, SortFields sortField)
        {
            return FolderSortContents(pObject, bAscending, (int)sortField);
        }

        /// <summary>
        /// Places a restriction on the current contents to pass unread items only
        /// </summary>
        /// <param name="bUnreadOnly">true to filter unread items</param>
        /// <returns>true on success</returns>
        public bool SetUnreadOnly(bool bUnreadOnly)
        {
            return FolderSetUnreadOnly(pObject, bUnreadOnly);
        }

        /// <summary>
        /// Gets the next message in this folder
        /// </summary>
        /// <param name="pMessage">MAPIMessage object</param>
        /// <returns>true on success</returns>
        public bool GetNextMessage(out MAPIMessage message)
        {
            IntPtr pMessage;
            message = null;
            if (FolderGetNextMessage(pObject, out pMessage))
            {
                message = new MAPIMessage(pMessage);
            }
            return (message != null);
        }

        /// <summary>
        /// Gets the next contact (assuming this is the Contacts folder)
        /// </summary>
        /// <param name="contact">MAPIContact object</param>
        /// <returns>true on success</returns>
        public bool GetNextContact(out MAPIContact contact)
        {
            IntPtr pContact;
            contact = null;
            if (FolderGetNextContact(pObject, out pContact))
            {
                contact = new MAPIContact(pContact);
            }
            return (contact != null);
        }

        /// <summary>
        /// Gets the next appointment (assuming this is the Calendar folder)
        /// </summary>
        /// <param name="appointment">MAPIAppointment object</param>
        /// <returns>true on success</returns>
        public bool GetNextAppointment(out MAPIAppointment appointment)
        {
            IntPtr pAppointment;
            appointment = null;
            if (FolderGetNextAppointment(pObject, out pAppointment))
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
            if (FolderGetNextSubFolder(pObject, out pFolder, strFolderName, strFolderName.Capacity))
            {
                folder = new MAPIFolder(pFolder);
            }
            return (folder != null);
        }

        public bool DeleteMessage(MAPIMessage message)
        {
            return FolderDeleteMessage(pObject, message.Pointer);
        }

        public bool CopyMessage(MAPIMessage message, MAPIFolder folderDest)
        {
            return FolderCopyMessage(pObject, message.Pointer, folderDest.Pointer);
        }

        public bool MoveMessage(MAPIMessage message, MAPIFolder folderDest)
        {
            return FolderMoveMessage(pObject, message.Pointer, folderDest.Pointer);
        }

        public bool DeleteContact(MAPIContact contact)
        {
            return FolderDeleteContact(pObject, contact.Pointer);
        }

        public bool DeleteAppointment(MAPIAppointment appointment)
        {
            return FolderDeleteAppointment(pObject, appointment.Pointer);
        }

        #endregion

        #region DLLCalls

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderGetHierarchy(IntPtr pFolder);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern IntPtr FolderOpenSubFolder(IntPtr pFolder, string strSubFolder);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern IntPtr FolderCreateSubFolder(IntPtr pFolder, string strSubFolder);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderDeleteSubFolderByName(IntPtr pFolder, string strSubFolder);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderDeleteSubFolder(IntPtr pFolder, IntPtr pSubFolder);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderGetContents(IntPtr pFolder);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int FolderGetRowCount(IntPtr pFolder);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderSortContents(IntPtr pFolder, bool bAscending, int nSortField);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderSetUnreadOnly(IntPtr pFolder, bool bUnreadOnly);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderGetNextMessage(IntPtr pFolder,out IntPtr pMessage);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderGetNextContact(IntPtr pFolder,out IntPtr pContact);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderGetNextAppointment(IntPtr pFolder, out IntPtr pAppointment);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderGetNextSubFolder(IntPtr pFolder,out IntPtr pSubFolder,StringBuilder strFolder, int nMaxLength);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderDeleteMessage(IntPtr pFolder, IntPtr pMessage);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderCopyMessage(IntPtr pFolder, IntPtr pMessage, IntPtr pFolderDest);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderMoveMessage(IntPtr pFolder, IntPtr pMessage, IntPtr pFolderDest);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderDeleteContact(IntPtr pFolder, IntPtr pContact);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool FolderDeleteAppointment(IntPtr pFolder, IntPtr pAppointment);
        
        #endregion
    }
}

