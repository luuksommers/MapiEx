////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIObject.cs
// Description: .NET Extended MAPI base class for MAPI Items (Messages, Contacts, etc)
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
    /// Contacts
    /// </summary>
    public class MAPIObject : IDisposable
    {
        public enum MessageEditorFormat { EDITOR_FORMAT_DONTKNOW, EDITOR_FORMAT_PLAINTEXT, EDITOR_FORMAT_HTML, EDITOR_FORMAT_RTF };
        
        protected IntPtr pObject;

        public MAPIObject()
        {
            pObject = IntPtr.Zero;
        }

        public MAPIObject(IntPtr pObject)
        {
            this.pObject = pObject;
        }

        ~MAPIObject()
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
            if (pObject != IntPtr.Zero)
            {
                ObjectClose(pObject);
                pObject = IntPtr.Zero;
            }
        }

        public IntPtr Pointer { get { return pObject; } }

        #region Object Functions

        /// <summary>
        /// Saves (and closes) the message
        /// </summary>
        /// <returns>true on success</returns>
        public bool Save()
        {
            return ObjectSave(pObject, true);
        }

        /// <summary>
        /// Saves the object
        /// </summary>
        /// <param name="bClose">if false, the object will remain open for further changes</param>
        /// <returns>true on success</returns>
        public bool Save(bool bClose)
        {
            return ObjectSave(pObject, bClose);
        }

        /// <summary>
        /// Get the message flags 
        /// </summary>
        /// <returns>all flags set (MSG_UNSENT etc)</returns>
        public int GetMessageFlags()
        {
            return ObjectGetMessageFlags(pObject);
        }

        /// <summary>
        /// Gets string representation of the Entry ID
        /// </summary>
        /// <param name="strEntryID">field to store the entry id</param>
        /// <returns>true on success</returns>
        public bool GetEntryID(StringBuilder strEntryID)
        {
            return ObjectGetEntryID(pObject, strEntryID, strEntryID.Capacity);
        }

        /// <summary>
        /// Get the message class of the object
        /// </summary>
        /// <param name="strMessageClass">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetMessageClass(StringBuilder strMessageClass)
        {
            return ObjectGetMessageClass(pObject, strMessageClass, strMessageClass.Capacity);
        }

        /// <summary>
        /// Gets the Message Editor Format
        /// </summary>
        /// <returns>format of the item</returns>
        public MessageEditorFormat GetMessageEditorFormat()
        {
            return (MessageEditorFormat)ObjectGetMessageEditorFormat(pObject);
        }

        /// <summary>
        /// Advanced function for getting properties that may not be exposed via the interface below
        /// </summary>
        /// <param name="strProperty">buffer to contain output string</param>
        /// <param name="ulProperty">MAPI property you want to retrieve</param>
        /// <param name="bStream">used for large strings that must be streamed</param>
        /// <returns>true on success</returns>
        public bool GetPropertyString(StringBuilder strProperty, uint ulProperty, bool bStream)
        {
            return ObjectGetPropertyString(pObject, ulProperty, strProperty, strProperty.Capacity, bStream);
        }

        /// <summary>
        /// Gets a named property
        /// </summary>
        /// <param name="strFieldName">name of the field</param>
        /// <param name="strField">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetNamedProperty(string strFieldName, StringBuilder strField)
        {
            return ObjectGetNamedProperty(pObject, strFieldName, strField, strField.Capacity);
        }

        /// <summary>
        /// Gets the ANSI text field of the item, or optionally the "suitable" text
        /// </summary>
        /// <param name="strBody">text content</param>
        /// <param name="bAutoDetect">use the MessageEditorFormat to determine which field to query</param>
        /// <returns>true on success</returns>
        public bool GetBody(out string strBody, bool bAutoDetect)
        {
            IntPtr szString;
            if (!ObjectGetBody(pObject, out szString, bAutoDetect)) szString = IntPtr.Zero;
            strBody = NetMAPI.MarshalString(szString);
            return (szString != IntPtr.Zero);
        }

        /// <summary>
        /// Gets the HTML text field of the item
        /// </summary>
        /// <param name="strHTML">HTML text content</param>
        /// <returns>true on success</returns>
        public bool GetHTML(ref string strHTML)
        {
            IntPtr szString;
            if (!ObjectGetHTML(pObject, out szString)) szString = IntPtr.Zero;
            strHTML = NetMAPI.MarshalString(szString);
            return (szString!=IntPtr.Zero);
        }

        /// <summary>
        /// Gets the rich text field of the item
        /// </summary>
        /// <param name="strRTF">rich text content</param>
        /// <returns>true on success</returns>
        public bool GetRTF(ref string strRTF)
        {
            IntPtr szString;
            if (!ObjectGetRTF(pObject, out szString)) szString = IntPtr.Zero;
            strRTF = NetMAPI.MarshalString(szString);
            return (szString != IntPtr.Zero);
        }

        /// <summary>
        /// Advanced function, frees internal memory used by the last GetBody, GetHTML, or GetRTF call.
        /// no need to call this normally (it will be freed automatically) but some users will want control of when.
        /// </summary>
        public void FreeBody()
        {
            ObjectFreeBody();
        }

        /// <summary>
        /// Sets the message flags
        /// </summary>
        /// <param name="nFlags">value to set (may want to get flags and OR before calling)</param>
        /// <returns>true on success</returns>
        public bool SetMessageFlags(int nFlags)
        {
            return ObjectSetMessageFlags(pObject, nFlags);
        }

        /// <summary>
        /// Sets the Message Editor Format
        /// </summary>
        /// <param name="nFormat">format to set</param>
        /// <returns>true on success</returns>
        public bool SetMessageEditorFormat(MessageEditorFormat nFormat)
        {
            return ObjectSetMessageEditorFormat(pObject, (int)nFormat);
        }

        /// <summary>
        /// Advanced function for setting properties that may not be exposed via the interface below
        /// </summary>
        /// <param name="strProperty">string to set</param>
        /// <param name="ulProperty">MAPI property you want to retrieve</param>
        /// <param name="bStream">used for large strings that must be streamed</param>
        /// <returns>true on success</returns>
        public bool SetPropertyString(string strProperty, uint ulProperty, bool bStream)
        {
            return ObjectSetPropertyString(pObject, ulProperty, strProperty, bStream);
        }

        /// <summary>
        /// Sets a named property
        /// </summary>
        /// <param name="strFieldName">name of the field</param>
        /// <param name="strField">string to set field value</param>
        /// <returns>true on success</returns>
        public bool SetNamedProperty(string strFieldName, string strField, bool bCreate)
        {
            return ObjectSetNamedProperty(pObject, strFieldName, strField, bCreate);
        }

        /// <summary>
        /// Sets the ANSI text body of the Object
        /// </summary>
        /// <param name="strBody">ANSI text to set</param>
        /// <returns>true on success</returns>
        public bool SetBody(string strBody)
        {
            return ObjectSetBody(pObject, strBody);
        }

        /// <summary>
        /// Sets the HTML text body of the Object
        /// </summary>
        /// <param name="strBody">HTML text to set</param>
        /// <returns>true on success</returns>
        public bool SetHTML(string strHTML)
        {
            return ObjectSetHTML(pObject, strHTML);
        }

        /// <summary>
        /// Sets the rich text body of the Object
        /// </summary>
        /// <param name="strBody">rich text to set</param>
        /// <returns>true on success</returns>
        public bool SetRTF(string strRTF)
        {
            return ObjectSetRTF(pObject, strRTF);
        }

        #endregion

        #region DLLCalls

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern void ObjectClose(IntPtr pObject);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSave(IntPtr pObject, bool bClose);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern int ObjectGetMessageFlags(IntPtr pObject);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectGetEntryID(IntPtr pObject, StringBuilder strEntryID, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectGetMessageClass(IntPtr pObject, StringBuilder szMessageClass, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern int ObjectGetMessageEditorFormat(IntPtr pObject);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectGetPropertyString(IntPtr pObject, uint ulProperty, StringBuilder strProperty, int nMaxLength, bool bStream);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectGetNamedProperty(IntPtr pObject, string strFieldName, StringBuilder strField, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectGetBody(IntPtr pObject, out IntPtr szBody, bool bAutoDetect);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectGetHTML(IntPtr pObject, out IntPtr szHTML);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectGetRTF(IntPtr pObject, out IntPtr szRTF);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern void ObjectFreeBody();

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSetMessageFlags(IntPtr pObject, int nFlags);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSetMessageEditorFormat(IntPtr pObject, int nFormat);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSetPropertyString(IntPtr pObject, uint ulProperty, string strProperty, bool bStream);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSetNamedProperty(IntPtr pObject, string strFieldName, string strField, bool bCreate);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSetBody(IntPtr pObject, string strBody);
        
        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSetHTML(IntPtr pObject, string strHTML);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet)]
        protected static extern bool ObjectSetRTF(IntPtr pObject, string strRTF);

        #endregion
    }
}
