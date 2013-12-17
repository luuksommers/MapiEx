////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIContact.cs
// Description: .NET Extended MAPI wrapper for Contacts
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
    public class MAPIContact : MAPIObject
    {
        public enum NameType { DISPLAY_NAME, GIVEN_NAME, MIDDLE_NAME, SURNAME };
        public enum AddressType { HOME, BUSINESS, OTHER };
        public enum PhoneType
        {
            PRIMARY_TELEPHONE_NUMBER, BUSINESS_TELEPHONE_NUMBER, HOME_TELEPHONE_NUMBER,
            CALLBACK_TELEPHONE_NUMBER, BUSINESS2_TELEPHONE_NUMBER, MOBILE_TELEPHONE_NUMBER,
            RADIO_TELEPHONE_NUMBER, CAR_TELEPHONE_NUMBER, OTHER_TELEPHONE_NUMBER,
            PAGER_TELEPHONE_NUMBER, PRIMARY_FAX_NUMBER, BUSINESS_FAX_NUMBER,
            HOME_FAX_NUMBER, TELEX_NUMBER, ISDN_NUMBER, ASSISTANT_TELEPHONE_NUMBER,
            HOME2_TELEPHONE_NUMBER, TTYTDD_PHONE_NUMBER, COMPANY_MAIN_PHONE_NUMBER
        }

        public MAPIContact()
        {
        }

        public MAPIContact(IntPtr pContact) : base(pContact)
        {
        }

        #region Contact Functions

        /// <summary>
        /// Create a new contact in the current folder
        /// </summary>
        /// <param name="mapi">NetMAPI session with an open folder</param>
        /// <returns>true on success</returns>
        public bool Create(NetMAPI mapi)
        {
            return ContactCreate(mapi.MAPI, out pObject, IntPtr.Zero);
        }

        /// <summary>
        /// Create a new contact in the current folder
        /// </summary>
        /// <param name="mapi">NetMAPI session with an open folder</param>
        /// <param name="pFolder">folder to create in</param>
        /// <returns>true on success</returns>
        public bool Create(NetMAPI mapi, IntPtr pFolder)
        {
            return ContactCreate(mapi.MAPI, out pObject, pFolder);
        }

        /// <summary>
        /// Gets the name of the contact
        /// </summary>
        /// <param name="strSubject">buffer to contain output string</param>
        /// <param name="nType">NameType to specify which name to retrieve</param>
        /// <returns>true on success</returns>
        public bool GetName(StringBuilder strSubject, NameType nType)
        {
            return ContactGetName(pObject, strSubject, strSubject.Capacity, (int)nType);
        }

        /// <summary>
        /// Gets the email address of the contact
        /// </summary>
        /// <param name="strEmail">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetEmail(StringBuilder strEmail)
        {
            return ContactGetEmail(pObject, strEmail, strEmail.Capacity, 1);
        }

        /// <summary>
        /// Gets the email address of the contact
        /// </summary>
        /// <param name="strEmail">buffer to contain output string</param>
        /// <param name="nIndex">index between 1 and 3</param>
        /// <returns>true on success</returns>
        public bool GetEmail(StringBuilder strEmail, int nIndex)
        {
            return ContactGetEmail(pObject, strEmail, strEmail.Capacity, nIndex);
        }

        /// <summary>
        /// Gets the email display as text of the contact
        /// </summary>
        /// <param name="strDisplayAs">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetEmailDisplayAs(StringBuilder strDisplayAs)
        {
            return ContactGetEmailDisplayAs(pObject, strDisplayAs, strDisplayAs.Capacity, 1);
        }

        /// <summary>
        /// Gets the email display as text of the contact
        /// </summary>
        /// <param name="strDisplayAs">buffer to contain output string</param>
        /// <param name="nIndex">index between 1 and 3</param>
        /// <returns>true on success</returns>
        public bool GetEmailDisplayAs(StringBuilder strDisplayAs, int nIndex)
        {
            return ContactGetEmailDisplayAs(pObject, strDisplayAs, strDisplayAs.Capacity, nIndex);
        }

        /// <summary>
        /// Gets the IM address of the contact
        /// </summary>
        /// <param name="strIMAddress">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetIMAddress(StringBuilder strIMAddress)
        {
            return ContactGetIMAddress(pObject, strIMAddress, strIMAddress.Capacity);
        }

        /// <summary>
        /// Gets the profession of the contact
        /// </summary>
        /// <param name="strProfession">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetProfession(StringBuilder strProfession)
        {
            return ContactGetProfession(pObject, strProfession, strProfession.Capacity);
        }

        /// <summary>
        /// Gets the homepage of the contact
        /// </summary>
        /// <param name="strHomePage">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetHomePage(StringBuilder strHomePage)
        {
            return ContactGetHomePage(pObject, strHomePage, strHomePage.Capacity);
        }

        /// <summary>
        /// Gets the phone number of the contact
        /// </summary>
        /// <param name="strPhoneNumber">buffer to contain output string</param>
        /// <param name="nType">PhoneType to specify which phone number to retrieve</param>
        /// <returns>true on success</returns>
        public bool GetPhoneNumber(StringBuilder strPhoneNumber, PhoneType nType)
        {
            return ContactGetPhoneNumber(pObject, strPhoneNumber, strPhoneNumber.Capacity, (int)nType);
        }

        /// <summary>
        /// Gets the full address of the contact
        /// </summary>
        /// <param name="nType">AddressType to specify which address to retrieve</param>
        /// <returns>true on success</returns>
        public bool GetAddress(out ContactAddress address, AddressType nType)
        {
            IntPtr pAddress;
            if (ContactGetAddress(pObject, out pAddress, (int)nType))
            {
                address = new ContactAddress();
                address.Type = nType;

                StringBuilder s = new StringBuilder(NetMAPI.DefaultBufferSize);
                AddressGetStreet(pAddress, s, s.Capacity);
                address.Street = s.ToString();
                AddressGetCity(pAddress, s, s.Capacity);
                address.City = s.ToString();
                AddressGetStateOrProvince(pAddress, s, s.Capacity);
                address.StateOrProvince = s.ToString();
                AddressGetPostalCode(pAddress, s, s.Capacity);
                address.PostalCode = s.ToString();
                AddressGetCountry(pAddress, s, s.Capacity);
                address.Country = s.ToString();

                AddressClose(pAddress);
                return true;
            }

            address = null;
            return false;
        }

        /// <summary>
        /// Gets the postal address, for quick access to a single string
        /// </summary>
        /// <param name="strAddress">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetPostalAddress(StringBuilder strAddress)
        {
            return ContactGetPostalAddress(pObject, strAddress, strAddress.Capacity);
        }

        /// <summary>
        /// Gets the size of the notes field, so you can allocate a large enough StringBuilder
        /// </summary>
        /// <param name="bRTF"></param>
        /// <returns>size in characters of the notes field</returns>
        public int GetNotesSize(bool bRTF)
        {
            return ContactGetNotesSize(pObject, bRTF);
        }

        /// <summary>
        /// Gets the notes field
        /// </summary>
        /// <param name="strNotes">buffer to contain output string</param>
        /// <param name="bRTF">in Rich Text Format or just plain text</param>
        /// <returns>true on success</returns>
        public bool GetNotes(StringBuilder strNotes, bool bRTF)
        {
            return ContactGetNotes(pObject, strNotes, strNotes.Capacity, bRTF);
        }

        /// <summary>
        /// Gets the sensitivity of the contact
        /// </summary>
        /// <returns>see NetMAPI.cs for Sensitivity enums</returns>
        public Sensitivity GetSensitivity()
        {
            return (Sensitivity)ContactGetSensitivity(pObject);
        }

        /// <summary>
        /// Gets the title field
        /// </summary>
        /// <param name="strTitle">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetTitle(StringBuilder strTitle)
        {
            return ContactGetTitle(pObject, strTitle, strTitle.Capacity);
        }

        /// <summary>
        /// Gets the Company field
        /// </summary>
        /// <param name="strCompany">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetCompany(StringBuilder strCompany)
        {
            return ContactGetCompany(pObject, strCompany, strCompany.Capacity);
        }

        /// <summary>
        /// Gets the prefix (Mr., Dr. etc)
        /// </summary>
        /// <param name="strPrefix">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetDisplayNamePrefix(StringBuilder strPrefix)
        {
            return ContactGetDisplayNamePrefix(pObject, strPrefix, strPrefix.Capacity);
        }

        /// <summary>
        /// Gets the generation (Sr, Jr, III etc)
        /// </summary>
        /// <param name="strGeneration">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetGeneration(StringBuilder strGeneration)
        {
            return ContactGetGeneration(pObject, strGeneration, strGeneration.Capacity);
        }

        /// <summary>
        /// Gets the department field
        /// </summary>
        /// <param name="strDepartment">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetDepartment(StringBuilder strDepartment)
        {
            return ContactGetDepartment(pObject, strDepartment, strDepartment.Capacity);
        }

        /// <summary>
        /// Gets the office field
        /// </summary>
        /// <param name="strOffice">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetOffice(StringBuilder strOffice)
        {
            return ContactGetOffice(pObject, strOffice, strOffice.Capacity);
        }

        /// <summary>
        /// Gets the name of the contact's manager
        /// </summary>
        /// <param name="strManagerName">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetManagerName(StringBuilder strManagerName)
        {
            return ContactGetManagerName(pObject, strManagerName, strManagerName.Capacity);
        }

        /// <summary>
        /// Gets the name of the contact's assistant
        /// </summary>
        /// <param name="strAssistantName">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetAssistantName(StringBuilder strAssistantName)
        {
            return ContactGetAssistantName(pObject, strAssistantName, strAssistantName.Capacity);
        }

        /// <summary>
        /// Get's the contact's nickname
        /// </summary>
        /// <param name="strNickName">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetNickName(StringBuilder strNickName)
        {
            return ContactGetNickName(pObject, strNickName, strNickName.Capacity);
        }

        /// <summary>
        /// Gets the name of the contact's spouse
        /// </summary>
        /// <param name="strSpouseName">buffer to contain output string</param>
        /// <returns>true on success</returns>
        public bool GetSpouseName(StringBuilder strSpouseName)
        {
            return ContactGetSpouseName(pObject, strSpouseName, strSpouseName.Capacity);
        }

        /// <summary>
        /// Gets the contact's birthday
        /// </summary>
        /// <param name="dtBirthday">DateTime of Birthday</param>
        /// <returns>true on success</returns>
        public bool GetBirthday(out DateTime dtBirthday)
        {
            int nYear, nMonth, nDay;
            bool bResult = ContactGetBirthday(pObject, out nYear, out nMonth, out nDay);
            dtBirthday = new DateTime(nYear, nMonth, nDay);
            return bResult;
        }

        /// <summary>
        /// Gets the contact's wedding anniversary
        /// </summary>
        /// <param name="dtAnniversary">DateTime of Anniversary</param>
        /// <returns>true on success</returns>
        public bool GetAnniversary(out DateTime dtAnniversary)
        {
            int nYear, nMonth, nDay;
            bool bResult = ContactGetAnniversary(pObject, out nYear, out nMonth, out nDay);
            dtAnniversary = new DateTime(nYear, nMonth, nDay);
            return bResult;
        }

        /// <summary>
        /// Gets the categories (stored under MV property "Keywords"
        /// </summary>
        /// <param name="strCategories">buffer to contain categories separated by semicolons</param>
        /// <returns>true on success</returns>
        public bool GetCategories(StringBuilder strCategories)
        {
            return ContactGetCategories(pObject, strCategories, strCategories.Capacity);
        }

        /// <summary>
        /// Sets the name field
        /// After setting all name fields you should call UpdateDisplayName
        /// </summary>
        /// <param name="strName">string to set</param>
        /// <param name="nType">NameType to set to</param>
        /// <returns>true on success</returns>
        public bool SetName(string strName, NameType nType)
        {
            return ContactSetName(pObject, strName, (int)nType);
        }

        /// <summary>
        /// Sets the outlook email address field
        /// </summary>
        /// <param name="strEmail">email address to set to</param>
        /// <param name="nIndex">index (1 to 3)</param>
        /// <returns>true on success</returns>
        public bool SetEmail(string strEmail, int nIndex)
        {
            return ContactSetEmail(pObject, strEmail, nIndex);
        }

        /// <summary>
        /// Sets the outlook email address display as field
        /// </summary>
        /// <param name="strDisplayAs">display as string</param>
        /// <param name="nIndex">index (1 to 3)</param>
        /// <returns>true on success</returns>
        public bool SetEmailDisplayAs(string strDisplayAs, int nIndex)
        {
            return ContactSetEmailDisplayAs(pObject, strDisplayAs, nIndex);
        }

        /// <summary>
        /// Sets the home page field
        /// </summary>
        /// <param name="strHomePage">string to set</param>
        /// <returns>true on success</returns>
        public bool SetHomePage(string strHomePage)
        {
            return ContactSetHomePage(pObject, strHomePage);
        }

        /// <summary>
        /// Sets the IM address field
        /// </summary>
        /// <param name="strIMAddress">string to set</param>
        /// <returns>true on success</returns>
        public bool SetIMAddress(string strIMAddress)
        {
            return ContactSetIMAddress(pObject, strIMAddress);
        }

        /// <summary>
        /// Sets the profession field
        /// </summary>
        /// <param name="strProfession">string to set</param>
        /// <returns>true on success</returns>
        public bool SetProfession(string strProfession)
        {
            return ContactSetProfession(pObject, strProfession);
        }

        /// <summary>
        /// Sets the phone number
        /// </summary>
        /// <param name="strPhoneNumber">string to set</param>
        /// <param name="nType">PhoneType to specify which phone number to set</param>
        /// <returns>true on success</returns>
        public bool SetPhoneNumber(string strPhoneNumber, PhoneType nType)
        {
            return ContactSetPhoneNumber(pObject, strPhoneNumber, (int)nType);
        }

        /// <summary>
        /// Sets the full address of the contact
        /// </summary>
        /// <param name="nType">AddressType to specify which address to set</param>
        /// <returns>true on success</returns>
        public bool SetAddress(ContactAddress address, AddressType nType)
        {
            IntPtr pAddress;
            if (!ContactGetAddress(pObject, out pAddress, (int)nType)) return false;

            AddressSetStreet(pAddress, address.Street);
            AddressSetCity(pAddress, address.City);
            AddressSetStateOrProvince(pAddress, address.StateOrProvince);
            AddressSetPostalCode(pAddress, address.PostalCode);
            AddressSetCountry(pAddress, address.Country);

            bool bResult = ContactSetAddress(pObject, pAddress, (int)nType);
            AddressClose(pAddress);
            return bResult;
        }

        /// <summary>
        /// Sets the postal address fields (and the corresponding outlook checkbox)
        /// </summary>
        /// <param name="nType">AddressType to set</param>
        /// <returns>true on success</returns>
        public bool SetPostalAddress(AddressType nType)
        {
            return ContactSetPostalAddress(pObject, (int)nType);
        }

        /// <summary>
        /// Updates the outlook display address field (use after changing address fields)
        /// </summary>
        /// <param name="nType">AddressType to set</param>
        /// <returns>true on success</returns>
        public bool UpdateDisplayAddress(AddressType nType)
        {
            return ContactUpdateDisplayAddress(pObject, (int)nType);
        }

        /// <summary>
        /// Sets the notes field
        /// </summary>
        /// <param name="strNotes">string to set</param>
        /// <param name="bRTF"></param>
        /// <returns>true on success</returns>
        public bool SetNotes(string strNotes, bool bRTF)
        {
            return ContactSetNotes(pObject, strNotes, bRTF);
        }

        /// <summary>
        /// Sets the sensitivity value
        /// </summary>
        /// <param name="nSensitivity"></param>
        /// <returns>true on success</returns>
        public bool SetSensitivity(int nSensitivity)
        {
            return ContactSetSensitivity(pObject, nSensitivity);
        }

        /// <summary>
        /// Sets the File As field
        /// </summary>
        /// <param name="strFileAs">string to set</param>
        /// <returns>true on success</returns>
        public bool SetFileAs(string strFileAs)
        {
            return ContactSetFileAs(pObject, strFileAs);
        }

        /// <summary>
        /// Sets the title field
        /// </summary>
        /// <param name="strTitle">string to set</param>
        /// <returns>true on success</returns>
        public bool SetTitle(string strTitle)
        {
            return ContactSetTitle(pObject, strTitle);
        }

        /// <summary>
        /// Sets the company field
        /// </summary>
        /// <param name="strCompany">string to set</param>
        /// <returns>true on success</returns>
        public bool SetCompany(string strCompany)
        {
            return ContactSetCompany(pObject, strCompany);
        }

        /// <summary>
        /// Sets the display name prefix (Mr., Dr., etc)
        /// After setting this you should call UpdateDisplayName
        /// </summary>
        /// <param name="strPrefix">string to set</param>
        /// <returns>true on success</returns>
        public bool SetDisplayNamePrefix(string strPrefix)
        {
            return ContactSetDisplayNamePrefix(pObject, strPrefix);
        }

        /// <summary>
        /// Sets the generation (Jr., III etc)
        /// After setting this you should call UpdateDisplayName
        /// </summary>
        /// <param name="strGeneration">string to set</param>
        /// <returns>true on success</returns>
        public bool SetGeneration(string strGeneration)
        {
            return ContactSetGeneration(pObject, strGeneration);
        }

        /// <summary>
        /// Updates the display name, use this after setting one or more name properties (name, generation, prefix etc)
        /// </summary>
        /// <returns>true on success</returns>
        public bool UpdateDisplayName()
        {
            return ContactUpdateDisplayName(pObject);
        }

        /// <summary>
        /// Sets the department field
        /// </summary>
        /// <param name="strDepartment">string to set to</param>
        /// <returns>true on success</returns>
        public bool SetDepartment(string strDepartment)
        {
            return ContactSetDepartment(pObject, strDepartment);
        }

        /// <summary>
        /// sets the office field
        /// </summary>
        /// <param name="strOffice">string to set to</param>
        /// <returns>true on success</returns>
        public bool SetOffice(string strOffice)
        {
            return ContactSetOffice(pObject, strOffice);
        }

        /// <summary>
        /// Sets the manager name field
        /// </summary>
        /// <param name="strManagerName">string to set to</param>
        /// <returns>true on success</returns>
        public bool SetManagerName(string strManagerName)
        {
            return ContactSetManagerName(pObject, strManagerName);
        }

        /// <summary>
        /// Sets the assistant name field
        /// </summary>
        /// <param name="strAssistantName">string to set to</param>
        /// <returns>true on success</returns>
        public bool SetAssistantName(string strAssistantName)
        {
            return ContactSetAssistantName(pObject, strAssistantName);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strNickName"></param>
        /// <returns>true on success</returns>
        public bool SetNickName(string strNickName)
        {
            return ContactSetNickName(pObject, strNickName);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strSpouseName"></param>
        /// <returns>true on success</returns>
        public bool SetSpouseName(string strSpouseName)
        {
            return ContactSetSpouseName(pObject, strSpouseName);
        }

        /// <summary>
        /// Sets the birthday
        /// </summary>
        /// <param name="dtBirthday">date of birth</param>
        /// <returns>true on success</returns>
        public bool SetBirthday(DateTime dtBirthday)
        {
            return ContactSetBirthday(pObject, dtBirthday.Year, dtBirthday.Month, dtBirthday.Day);
        }

        /// <summary>
        /// Sets the wedding anniversary
        /// </summary>
        /// <param name="dtAnniversary">date of anniversary</param>
        /// <returns>true on success</returns>
        public bool SetAnniversary(DateTime dtAnniversary)
        {
            return ContactSetAnniversary(pObject, dtAnniversary.Year, dtAnniversary.Month, dtAnniversary.Day);
        }

        /// <summary>
        /// Sets the categories
        /// </summary>
        /// <param name="strCategories">string of semicolon separated categories</param>
        /// <returns>true on success</returns>
        public bool SetCategories(string strCategories)
        {
            return ContactSetCategories(pObject, strCategories);
        }

        /// <summary>
        /// Sets the outlook picture
        /// </summary>
        /// <param name="strPath">Path to a valid JPG or PNG (other may work) file</param>
        /// <returns>true on success</returns>
        public bool SetPicture(string strPath)
        {
            return ContactSetPicture(pObject, strPath);
        }

        #endregion

        #region DLLCalls

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactCreate(IntPtr pMAPI, out IntPtr pContact, IntPtr pFolder);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetName(IntPtr pContact, StringBuilder strName, int nMaxLength, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetEmail(IntPtr pContact, StringBuilder strEmail, int nMaxLength, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetEmailDisplayAs(IntPtr pContact, StringBuilder strDisplayAs, int nMaxLength, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetHomePage(IntPtr pContact, StringBuilder strHomePage, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetIMAddress(IntPtr pContact, StringBuilder strIMAddress, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetProfession(IntPtr pContact, StringBuilder strProfession, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetPhoneNumber(IntPtr pContact, StringBuilder strPhoneNumber, int nMaxLength, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetAddress(IntPtr pContact, out IntPtr pAddress, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetPostalAddress(IntPtr pContact, StringBuilder strAddress, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int ContactGetNotesSize(IntPtr pContact, bool bRTF);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetNotes(IntPtr pContact, StringBuilder strNotes, int nMaxLength, bool bRTF);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern int ContactGetSensitivity(IntPtr pContact);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetTitle(IntPtr pContact, StringBuilder strTitle, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetCompany(IntPtr pContact, StringBuilder strCompany, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetDisplayNamePrefix(IntPtr pContact, StringBuilder strDisplayNamePrefix, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetGeneration(IntPtr pContact, StringBuilder strGeneration, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetDepartment(IntPtr pContact, StringBuilder strDepartment, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetOffice(IntPtr pContact, StringBuilder strOffice, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetManagerName(IntPtr pContact, StringBuilder strManagerName, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetAssistantName(IntPtr pContact, StringBuilder strAssistantName, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetNickName(IntPtr pContact, StringBuilder strNickName, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetSpouseName(IntPtr pContact, StringBuilder strSpouseName, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetBirthday(IntPtr pContact, out int nYear, out int nMonth, out int nDay);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetAnniversary(IntPtr pContact, out int nYear, out int nMonth, out int nDay);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactGetCategories(IntPtr pContact, StringBuilder strField, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetName(IntPtr pContact, string strName, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetEmail(IntPtr pContact, string strEmail, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetEmailDisplayAs(IntPtr pContact, string strDisplayAs, int nIndex);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetHomePage(IntPtr pContact, string strHomePage);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetIMAddress(IntPtr pContact, string strIMAddress);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetProfession(IntPtr pContact, string strProfession);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetPhoneNumber(IntPtr pContact, string strPhoneNumber, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetAddress(IntPtr pContact, IntPtr pAddress, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetPostalAddress(IntPtr pContact, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactUpdateDisplayAddress(IntPtr pContact, int nType);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetNotes(IntPtr pContact, string strNotes, bool bRTF);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetSensitivity(IntPtr pContact, int nSensitivity);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetFileAs(IntPtr pContact, string strFileAs);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetTitle(IntPtr pContact, string strTitle);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetCompany(IntPtr pContact, string strCompany);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetDisplayNamePrefix(IntPtr pContact, string strPrefix);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetGeneration(IntPtr pContact, string strGeneration);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactUpdateDisplayName(IntPtr pContact);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetDepartment(IntPtr pContact, string strDepartment);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetOffice(IntPtr pContact, string strOffice);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetManagerName(IntPtr pContact, string strManagerName);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetAssistantName(IntPtr pContact, string strAssistantName);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetNickName(IntPtr pContact, string strNickName);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetSpouseName(IntPtr pContact, string strSpouseName);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetBirthday(IntPtr pContact, int nYear, int nMonth, int nDay);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetAnniversary(IntPtr pContact, int nYear, int nMonth, int nDay);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetCategories(IntPtr pContact, string strCategories);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool ContactSetPicture(IntPtr pContact, string strPath);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressClose(IntPtr pAddress);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressGetStreet(IntPtr pAddress, StringBuilder strStreet, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressGetCity(IntPtr pAddress, StringBuilder strCity, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressGetStateOrProvince(IntPtr pAddress, StringBuilder strStateOrProvince, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressGetPostalCode(IntPtr pAddress, StringBuilder strPostalCode, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressGetCountry(IntPtr pAddress, StringBuilder strCountry, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressSetStreet(IntPtr pAddress, string strStreet);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressSetCity(IntPtr pAddress, string strCity);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressSetStateOrProvince(IntPtr pAddress, string strStateOrProvince);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressSetPostalCode(IntPtr pAddress, string strPostalCode);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern void AddressSetCountry(IntPtr pAddress, string strCountry);

        #endregion

        public class ContactAddress
        {
            public ContactAddress()
            {
            }

            private AddressType nType;
            private string strStreet;
            private string strCity;
            private string strStateOrProvince;
            private string strPostalCode;
            private string strCountry;

            public AddressType Type
            {
                get { return nType; }
                set { nType = value; }
            }

            public string Street
            {
                get { return strStreet; }
                set { strStreet = value; }
            }

            public string City
            {
                get { return strCity; }
                set { strCity = value; }
            }

            public string StateOrProvince
            {
                get { return strStateOrProvince; }
                set { strStateOrProvince = value; }
            }

            public string PostalCode
            {
                get { return strPostalCode; }
                set { strPostalCode = value; }
            }

            public string Country
            {
                get { return strCountry; }
                set { strCountry = value; }
            }
        };
    }
}
