////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: TestNetMAPI.cs
// Description: Test program for .NET Extended MAPI wrapper
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
using System.Text;
using MAPIEx;

namespace TestNetMAPI
{
    /// <summary>
    /// Test program for NetMAPI Wrapper
    /// </summary>
    class TestNetMAPI
    {
        /// <summary>
        /// TestMAPI Main
        /// </summary>
        [STAThread]

        static void Main(string[] args)
        {
            if (NetMAPI.Init())
            {
                NetMAPI mapi = new NetMAPI();
                if (mapi.Login())
                {
                    StringBuilder strText = new StringBuilder(NetMAPI.DefaultBufferSize);
                    if (mapi.GetProfileName(strText)) Console.WriteLine("Profile Name: " + strText);
                    if (mapi.GetProfileEmail(strText)) Console.WriteLine("Profile Email: " + strText);

                    if (mapi.OpenMessageStore())
                    {
                        // uncomment the functions you're interested in and/or step through these to see how each thing works.
//                         SendTest(mapi);
//                         SendCIDTest(mapi);
//                         FolderTest(mapi);
                         ReceiveTest(mapi);
//                         ContactsTest(mapi);
//                         CopyMessageTest(mapi);
//                        MoveMessageTest(mapi);
//                         DeleteMessageTest(mapi);
//                         AppointmentTest(mapi);
                    }
                    mapi.Logout();
                }
                NetMAPI.Term();
            }
            Console.ReadLine();
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // To iterate through folders:
        //		-first open up a MAPI session and login
        //		-then open the message store you want to access
        //		-then open the folder and get the hierarchy table
        //		-iterate through the folders using GetNextFolder()
        //		
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public static void FolderTest(NetMAPI mapi)
        {
            // another way to do this is with the line below
            // if (mapi.OpenRootFolder()) EnumerateSubFolders(mapi.Folder); 

            if (mapi.OpenRootFolder() && mapi.GetHierarchy()) 
            {
                StringBuilder s = new StringBuilder(NetMAPI.DefaultBufferSize);
                MAPIFolder folder;
                while (mapi.GetNextSubFolder(out folder, s))
                {
                    Console.WriteLine("Folder: " + s.ToString());
                    EnumerateSubFolders(folder);
                    folder.Dispose();
                }
            }
        }

        private static void EnumerateSubFolders(MAPIFolder folder)
        {
            if (folder.GetHierarchy())
            {
                StringBuilder s = new StringBuilder(NetMAPI.DefaultBufferSize);
                MAPIFolder subFolder;
                while (folder.GetNextSubFolder(out subFolder, s))
                {
                    Console.WriteLine("SubFolder: " + s.ToString());
                    EnumerateSubFolders(subFolder);
                    subFolder.Dispose();
                }
            }
        } 

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // This function displays your first contact in your Contacts folder
        //
        // To do this:
        //		-first open up a MAPI session and login
        //		-then open the message store you want to send 
        //		-open your contacts folder 
        //
        // Use GetName to get the name (default DISPLAY NAME, but can be Initials, First Name etc) 
        // Use GetEmail to get the email address (named property) 
        // Use GetAddress to get the mailing address of CContactAddress::AddressType
        // Use GetPhoneNumber supplying a phone number property (ie BUSINESS_TELEPHONE_NUMBER)
        // Use GetNotes to get the notes in either plain text (default) or RTF
        //
        // Remember to Dispose of the contact when you're done with it!
        //
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void ContactsTest(NetMAPI mapi)
        {
            if (mapi.OpenContacts() && mapi.GetContents())
            {
                // sort by name (stored in PR_SUBJECT)
                mapi.SortContents(true, SortFields.SORT_SUBJECT);

                MAPIContact contact;
                StringBuilder strText = new StringBuilder(NetMAPI.DefaultBufferSize);
                if (mapi.GetNextContact(out contact))
                {
                    if (contact.GetName(strText, MAPIContact.NameType.DISPLAY_NAME)) Console.WriteLine("Contact: " + strText);
                    if (contact.GetEmail(strText)) Console.WriteLine("Email: " + strText);
                    if (contact.GetCategories(strText)) Console.WriteLine("Categories: " + strText);

                    MAPIContact.ContactAddress address;
                    if (contact.GetAddress(out address, MAPIContact.AddressType.BUSINESS))
                    {
                        Console.WriteLine(address.Street);
                        Console.WriteLine(address.City);
                        Console.WriteLine(address.StateOrProvince);
                        Console.WriteLine(address.Country);
                        Console.WriteLine(address.PostalCode);
                    }

                    if (contact.GetPhoneNumber(strText, MAPIContact.PhoneType.BUSINESS_TELEPHONE_NUMBER)) Console.WriteLine("Phone: " + strText);

                    //usually you would call GetNotesSize first to ensure the buffer is large enough
                    if (contact.GetNotes(strText, false)) Console.WriteLine("Notes: " + strText);

                    contact.Dispose();
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // To send a message:
        //		-first open up a MAPI session and login
        //		-then open the message store you want to access 
        //		-then open the outbox
        //		-create a new message, set its priority if you like 
        //		-set its properties, recipients and attachments
        //		-call send
        //
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void SendTest(NetMAPI mapi)
        {
            if (mapi.OpenOutbox())
            {
                MAPIMessage message = new MAPIMessage();
                if (message.Create(mapi, MAPIMessage.Importance.IMPORTANCE_LOW))
                {
                    message.SetSender("Support", "support@nospam.com");
                    message.SetSubject("Subject");

                    // user SetBody for ANSI text, SetRTF for HTML and Rich Text
                    message.SetRTF("<html><body><font size=2 color=red face=Arial><span style='font-size:10.0pt;font-family:Arial;color:red'>Body</font></body></html>");

                    message.AddRecipient("noel@nospam.com");
                    message.AddRecipient("noel@nospam.com");

                    if (message.Send()) Console.WriteLine("Sent Successfully");
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // To embed images inside of HTML text:
        //		-send a message as usual, setting its HTML text
        //		-add an <IMG tag with src="cid:<contentID>"
        //		-add an attachment using the same CID
        //
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void SendCIDTest(NetMAPI mapi)
        {
            if (mapi.OpenOutbox())
            {
                MAPIMessage message = new MAPIMessage();
                if (message.Create(mapi, MAPIMessage.Importance.IMPORTANCE_LOW))
                {
                    message.SetSender("Support", "support@nospam.com");
                    message.SetSubject("Subject");

                    string strCID = "123456789";
                    StringBuilder strImage = new StringBuilder();
                    strImage.Append("<html><body><IMG alt=\"\" src=\"cid:");
                    strImage.Append(strCID);
                    strImage.Append("\" border=0></body></html>");
                    message.SetRTF(strImage.ToString());


                    message.AddRecipient("noel@nospam.com");
                    message.AddAttachment(@"c:\Temp\Pic.jpg", "", strCID); // obviously you'll have to supply an image here

                    if (message.Send()) Console.WriteLine("Sent Successfully");
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // To receive a message:
        //		-first open up a MAPI session and login
        //		-then open the message store you want to access 
        //		-then open the inbox and get the contents table
        //		-iterate through the message using GetNextMessage() (sample below gets only unread messages)
        //		-save attachments (if any) using SaveAttachment() if you like
        //
        // Remember to Dispose of the message when you're done with it!
        //
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void ReceiveTest(NetMAPI mapi)
        {
            if (mapi.OpenInbox() && mapi.GetContents())
            {
                mapi.SortContents(false);
                mapi.SetUnreadOnly(false);

                MAPIMessage message;
                StringBuilder s = new StringBuilder(NetMAPI.DefaultBufferSize);
                while (mapi.GetNextMessage(out message))
                {
                    Console.Write("Message from '");
                    message.GetSenderName(s);
                    Console.Write(s.ToString() + "' (");
                    message.GetSenderEmail(s);
                    Console.Write(s.ToString() + "), subject '");
                    message.GetSubject(s);
                    Console.Write(s.ToString() + "', received: ");
                    message.GetReceivedTime(s);
                    Console.Write(s.ToString() + "\n\n");

                    // use message.GetBody(), message.GetHTML(), or message.GetRTF() to get the text body
                    // GetBody() can autodetect the source
                    string strBody;
                    message.GetBody(out strBody, true);
                    Console.Write(strBody + "\n");

                    message.Dispose();
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // This function creates a folder (opens if exists) and copies the first unread message if any to this folder
        //
        // To do this:
        //		-first open up a MAPI session and login
        //		-then open the message store you want to access 
        //		-then open the folder (probably inbox) and get the contents table
        //		-open the message you want to move
        //		-create (open if exists) the folder you want to move to
        //		-copy the message 
        //
        // You can also move and delete the message, but I wanted the sample to be non destructive just in case
        //
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void CopyMessageTest(NetMAPI mapi)
        {
            if (mapi.OpenInbox() && mapi.GetContents())
            {
                mapi.SetUnreadOnly(true);

                MAPIMessage message;
                StringBuilder s = new StringBuilder(NetMAPI.DefaultBufferSize);
                if (mapi.GetNextMessage(out message))
                {
                    Console.Write("Copying message from '");
                    message.GetSenderName(s);
                    Console.Write(s.ToString() + "' (");
                    message.GetSenderEmail(s);
                    Console.Write(s.ToString() + "), subject '");
                    message.GetSubject(s);
                    Console.Write(s.ToString() + "'\n");

                    MAPIFolder folder=mapi.Folder;
                    MAPIFolder subfolder;
                    if (folder.CreateSubFolder("TestFolder", out subfolder)) 
                    {
                        if (folder.CopyMessage(message, subfolder))
                        {
                            Console.WriteLine("Message copied successfully");
                        }
                        subfolder.Dispose();
                    }
                    message.Dispose();
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // This function moves the first message from the Inbox to the Sent Items folder (requested by user)
        //
        // NOTE: MAPIEx takes care of the internal folder, but in this case you have an external folder that must 
        // be disposed of.
        // 
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void MoveMessageTest(NetMAPI mapi)
        {
            if (mapi.OpenInbox() && mapi.GetContents())
            {
                mapi.SortContents(false);

                MAPIMessage message;
                StringBuilder s = new StringBuilder(NetMAPI.DefaultBufferSize);
                if (mapi.GetNextMessage(out message))
                {
                    Console.Write("Moving message from '");
                    message.GetSenderName(s);
                    Console.Write(s.ToString() + "' (");
                    message.GetSenderEmail(s);
                    Console.Write(s.ToString() + "), subject '");
                    message.GetSubject(s);
                    Console.Write(s.ToString() + "'\n");

                    MAPIFolder sentItems = mapi.OpenSentItems(false);
                    if (mapi.Folder.MoveMessage(message, sentItems))
                    {
                        Console.WriteLine("Message moved successfully");
                    }
                    sentItems.Dispose();
                    message.Dispose();
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // This function deletes the first message from the Inbox 
        //
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void DeleteMessageTest(NetMAPI mapi)
        {
            if (mapi.OpenInbox() && mapi.GetContents())
            {
                mapi.SortContents(false);

                MAPIMessage message;
                StringBuilder s = new StringBuilder(NetMAPI.DefaultBufferSize);
                if (mapi.GetNextMessage(out message))
                {
                    Console.Write("Deleting message from '");
                    message.GetSenderName(s);
                    Console.Write(s.ToString() + "'\n");

                    mapi.Folder.DeleteMessage(message);
                    message.Dispose();
                }
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        // This function reads an appointment from the Calendar
        //
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static void AppointmentTest(NetMAPI mapi)
        {
            if (mapi.OpenCalendar() && mapi.GetContents())
            {
                mapi.SortContents(false, SortFields.SORT_RECEIVED_TIME);

                MAPIAppointment appointment;
                StringBuilder strText = new StringBuilder(NetMAPI.DefaultBufferSize);
                if (mapi.GetNextAppointment(out appointment))
                {
                    if (appointment.GetSubject(strText)) Console.WriteLine("Subject: " + strText);
                    if (appointment.GetLocation(strText)) Console.WriteLine("Location: " + strText);
                    if (appointment.GetStartTime(strText)) Console.WriteLine("Start Time: " + strText);
                    if (appointment.GetEndTime(strText)) Console.WriteLine("End Time: " + strText);

                    appointment.Dispose();
                }
            }
        }
    }
}
