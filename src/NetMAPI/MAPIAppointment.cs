////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: MAPIAppointment.cs
// Description: .NET Extended MAPI wrapper for Appointments
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
    /// Appointments
    /// </summary>
    public class MAPIAppointment : MAPIObject
    {
        public MAPIAppointment()
        {
        }

        public MAPIAppointment(IntPtr pAppointment) : base(pAppointment)
        {
            
        }

        #region Appointment Functions

        /// <summary>
        /// Get the appointment subject
        /// </summary>
        /// <param name="strSubject">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetSubject(StringBuilder strSubject)
        {
            return AppointmentGetSubject(pObject, strSubject, strSubject.Capacity);
        }

        /// <summary>
        /// Get the appointment location
        /// </summary>
        /// <param name="strSubject">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetLocation(StringBuilder strLocation)
        {
            return AppointmentGetLocation(pObject, strLocation, strLocation.Capacity);
        }

        /// <summary>
        /// Gets the start time
        /// </summary>
        /// <param name="dt">DateTime to receive</param>
        /// <returns>true on success</returns>
        public bool GetStartTime(out DateTime dt)
        {
            int nYear, nMonth, nDay, nHour, nMinute, nSecond;
            bool bResult = AppointmentGetStartTime(pObject, out nYear, out nMonth, out nDay, out nHour, out nMinute, out nSecond);
            dt = new DateTime(nYear, nMonth, nDay, nHour, nMinute, nSecond);
            return bResult;
        }

        /// <summary>
        /// Gets the start time using the default format (MM/dd/yyyy hh:mm:ss tt)
        /// </summary>
        /// <param name="strStartTime">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetStartTime(StringBuilder strStartTime)
        {
            return AppointmentGetStartTimeString(pObject, strStartTime, strStartTime.Capacity, "");
        }

        /// <summary>
        /// Gets the start time
        /// </summary>
        /// <param name="strStartTime">buffer to receive</param>
        /// <param name="strFormat">format string for date (empty for default)</param>
        /// <returns>true on success</returns>
        public bool GetStartTime(StringBuilder strStartTime, string strFormat)
        {
            return AppointmentGetStartTimeString(pObject, strStartTime, strStartTime.Capacity, strFormat);
        }

        /// <summary>
        /// Gets the end time
        /// </summary>
        /// <param name="dt">DateTime to receive</param>
        /// <returns>true on success</returns>
        public bool GetEndTime(out DateTime dt)
        {
            int nYear, nMonth, nDay, nHour, nMinute, nSecond;
            bool bResult = AppointmentGetEndTime(pObject, out nYear, out nMonth, out nDay, out nHour, out nMinute, out nSecond);
            dt = new DateTime(nYear, nMonth, nDay, nHour, nMinute, nSecond);
            return bResult;
        }

        /// <summary>
        /// Gets the end time using the default format (MM/dd/yyyy hh:mm:ss tt)
        /// </summary>
        /// <param name="strEndTime">buffer to receive</param>
        /// <returns>true on success</returns>
        public bool GetEndTime(StringBuilder strEndTime)
        {
            return AppointmentGetEndTimeString(pObject, strEndTime, strEndTime.Capacity, "");
        }

        /// <summary>
        /// Gets the end time
        /// </summary>
        /// <param name="strEndTime">buffer to receive</param>
        /// <param name="strFormat">format string for date (empty for default)</param>
        /// <returns>true on success</returns>
        public bool GetEndTime(StringBuilder strEndTime, string strFormat)
        {
            return AppointmentGetEndTimeString(pObject, strEndTime, strEndTime.Capacity, strFormat);
        }

        /// <summary>
        /// Set the appointment subject
        /// </summary>
        /// <param name="strSubject">subject to set</param>
        /// <returns>true on success</returns>
        public bool SetSubject(string strSubject)
        {
            return AppointmentSetSubject(pObject, strSubject);
        }

        /// <summary>
        /// Set the appointment location
        /// </summary>
        /// <param name="strSubject">location to set</param>
        /// <returns>true on success</returns>
        public bool SetLocation(string strLocation)
        {
            return AppointmentSetLocation(pObject, strLocation);
        }

        /// <summary>
        /// Sets the start time
        /// </summary>
        /// <param name="dtStart">DateTime of start</param>
        /// <returns>true on success</returns>
        public bool SetStartTime(DateTime dtStart)
        {
            return AppointmentSetStartTime(pObject, dtStart.Year, dtStart.Month, dtStart.Day, dtStart.Hour, dtStart.Minute, dtStart.Second);
        }

        /// <summary>
        /// Sets the end time
        /// </summary>
        /// <param name="dtEnd">DateTime of end</param>
        /// <returns>true on success</returns>
        public bool SetEndTime(DateTime dtEnd)
        {
            return AppointmentSetEndTime(pObject, dtEnd.Year, dtEnd.Month, dtEnd.Day, dtEnd.Hour, dtEnd.Minute, dtEnd.Second);
        }

        #endregion

        #region DLLCalls

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentGetSubject(IntPtr pAppointment, StringBuilder strSubject, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentGetLocation(IntPtr pAppointment, StringBuilder strLocation, int nMaxLength);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentGetStartTime(IntPtr pAppointment, out int nYear, out int nMonth, out int nDay, out int nHour, out int nMinute, out int nSecond);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentGetStartTimeString(IntPtr pAppointment, StringBuilder strStartTime, int nMaxLength, string szFormat);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentGetEndTime(IntPtr pAppointment, out int nYear, out int nMonth, out int nDay, out int nHour, out int nMinute, out int nSecond);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentGetEndTimeString(IntPtr pAppointment, StringBuilder strEndTime, int nMaxLength, string szFormat);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentSetSubject(IntPtr pAppointment, string strSubject);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentSetLocation(IntPtr pAppointment, string strLocation);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentSetStartTime(IntPtr pAppointment, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond);

        [DllImport(NetMAPI.MAPIExDLL, CharSet = NetMAPI.DefaultCharSet, CallingConvention = NetMAPI.DefaultCallingConvention)]
        protected static extern bool AppointmentSetEndTime(IntPtr pAppointment, int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond);

        #endregion
    }
}

