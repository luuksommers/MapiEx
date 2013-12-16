////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// File: Defines.cs
// Description: .NET Extended MAPI wrapper defines
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
    partial class NetMAPI
    {
        public const CharSet DefaultCharSet = CharSet.Ansi;
        public const CallingConvention DefaultCallingConvention = CallingConvention.Cdecl;
        
        public const string MAPIExDLL = "MAPIEx.dll";
        public const int DefaultBufferSize = 1024;
    };
}
