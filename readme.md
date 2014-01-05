# Extended Mapi for C# #

### Original code by Noel Dillabough ###
http://www.codeproject.com/Articles/10881/MAPIEx-Extended-MAPI-Wrapper

### Usage ###
```csharp
if (NetMAPI.Init())
{
    NetMAPI mapi = new NetMAPI();
    if (mapi.Login())
    {
        if (mapi.OpenMessageStore())
        {
            if (mapi.OpenInbox() && mapi.GetContents())
            {
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
        mapi.Logout();
    }
    NetMAPI.Term();
}
```


### Setup ###
1. Visual Studio 2010 Professional or higher for MFC headers
2. When using VS2013, download [Visual Studio 2013 C++ Multi Byte Support](http://go.microsoft.com/?linkid=9832071)
3. When using VS2012/VS2013, VS2010 must also be installed to support the platform toolset v100 (C++ 2010 redistributable)

### Deployment ###
This library uses the [Visual C++ 2010 redistributable](http://www.microsoft.com/en-us/download/details.aspx?id=5555)

### FAQ ###
__Does this support x64 MAPI__

No, MAPI only supports x64 when an x64 office is installed. See [Building MAPI Applications on 32-Bit and 64-Bit Platforms](http://msdn.microsoft.com/en-us/library/dd941355.aspx). Not many people do have this installed. On request I can try to build one.

__Why Multibyte, isn't it deprecated for MFC?__

MAPI only supports unicode for a couple of functions. See [MAPI Character Sets](http://msdn.microsoft.com/en-us/library/ms530680.aspx)
