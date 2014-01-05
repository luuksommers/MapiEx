# Extended Mapi for C# #

### Original code by Noel Dillabough ###
http://www.codeproject.com/Articles/10881/MAPIEx-Extended-MAPI-Wrapper

### Setup ###
1. Visual Studio 2013 Professional or higher for MFC headers
2. Download [Visual Studio 2013 C++ Multi Byte Support](http://go.microsoft.com/?linkid=9832071)

### Deployment ###
This library uses the [Visual C++ 2010 redistributable](http://www.microsoft.com/en-us/download/details.aspx?id=5555)

### FAQ ###
__Does this support x64 MAPI__
No, MAPI only supports x64 when an x64 office is installed. See [Building MAPI Applications on 32-Bit and 64-Bit Platforms](http://msdn.microsoft.com/en-us/library/dd941355.aspx). Not many people do have this installed. On request I can try to build one.

__Why Multibyte, isn't it deprecated for MFC?__
MAPI only supports unicode for a couple of functions. See [MAPI Character Sets](http://msdn.microsoft.com/en-us/library/ms530680.aspx)
