The WMFSDK idl was originally made into an interop dll with midl and tlbimp, then many changes were made to the interfaces to make the marshalling work correctly.  These changes were applied directly to the IL.  Therefore ManagedWM.il should be considered the source file.  The dll is rebuilt from ManageWM.il using ilasm.  "ilasm /DLL ManagedWM.il

The original version of the dll was built this way:

midl ManagedWM.idl /I "C:\Program Files\Microsoft Visual Studio .NET\Vc7\PlatformSDK\Include" /I D:\WMSDK\WMFSDK9\include
tlbimp ManagedWM.tlb /namespace:UW.CSE.ManagedWM
