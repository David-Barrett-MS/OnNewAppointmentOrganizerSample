# OnNewAppointmentOrganizerSample

This sample show how to develop an event based add-in in Visual Studio.

Note that you may need to update the Visual Studio add-in schema definition file as follows (the files that ship with Visual Studio do not contain the latest updates):

- Ensure Visual Studio is closed
- Obtain the latest xsd file from the [Exchange Protocol Documents](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-owemxml/8e722c85-eb78-438c-94a4-edac7e9c533a)
- Open Notepad or preferred text editor as Administrator
- Open file **C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Xml\Schemas\1033\MailAppVersionOverridesV1_1.xsd**
- Replace the contents of this file with the latest downloaded from the Protocol docs, then save
- Visual Studio will now be able to compile the sample [Event-based activation code](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch)
