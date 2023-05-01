# WordAddInDisableMacros

## Issue
Disabling VBA macros will throw a warning prompt to the user, and throw a COMException on the Documents.Open() method call.
The warning prompt shows even if the add-in is a bare-bones project that has no code added to it. The prompt does not show in Excel or PowerPoint, nor is the exception thrown in those add-ins.
Once the document is saved to SharePoint, using the native 'Compare' function in the ribbon shows the prompt several times; sometimes more than a dozen dialogs are shown.

Policy information: 

![image](https://user-images.githubusercontent.com/60905116/235055427-8b6abd4d-2766-48b8-a02d-9a10067e2560.png)


policy enabled via registry, if this key doesn't exist - create it:

![image](https://user-images.githubusercontent.com/60905116/235055446-1561c4a0-07c3-4844-9906-c5bb3f3ffe0e.png)


User warning prompt:

![image](https://user-images.githubusercontent.com/60905116/235055660-361ced0f-5dc2-4095-8d81-668c8f19bb87.png)


## Repro steps
* Clone repo
* Disable macros using either group policy or registry key
* Start debugging
* Open a new document

![image](https://user-images.githubusercontent.com/60905116/235056518-cfd3deb6-a8d4-4674-991e-5db588aa9a90.png)
* See the warning prompt, even though the project is bare and just created from the Word-VSTO-AddIn template
* Save the document to a SharePoint location
* Make changes to the document
* Save and compare using the native 'Compare' option

![image](https://user-images.githubusercontent.com/60905116/235388797-13b0cb44-436b-4b42-a641-68ec28a95d9e.png)

* Note the multiple warning prompts shown to the user

## Expected Outcomes
Answers to these questions:
1. Is there something we can do to open documents that won't trigger this exception, or get flagged by Word as using macros?
2. Why do the prompts and exceptions happen in Word but not in the other applications?
