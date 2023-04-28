# WordAddInDisableMacros

## Issue
Disabling VBA macros will throw a warning prompt to the user, and throw a COMException on the Documents.Open() method call.
The warning prompt shows even if the add-in is a bare-bones project that has no code added to it.

Policy information: 

![image](https://user-images.githubusercontent.com/60905116/235055427-8b6abd4d-2766-48b8-a02d-9a10067e2560.png)


policy enabled via registry:

![image](https://user-images.githubusercontent.com/60905116/235055446-1561c4a0-07c3-4844-9906-c5bb3f3ffe0e.png)


User warning prompt:

![image](https://user-images.githubusercontent.com/60905116/235055660-361ced0f-5dc2-4095-8d81-668c8f19bb87.png)


## Repro steps
* Clone repo
* Start debugging
* Open a new document

![image](https://user-images.githubusercontent.com/60905116/235056518-cfd3deb6-a8d4-4674-991e-5db588aa9a90.png)
* See the warning prompt, even though the project is bare and just created from the Word-VSTO-AddIn template
