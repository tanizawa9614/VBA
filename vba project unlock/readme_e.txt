======================================================================
[Software name          ] VBA lock & unlock
[Version                ] Ver 1.64
[Kind of software       ] Freeware
[Runnable OS            ] Windows XP/Vista/7/8/10/11
[Necessary thing        ] Microsoft Excel
[Development language   ] Excel VBA
[Date of publication    ] 2023/01/09
[Copyright              ] Copyright(C) K.Hiwasa
[URL                    ] http://srcedit.pekori.jp
[E-mail                 ] srcedit@a.pekori.jp
[Development environment] Windows 11/Microsoft Excel 2021
[Condition of reprint   ] Contact required
======================================================================
[Introduction of software]
Lock / unlock and hide / restore VBA projects.
Locking can be special locks and shared locks.
Unlocking is usually password lock, special lock, and shared lock.
Hiding hides the module source.
Restoring restores the hidden module source.
Supports Excel, Word, PowerPoint and Access VBA projects.


[Update of this version]
- Change the form to modeless.


[Files]
vbalckunl.xlam : Body of tool.
readme.tx      : Japanese version of this file.
readme_e.txt   : This file.


[How to install]
Only execute vbalckunl.xlam.
Excel must be able to use macros.


[How to uninstall]
Delete the entire folder. Not using the registry.


[How to use]
- Lock
-- Press the lock button and specify the file to be locked.
-- If a shared lock(*) is possible, 
   a confirmation dialog for the shared lock will appear.
-- Create a locked file with "_Lck" added to the original file name
   in the same folder as the target file.

- Unlock
-- Press the unlock button and specify the file to be unlocked.
-- Create a unlocked file with "_Unl" added to the original file name
   in the same folder as the target file.

- Hide
-- Hide the standard module source.
-- Press the hide button and specify the file to be hidden.
-- Create a hidden file with "_Hid" added to the original file name
   in the same folder as the target file.

- Restore
-- Restore the hidden standard module source.
-- Press the restore button and specify the file to be restored.
-- Create a restored file with "_Rst" added to the original file name
   in the same folder as the target file.

* Shared lock
  As the workbook becomes shared, the original process may not work properly.
  Enforce a shared lock only if there is no problem with sharing.


[Disclaimer]
- I'm not responsible for any troubles caused by using this software.
  Please use at your own risk.
- Requests for improvements and bug reports are welcome,
  but I'm not sure if it's technically or policy-wise possible,
  and if it's possible, I'm not sure if it's possible soon.


[History]
2022/10/26  Version 1.63
- Improved unlocking accuracy.
- Improved restoring accuracy.

2022/08/22  Version 1.62
- Added note about hiding.

2022/01/15  Version 1.61
- Improved unlocking accuracy.

2022/01/09  Version 1.60
- Added English mode.

2022/01/08  Version 1.57
- Added non-restorable level 3 to Excel file hiding.
- Improved restoring accuracy.

2022/01/03  Version 1.56
- Added non-restorable level 2 to Excel file hiding.

2022/01/02  Version 1.55
- Improved locking accuracy.

2021/12/31  Version 1.54
- Improved convenience of unlocking.

2021/12/29  Version 1.53
- Improved locking accuracy.

2021/12/27  Version 1.52
- Improved restoring accuracy.
- Added that it is possible to unlock the camouflage source in the unlock description.

2021/11/09  Version 1.51
- Minor bug fixes.

2021/10/15  Version 1.50
- Supports Access VBA projects.

2021/10/13  Version 1.40
- Supports Word and PowerPoint VBA projects.

2021/10/06  Version 1.35
- Added non-restorable mode to hiding.

2021/10/05  Version 1.34
- Improved hiding accuracy.
- Improved restoring accuracy.

2021/09/29  Version 1.33
- Improved hiding accuracy.
- Minor correction.

2021/09/27  Version 1.32
- Improved restoring accuracy.
- Minor correction.

2021/09/22  Version 1.31
- Minor correction.

2021/09/21  Version 1.30
- Added hiding.
- Improved locking accuracy.
- Improved restoring accuracy.
- Change the file format to add-in.

2021/09/20  Version 1.20
- Added restoring.
- Improved unlocking accuracy.

2021/09/18  Version 1.13
- Improved unlocking accuracy.

2021/08/04  Version 1.12
- Improved unlocking accuracy.
- Minor correction.

2021/08/04  Version 1.11
- Minor correction.

2021/08/03  Version 1.10
- Added locking.
- Improved unlocking accuracy.

2021/08/02  Version 1.00
- New release.
