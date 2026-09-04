# invSys Excel Operator Setup Requirements

## Who this is for

This guide is for a person who will sign in to invSys and operate Receiving,
Production, Shipping, Boxing, or Viewer. It is not a developer setup guide.

## Required license and device

- A Windows 10 or Windows 11 computer.
- A licensed desktop installation of Microsoft Excel for Windows with VBA and
  Excel add-ins available. The recommended standard is **Microsoft 365 Apps for
  business or enterprise, desktop Excel**.
- Excel for the web, iPad, Android, and macOS Excel are not supported operator
  platforms for Release 1. invSys uses Windows VBA UserForms and Windows API
  resizing behavior.
- The user must be able to open `.xlsm`, `.xlsb`, and `.xlam` files. Their IT
  license/policy must not disable VBA or block Excel add-ins.

The operator does **not** need Visual Basic Editor access, Trust Center access,
GitHub access, a developer license, or permission to install software.

## IT or invSys administrator prerequisites

Before the operator opens Excel, an administrator must ensure:

1. The user has a valid Microsoft Office/Excel desktop license and is signed in
   to Office if their organization requires it.
2. VBA macros are permitted for the invSys managed add-ins. Do not use a global
   “Enable all macros” setting. Apply the organization’s approved Excel macro
   policy, code-signing policy, or managed trusted-location policy to invSys.
3. The user can reach the NAS through the organization’s approved network/VPN
   (currently Tailscale where applicable) and has Windows/SMB permissions for
   the warehouse roots they are allowed to use.
4. The invSys station updater has been installed for that Windows account. It
   obtains the verified five-package release from the NAS and registers the
   Operations and Admin startup add-ins. Operators must not browse to or
   manually register individual XLAM files.
5. A warehouse administrator has created the warehouse/station configuration
   and assigned the person an invSys user ID and applicable capabilities.

## First-use check

1. Close every Excel window.
2. Start desktop Excel normally. Allow the station updater to finish before
   opening invSys if Windows indicates it is running.
3. Confirm the **invSys** ribbon is visible. The expected operator add-in is
   **invSys.Operations**; **invSys.Admin** is present only when the account is
   authorized for Admin functions.
4. Use **Server Sign In** / **Send To** only to connect to an authorized NAS
   warehouse. Then use **invSys Sign In** with the assigned invSys credentials.
5. Confirm that only the allowed role buttons are enabled. A disabled role is a
   capability assignment issue, not an Excel license problem.

## Common Excel-policy failures

| Symptom | Likely owner | Correct resolution |
|---|---|---|
| invSys ribbon is missing | Station deployment / Excel add-in policy | Close Excel; have IT or an invSys admin verify the station updater and managed add-in registration. |
| Yellow or red macro-security warning | Office security policy | Have IT approve the managed invSys release through the organization’s normal macro policy. Do not lower global macro security. |
| “Macros have been disabled” | Office security policy | Same as above; the operator should not click around Trust Center settings. |
| Server connection fails | Network, VPN/Tailscale, or NAS permissions | Verify the user is on the approved network and has SMB rights to the intended NAS root. |
| invSys Sign In is unavailable after server login | Warehouse has not been selected | Use **Send To** to select an authorized warehouse, then sign in. |
| Role button is disabled after invSys Sign In | invSys authorization | Have an Admin review the user’s capability assignment for the selected warehouse/station. |

## Operator rules

- Do not manually copy, rename, edit, or browse-open invSys XLAM files.
- Do not enable untrusted macros or add a broad trusted location to work around
  a warning.
- Do not use personal cloud-sync folders for warehouse workbooks or invSys
  runtime files.
- Close Excel before a scheduled update, rollback, or support-directed add-in
  repair. The updater safely defers while Excel is open.

## Support information to provide

When requesting help, send the support person:

- Windows version and Excel desktop version/build;
- whether the invSys ribbon appears;
- the exact warning/error text or screenshot;
- whether the NAS/VPN connection works; and
- the selected warehouse and role (but never passwords or PINs).

This guide reflects Release 1 Architecture v4.11 and the NAS-only operational
deployment model. GitHub is developer-only.
