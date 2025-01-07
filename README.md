# IT Helpdesk: Outlook Won't Open - How To Fix it Easy Steps

## Issue Overview
This document provides a step-by-step guide, similar to a video tutorial, on how I resolved the issue where Microsoft Outlook wouldn't open. The issue was affecting multiple users, and the solution addresses common causes for this problem. Follow along with the steps to fix it.

## Steps Taken to Resolve the Issue

### Step 1: Run Outlook in Safe Mode and Deactivate Add-ins
- **Open Outlook in Safe Mode:** Press **Ctrl** and hold it while launching Outlook, or type `outlook.exe /safe` in the Run dialog (Win+R). This opens Outlook without any add-ins or customizations that might be causing the issue.
- **Deactivate Add-ins:** If Outlook opens successfully in Safe Mode, it indicates that one of the add-ins or customizations may be causing the problem. To deactivate add-ins:
  - Go to **File > Options > Add-ins**.
  - At the bottom of the window, select **Go** next to Manage: COM Add-ins.
  - Uncheck all add-ins and click **OK**.
  - Restart Outlook in normal mode to check if it opens without issues. If the problem is resolved, re-enable the add-ins one by one to identify the culprit.

### Step 2: Create a New Outlook Profile
- **Open Control Panel:** Navigate to **Control Panel > Mail (Microsoft Outlook)**.
- **Show Profiles:** In the Mail Setup window, click **Show Profiles**.
- **Add New Profile:** Click **Add** to create a new profile. Enter a name for the new profile and follow the prompts to configure your email account.
- **Set Default Profile:** After creating the new profile, set it as the default profile by selecting **Always use this profile** and choosing the newly created profile from the dropdown menu.
- **Test New Profile:** Open Outlook using the new profile to check if the issue is resolved. If the problem is fixed, the original profile may have been corrupted.

### Step 3: Run the Scan.pst Tool
- **Locate Scan.pst:** The **scan.pst** tool (Inbox Repair Tool) is usually located in the following path: 
  - **C:\Program Files (x86)\Microsoft Office\root\OfficeXX** (where **XX** corresponds to your Office version number, e.g., Office16 for Office 2016).
- **Run the Tool:** Launch the **scan.pst** tool by double-clicking the **scanpst.exe** file.
- **Repair PST File:** In the tool, select the **PST** file that needs to be repaired (usually found in **C:\Users\<username>\AppData\Local\Microsoft\Outlook**). Click **Start** to begin scanning for errors in the file.
- **Fix Errors:** If errors are found, the tool will prompt you to repair the file. Click **Repair** to proceed. Once the repair is complete, restart Outlook and check if it opens properly.

### Step 4: Repair Outlook via Control Panel
- **Open Control Panel:** Go to the **Control Panel** on the user’s computer.
- **Programs and Features:** Click on **Programs** and then **Programs and Features**.
- **Repair Office Installation:** Find **Microsoft Office** in the list of installed programs, right-click on it, and select **Change**. Choose the **Repair** option and follow the on-screen instructions to fix any issues with the Outlook installation.

## Conclusion
By following these easy steps, I was able to resolve the issue of Outlook not opening for the affected user. Regular maintenance and monitoring can help prevent similar issues in the future.

## Tools Used
- **Microsoft Outlook**
- **CMD (Command Prompt)**
- **Task Manager**
- **Control Panel**
- **Scan.pst Tool (Inbox Repair)**
