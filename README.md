# MS Office 2019 ProPlus Installation and Activation Guide

## Installation

1. **Download MS Office 2019 ProPlus:**
   - Click this link to download: [MS Office 2019ProPlus.img](http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/enus/ProPlus2019Retail.img)

2. **Uninstall Existing MS Office Installation:**
   - **VERY IMPORTANT:** Please uninstall any existing old MS Office installation on your system.
   - **Steps:**
     - Go to **System** > **Apps** > **Apps & Features**.
     - Find the existing MS Office installation and uninstall it.

3. **Mount the Installer:**
   - After downloading the installer, right-click on it and select **Mount**.
   - Alternatively, right-click on it and select **Open with** > **Windows Explorer**.

4. **Run the Setup File:**
   - Right-click on the **Setup** file and select **Run as administrator**.

5. **Wait for Installation:**
   - The installation may take a while. Please be patient.

6. **Complete Installation:**
   - Upon successful installation, close the installer and proceed to MS Office Activation.

## MS Office Activation (Internet Connection Required)

1. **Open Command Prompt as Administrator:**
   - Type `CMD` in the search box.
   - Right-click on **Command Prompt** and select **Run as administrator**.

2. **Navigate to Office Directory:**
   - Copy and paste the following command into the Command Prompt, then press **Enter**:
     ```cmd
     cd /d %ProgramFiles%\Microsoft Office\Office16
     ```
   - If you receive an error message saying “The system cannot find the path specified”, use this command instead:
     ```cmd
     cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
     ```

3. **Install License Key:**
   - Copy and paste the following command into the Command Prompt, then press **Enter**:
     ```cmd
     for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
     ```

4. **Set KMS Port:**
   - Copy and paste the following command into the Command Prompt, then press **Enter**:
     ```cmd
     cscript ospp.vbs /setprt:1688
     ```

5. **Uninstall Old Product Key:**
   - Copy and paste the following command into the Command Prompt, then press **Enter**:
     ```cmd
     cscript ospp.vbs /unpkey:6MWKP >nul
     ```

6. **Install New Product Key:**
   - Copy and paste the following command into the Command Prompt, then press **Enter**:
     ```cmd
     cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
     ```

7. **Set KMS Host:**
   - Copy and paste the following command into the Command Prompt, then press **Enter**:
     ```cmd
     cscript ospp.vbs /sethst:kms8.msguides.com
     ```

8. **Activate Office:**
   - Copy and paste the following command into the Command Prompt, then press **Enter**:
     ```cmd
     cscript ospp.vbs /act
     ```

9. **Verify Activation:**
   - Close Command Prompt.
   - Open **Word** or **Excel**.
   - Go to **File** > **Account**.
   - Check if MS Office 2019 ProPlus is activated.

Enjoy your MS Office 2019 ProPlus!
