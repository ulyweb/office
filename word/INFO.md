Restoring the default Normal.dotm template in Microsoft Word on Windows 11 is a straightforward process. Here's how you can do it:

**Method 1: Rename the Existing Normal.dotm File (Recommended)**

This is the safest and most common method because it allows you to recover your customized template later if you change your mind.

1.  **Close all instances of Microsoft Word.** Make sure no Word documents are open.

2.  **Open File Explorer.** You can do this by pressing the `Windows Key + E`.

3.  **Navigate to the correct folder.** The Normal.dotm file is located in a hidden folder. To get there quickly, you can paste the following path into the address bar of File Explorer and press Enter:

    `%appdata%\Microsoft\Templates`

    Alternatively, you can manually navigate there:
    * `C:\Users\<YourUsername>\AppData\Roaming\Microsoft\Templates`
    * **Note:** The `AppData` folder is hidden. If you can't see it, go to the "View" tab in File Explorer, click "Show," and select "Hidden items."

4.  **Find the `Normal.dotm` file.**

5.  **Rename it.** Right-click the file and select "Rename." Change the name to something like `Normal.dotm.old` or `Normal_backup.dotm`.

6.  **Open Microsoft Word.** When you open Word, it will not find the `Normal.dotm` file and will automatically create a brand new, default version of the template.

**Method 2: Delete the Existing Normal.dotm File**

This method is faster, but a little less forgiving as the old file is permanently deleted. Only do this if you are certain you don't need your old customized template.

1.  **Close all instances of Microsoft Word.**

2.  **Open File Explorer** and navigate to the `Templates` folder using the path from Method 1: `%appdata%\Microsoft\Templates`

3.  **Find the `Normal.dotm` file.**

4.  **Delete it.** Right-click the file and select "Delete."

5.  **Open Microsoft Word.** Word will automatically generate a new, default `Normal.dotm` file to replace the one you deleted.

**What is the Normal.dotm template?**

The `Normal.dotm` file is the master template for all new, blank documents in Microsoft Word. It contains default settings for:
* Paragraph styles (e.g., Normal, Heading 1, etc.)
* Page layout (margins, paper size, orientation)
* Fonts
* Macros and custom toolbars

When you make changes to a new blank document and save the changes to the "Normal" template, you are modifying this file. Restoring it is useful if you accidentally made changes that you want to undo or if the file becomes corrupted.
