# Aspose.PDF for Excel Add-in

Export Excel sheets & ranges into editable PDFs with our Excel Add-in.

Our PDF exporter tool is intended for users and system administrators who want more flexibility when working with PDFs.

Work on your documents within the Excel application, then use the Aspose.PDF Add-in to convert these spreadsheets into PDFs in seconds.
You can also compress your file before conversion to ensure your PDF is in an optimal file size for easy shareability.

## 📓 Features

When this add-in is used, you can:

- Convert your XLSX files into PDFs in seconds.
- Convert selected ranges into PDFs.

![Add-in](./assets/addin01.png)

## 🛠 How to install

1. Go to the [Releases](https://github.com/aspose-pdf/aspose-pdf-for-excel-addin/releases) folder.
2. Download the latest version.
3. Run the downloaded file and follow the on-screen instructions to install the add-in.

## 🚀 Activatiton

**Warning!** The current add-in version supports the desktop version of Excel.

### Create Trusted Add-in Catalogs using Shared folders / Windows 11

1. Share a folder:
    1. In `File Explorer`, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.
    1. Open the context menu for the folder you want to use as your shared folder catalog (for example, right-click the folder) and choose `Properties`.
    1. In the `Properties` dialog window, open the `Sharing` tab and choose the `Share` button.
    1. Within the `Network access` dialog window, add yourself and any other users and/or groups with whom you want to share your add-in. The folder must have at least `Read/Write` permission. After choosing people to share with, choose the `Share` button.
    1. When you see the Your folder is shared confirmation, note the full network path displayed immediately following the folder name. (You'll need to enter this value as the Catalog Url when you specify the shared folder as a trusted catalog, as described in the next section of this article.) Choose the `Done` button to close the `Network Access` dialog window.
    1. Choose the `Close` button to close the `Properties` dialog window.

1. Specify the shared folder as a trusted catalog:
    1. Open a new document in Excel.
    1. Choose the `File` tab, and then choose `Options`.
    1. Choose `Trust Center` and then the `Trust Center Settings` button.
    ![Option dialog](./assets/image01.png)
    1. Choose `Trusted Add-in Catalogs`.
    1. In the `Catalog Url` box, enter the full network path to the folder that you shared previously. If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's Properties dialog window.
    1. After entering the folder's full network path into the `Catalog Url` box, choose the `Add catalog button`.
    1. Select the `Show in Menu` check box for the newly added item, and then click the `OK` button to close the Trust Center dialog window.

    ![Trusted Add-in Catalogs](./assets/image02.png)

    1. Choose the `OK` button to close the `Options` dialog window.
    1. Close and reopen the Excel application so your changes will occur.

1. Activate the add-in:
    1. Put the manifest XML file (the one you have edited and saved) into the shared folder.
    1. Select `Home > Add-ins` from the Excel ribbon, then select `Get Add-ins`.
    1. Choose `SHARED FOLDER` at the top of the Office Add-ins dialog box.
    ![Shared folder](./assets/image03.png)
    1. Select the Aspose.PDF for Excel add-in and choose `Add to insert the add-in`.
    1. Verify that your add-in is installed. The new button Aspose.PDF for Excel should appear on the Home ribbon.
    1. For more information regarding add-in installation in Excel software on Windows, please refer to [Sideload Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins#share-a-folder) for testing from a network share.

## Support

- 🙏 [Aspose.PDF for .NET Free Support Forum](https://forum.aspose.com/c/pdf/10)
- 💡 [Aspose.PDF for .NET Paid Support Helpdesk](https://helpdesk.aspose.com/)
