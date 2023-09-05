# Stractivity-Excelizer

![logo-filled](https://github.com/datacrat/Stractivity-Excelizer/assets/62040535/052effb4-c17c-48ad-b3ea-e9dab7e443b9)

## What is Stractivity Excelizer?

Stractivity Excelizer is an Excel Add-in to import your all Strava activities into a worksheet. You can sort, search and analyze the activities as you like by using Excel after importing the activities.

## Why use Stractivity Excelizer?

Strava is an excellent service to record and measure your fitness activities. It provides great summary views such as training log, power curve, fitness & freshness (trend of fitness strength), and so on, however, you cannot customize those views provided in Strava's Web site or in Strava's mobile apps for your special purposes. For example, if you would like to see your yearly training trend in the latest 10 years, or if you would like to see the top 10 heaviest activities ever, how would you do? Stractivity Excelizer enables you to perform search, analyze your activities in arbitrary manner in Excel.

## How to use Stractivity Excelizer?

### Prerequisites

- Either version of Microsoft Excel listed below
    - Office on Windows - Microsoft 365 subscription
    - Office on Windows - retail perpetual Office 2016 (Version 1509) or later
    - Office on Windows - volume-licensed perpetual Office 2016 (Version 1509) or later
    - Office on Mac - 15.20 or later
    - Office on the web
 
- Your Strava account

### Steps

1. (First time only) Setup "My API Application" in Strava
  - In a Web browser, open [My API Application](https://www.strava.com/settings/api)
  - You will see "Create an App" form. Fill each form as below then press "Create".
    | FORM | VALUE |
    | ---- | ---- |
    | Application Name | `Stractivity Excelizer` |
    | Category | `Data Importer` |
    | Club | - Blank - |
    | Website | `https://datacrat.github.io/Stractivity-Excelizer/` |
    | Application Description | - you can leave this blank - |
    | Authorization Callback Domain | `datacrat.github.io` |

#### If you have already registered an API Application
Strava provides ONLY ONE API Application entry for each user and you cannot create another application. In this case, you have to edit the existing API Application as above for Stractivity Excelizer to work. Before editing the existing API Application, take notes of the original so you can restore API Application setup for the original application.

2. Load Stractivity Excelizer Add-in - See [How to Load Stractivity Excelizer Add-in](#how-to-load-stractivity-excelizer-add-in)

4. Choose **Show Stractivity Excelizer** button in the home ribbon and a task pane appears

5. Choose **Connect with Strava** button in the task pane and a popup windows appears

6. Put Client ID into the popup. You can copy it from [My API Application](https://www.strava.com/settings/api). Then choose **Send**

7. If the popup shows Strava Login dialog, perform login to Strava in the dialog

8. The popup shows Authorize Screen. Make sure **View data about your private activities** is checked and then choose **Authorize**

9. Put Client Secret into the box in the popup. You can copy and paste it from [My API Application](https://www.strava.com/settings/api). Then choose **Send**

10. Press **Retrieve Activities** button. Depending on the number of your activities, retrieving takes a couple of minutes. You see the progress by page counter in _Read X pages._

11. Your activities will be listed in a table "MyActivities" in the current worksheet. Now you can sort, search and analyze your activities as you like. Enjoy!

### How to Load Stractivity Excelizer Add-in

At this moment, Stractivity Excelizer Add-in is not listed in Office Store and you will have to take following steps, instead of loading the Add-in from Office Store.

<details>
  <summary>Load Stractivity Excelizer Add-in in Office on Windows</summary>

  #### Load Stractivity Excelizer Add-in in Office on Windows
  1. Download [manifest.xml](https://datacrat.github.io/Stractivity-Excelizer/dist/manifest.xml)
  2. Follow the instructions in [Sideload Office Add-ins for testing from a network share](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins?view=common-js-preview)
     ##### Example:
       1. Locate `manifest.xml` in the folder `C:\Users\<user>\Documents\Stractivity-Excelizer\`
       2. Share the folder `C:\Users\<user>\Documents\Stractivity-Excelizer\`. You need to configure the share so you have Read/Write permission.
       3. Take note of the full network path of the shared folder. Typically, it would be `\\<Computer Name>\Stractivity-Excelizer`
       4. Open Excel
       5. Choose the **File** tab and then choose **Options**
       6. Choose **Trust Center** and then choose the **Trust Center Settings** button
       7. Choose **Trusted Add-in Catalogs**
       8. In the Catalog Url box, exter the full network path you noted previously - typically  it would be `\\<Computer Name>\Stractivity-Excelizer`
       9. Check **Show in Menu** and then choose **OK** button to close the **Trust Center** dialog window
       10. Choose **OK** button to close the **Options** dialog window
       11. Close and reopen Excel so the changes will take effect
       12. Load Stractivity Excelizer Add-in from **Office Add-ins** window. How to open **Office Add-ins** window is slightly different by Office versions. You may have to colsult **Help** of Excel. In case of version 2307, you can locate **Add-ins** button in the Home ribbon. Choose it and then choose **More Add-ins** so **Office Add-ins** window opens.
       13. In **Office Add-ins** window, choose **SHARED FOLDER** tab and you will find Stractivity Excelizer. Choose it and then choose **Add** so **Show Stractivity Excelizer** will be located in the home ribbon
</details>

<details>
  <summary>Load Stractivity Excelizer Add-in in Office on Mac</summary>

  #### Load Stractivity Excelizer Add-in in Office on Mac
  1. Download [manifest.xml](https://datacrat.github.io/Stractivity-Excelizer/dist/manifest.xml)
  2. Follow the instructions in [Sideload Office Add-ins on Mac for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac?view=common-js-preview)
     ##### Example:
       1. Locate `manifest.xml` in the folder `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
       2. Open Excel and then open a document
       3. On the ribbon, choose the **Insert** tab and then select **My Add-ins** (dropdown menu) in the **Add-ins** group. On the dropdown menu, you can choose Stractivity Excelizer Addin.
       4. **Show Stractivity Excelizer** will be located in the home ribbon
</details>

<details>
  <summary>Load Stractivity Excelizer Add-in in Office on the web</summary>

  #### Load Stractivity Excelizer Add-in in Office on the web
   1. Download [manifest.xml](https://datacrat.github.io/Stractivity-Excelizer/dist/manifest.xml)
   2. Follow the instructions in [Manually sideload an add-in to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing?view=common-js-preview#manually-sideload-an-add-in-to-office-on-the-web)
      ##### Example:
        1. Open Excel on the web and open a document
        2. Choose **Add-ins** and then choose **More Add-ins** so **Office Add-ins** window opens
        3. In **Office Add-ins**, choose **MY ADD-INS** tab so you will locate **Upload My Add-in**
        4. Select **Upload My Add-in** and upload `manifest.xml`
        5. **Show Stractivity Excelizer** will be located in the home ribbon
</details>
