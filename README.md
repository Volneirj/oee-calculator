# OEE Calculator

OEE Calculator is a tool to help you to indentify if your process is in a good shape,
the idea is make something simple to calculate your process availability, performance and quality
with this information calculate the OEE factor where will identify if your process have a good Overall equipment effectivess. (OEE)


[OEE calculator](https://oee-calculator-f9cea04d2a91.herokuapp.com/)

## Project Overview


### Purpose:

As part of a course project, this solution serves as an educational resource, demonstrating the application of python development skills to create some kind application or solution. Through the calculator, we aim to explore GoogleCloud API/GoogleSheets/Python/ features, creating an application with interact with the user and cloud services.

### Existing Features

#### Main Menu

- The main menu show the options available for the user.
- The options can be selected usin the number 1 to 4.

![Main Menu](docs/readme_images/main-menu.jpg)

#### Add new report

- On this option the user will answer the questions where will be processed and transfered for a file on
google drive.

![New Report](docs/readme_images/new-report.jpg)

#### Show All Reports

- On this option the user can acess all the saving data on googlesheet.

![Show Reports](docs/readme_images/show-reports.jpg)

#### Load OEE by date

- On this option the user can load the OEE calculated by date.


![OEE by date](docs/readme_images/oee-by-date.jpg)

#### Exit

- This option will close the application and show the credits.

![Exit](docs/readme_images/exit.jpg)

## Technical Design

### Overal Equipment Effectivenss (OEE) explanation

- The OEE is one of the tools used for process improvement, most of time used by a 6sigma Green belt or above.
- The idea is gather the process information and calculate the quality, availability and performance.
- With this three factors calculate the Overall Equipment Effectiveness.
- This data is used to understanding of your process and find oportunities to improvements.

- The data needed to calculate the OEE are:

1. Shift length in minutes - Ex. 480 minutes (8 hours)
2. Short and long break in minutes - Ex. 30 minutes (Lunch break)
3. Down time in minutes - Ex. 45 minutes (Machine Stopped work by overheating)
4. Ideal run rate in unit/parts per minute - Ex. 60 units (The expectation is assembly 60 box per minute)
5. Total parts done per shift in unit - Ex. 19000 units (Box assembled per day shifth length)
6. Rejected parts per shift in units - Ex. 500  (500 boxes were damaged so discarted)

- With this information is necessary to calculate the auxiliar variables which are:

1. Planned Production Time in minutes result from: 
    **Planned production time** = Shift Length - Breaks 
2. Operation time in minutes result from:
    **Operation time** = planned production time - down time 
3. Good pieces in unit result from:
    **Good pieces** = total pieces - rejected pieces

- Them using the auxiliar variables together with the production data is calculate the factors:

1. Availability in %:
    **Availability** = (operating time/production time)*100
2. Performance in %:
    **Performance** = ((total piece/operantion time)/Ideal run rate)*100
3. Quality in %:
    **Quality** = (good pieces/total pieces)*100 
4. Overall OEE = availability * performance * quality


### Flowchart

- For the development and logic decision maker has been draw a flowchart where show all user and system interaction flow.

![Logic Flow chart](docs/readme_images/flowchart.jpg)

## Technologies

- **Python:** Backend Development
- **Visual Studio Code:** Local IDE.
- **GitHub:** Source code hosting and deployment.
- **Git:** Version control.
- **Heroku:** Cloud Deployment.
- **Google Cloud Services** APIs to access and manipulate cloud based files
- **Google Drive** Store data in a googlesheet 
- **Google Sheet** Data Storage
- **Miro** Flowchart and design thinking

## Libraries
- **googleauth** Used to provide access to the application to interact with my google sheet.
- **gspread** Used to access google sheets document throughout the application, to access and edit data.
- **datetime** Used to validate datetime input
- **os** Used to interact with the operation system

## Google Cloud

- To the data management strategy, use the Google Cloud robust APIs to store and manipulate data effeciently. Specifically, rely on two APIs provided by GoogleCloud:
the Google Drive API and the Google Sheets API. These APIs empower us to seamlessly interact with data stored on Google Drive and within Google Sheets programmatically.
To ensure the utmost privacy and security to the data, has been implemented a sophisticated approach using service accounts.

### Google Drive API

- The Google Drive API allows developers to interact with files and folders stored on Google Drive programmatically. It provides methods for uploading, downloading, searching, and modifying files, as well as managing permissions and metadata.

![Google Drive API](docs/readme_images/google-drive-api.jpg)

### Google Sheets API

- The Google Sheets API enables developers to read, write, and manipulate Google Sheets data using code. It allows for tasks such as creating new sheets, updating existing ones, inserting and deleting rows and columns, and formatting cells.

![Google Sheet API](docs/readme_images/google-sheet-api.jpg)

### Service Account

- A service account is a special type of Google account that belongs to your application or a virtual machine (VM), instead of an individual user. It's typically used when the application needs to access Google Cloud services programmatically without user interaction. Service accounts are associated with cryptographic key pairs, which can be used to authenticate API requests.

![Service Account](docs/readme_images/service-account.jpg)

### Secure Management of Service Accounts

- In the setup, security is prioritized through the management of user service accounts via JSON files. Precautions have been taken to ensure the safety of these private keys. Specifically, the JSON file containing the service account credentials is added to the .gitignore file. This step prevents accidental commits of sensitive information to version control repositories, effectively safeguarding the private keys from unauthorized access.

## Testing and Fixing Bugs

### Debugging 

  - For debugging the python code has been used `print` together with debugging tools and the console logs to help identify bugs, loops, and condition interactions.. 

  **Debugging Tools**

   - [CI Python Linter - Code institute](https://pep8ci.herokuapp.com/)  
   - [OpenAI Chat](https://chat.openai.com/)
   - [Perplexity AI](https://www.perplexity.ai/) 


### Aplication interaction tests

After deployment, a batch of tests has been conducted, and the results are shown below.

### Test Table

#### Main program

| **Feature**            | **Expected Results**                                              | **Testing Performed**                                       | **Result** |
|:-----------------------:|:------------------------------------------------------------------:|:------------------------------------------------------------:|:----------:|
| **Run Program**           | Run program without errors                                 | Run the program                           | Pass       |
| **Main Menu**  | Accept only numbers 1 to 4                                               | Try different type of characters and different numbers.            | Pass       |
| **1 - Add New report** | Ask user the daily report information starting from the date        | Enter option 1 using the main menu.              | Pass       |
| **2 - Load report** | Connect to google drive, extract data and show to the user in a table       | Enter option 2 using the main menu.              | Pass       |
| **3 - Load OEE by date** | Connect to google drive, extract data and show oee for selected date     | Enter option 3 using the main menu.              | Pass       |
| **4 - Exit** | Close the application and show credits   | Enter option 4 using the main menu.              | Pass       |


#### 1 - Add New report

| **Feature**            | **Expected Results**                                              | **Testing Performed**                                       | **Result** |
|:-----------------------:|:------------------------------------------------------------------:|:------------------------------------------------------------:|:----------:|
| **Date** | Only accept the dd/mm/yyyy date format              | Tried different date format, character types, check return "q".                | Pass       |
| **Name** | Only accept letters and minimum 3, not only blank characters |Tested different type of characters, empty spaces and 0-2 letters only, check return "q".                | Pass       |
| **Shifth Length** | Only accept integer|Tested different type of characters, empty spaces, check return "q".            | Pass       |
| **Short Breaks** | Only accept integer|Tested different type of characters, empty spaces, check return "q".         | Pass       |
| **Meal Breaks** | Only accept integer|Tested different type of characters, empty spaces, check return "q".            | Pass       |
| **Machine Shutdown** | Only accept integer|Tested different type of characters, empty spaces, check return "q".            | Pass       |
| **Ideal Run Rate** | Only accept integer|Tested different type of characters, empty spaces, check return "q".           | Pass       |
| **Total Processed Pieces** | Only accept integer|Tested different type of characters, empty spaces, check return "q".           | Pass       |
| **Total Rejected Pieces** | Only accept integer|Tested different type of characters, empty spaces, check return "q".           | Pass       |
| **Show User all input data with Headers** | Print a dictionary with (Header: User input)|Add all requested data and check the output data       | Pass       |
| **Ask the user if all data is correct** |If user write **yes** start to export the data |Write yes on prompt and check the result      | Pass       |
| **Ask the user if all data is correct** |If user write **no** restart the data colection process |Write no on prompt and check the result      | Pass       |

#### 1 - Add New report - Data manipulation

| **Feature**            | **Expected Results**                                              | **Testing Performed**                                       | **Result** |
|:-----------------------:|:------------------------------------------------------------------:|:------------------------------------------------------------:|:----------:|
| **Update Report Sheet** | Update worksheet **report** on google sheet located on google drive       | Check if the data has been exported to a new line on worksheet| Pass       |
| **Calculate Auxiliar Variables** | Calculate auxiliar variables   | Print message with the data before export| Pass       |
| **Update variables Sheet** | Update worksheet **variables**  on google sheet located on google drive  | Check if the data has been exported to a new line on worksheet| Pass |
| **Calculate OEE** | Convert and Calculate OEE  | Print message with the data before export| Pass       |
| **Update OEE Sheet** | Update worksheet **oee-factor** on google sheet located on google drive    | Check if the data has been exported to a new line on worksheet| Pass |
| **Print Results** | Print a text showing the Overall OEE results      | Check if the printed data is correct and in the rigth format| Pass |

#### 2 - Load all reports - Data manipulation

| **Feature**            | **Expected Results**                                              | **Testing Performed**                                       | **Result** |
|:-----------------------:|:------------------------------------------------------------------:|:------------------------------------------------------------:|:----------:|
| **Print all reports** | Load worksheet **report** on google sheet located on google drive  | Select option and verify if the data is correct and proper showed| Pass       |

#### 3 - Load OEE by date - Data manipulation

| **Feature**            | **Expected Results**                                              | **Testing Performed**                                       | **Result** |
|:-----------------------:|:------------------------------------------------------------------:|:------------------------------------------------------------:|:----------:|
| **Date** | Only accept the dd/mm/yyyy date format              | Tried different date format, character types, check return "q".                   | Pass       |
| **Date Not available** |Print no date available for dd/mm/yyyy         | Tried different which is not on googlesheet          | Pass       |
| **Show OEE** | If the date is correct, extract data and print the OEE result for the selected date| Check if the results are correct and showing in the right format| Pass       |


### Validator Testing

#### Python
- Using the [CI Python Linter - Code institute](https://pep8ci.herokuapp.com/) inspect and validate the python code, the image bellow is showing the result.

#### Before validation

![Python code before validation](docs/readme_images/ci-python-code.jpg)

#### After validation

![Python code after validation](docs/readme_images/ci-python-code-noerror.jpg)

## Deployment

### Github Forking

- **Forking the GitHub Repository**
  If you want to make changes to your repository without affecting it, you can make a copy of it by 'Forking' it. This ensures your original repository remains unchanged.

  1. Find the relevant GitHub repository
  2. In the top right corner of the page, click the Fork button (under your account)
  3. Your repository has now been 'Forked' and you have a copy to work on

- **Cloning the GitHub Repository**
  Cloning your repository will allow you to download a local version of the repository to be worked on. Cloning can also be a great way to backup your work.

  1. Find the relevant GitHub repository
  2. Press the arrow on the Code button
  3. Copy the link that is shown in the drop-down
  4. Open the terminal 
  5. Move to the folder you want clone it
  6. In the terminal type 'git clone' & then paste the link you copied in GitHub
  7. Press enter and your local clone will be created.
- Live link: [GitHub Link](https://https://github.com/Volneirj/oee-calculator)

### Heroku Deployment

- **Creating Requirements.txt**
- To heroku be able to install the required dependencies is necessary to create the file where will be listed what is needed to run the project.

1. Create a file requirements.txt.
2. Run the command: pip3 freeze >requirements.txt.
3. Check if the file has been updated like the image bellow.

![Requirements.txt](docs/readme_images/requirements.jpg)

- **Creating an Application with Heroku**
- To be able to  deploy and run the application on heroku plataform, is necessary follow a few steps:

1. Login or create an account on Heroku website.
2. Click on create a new app.

![New app](docs/readme_images/create-new-app.jpg)

3. After create the new app, you need to configure the settings.
4. The first setting which need to be done is add your CREDS.json info to the plataform so it will be able to access the googlecloud service account.

![Settings CREDS](docs/readme_images/creds.jpg)

5. Add the Buildpacks necessary to run the application, in this case python and nodejs in this sequence.

![Buildpacks](docs/readme_images/buildpacks.jpg)

6. After done the settings we move to the deploy tab where we will configure the deployment setup.
7. Connect your Github.
8. Select your repository on github.
9. Connect to the repository.

![Github](docs/readme_images/github.jpg)

10. After all setting above been done you can select to deploy automatic or manual.

![Deploy](docs/readme_images/deploy.jpg)

11. After press to deploy your project if all settings are working you should see it building the application.

![Project being Deployed](docs/readme_images/deployed.jpg)

12. After all steps of deployment will show a button View, where you can click to open a new tab with the application.

![Deploy done](docs/readme_images/done.jpg)


## Credits


1. **Base Code reference** 

   - [Code institute Love Sandwiches Walkthrough](https://learn.codeinstitute.net/)

