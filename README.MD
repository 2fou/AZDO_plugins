# Azure DevOps Extension: Questionnaire Form

This project is an Azure DevOps extension that provides a questionnaire form for work items. It enables users to answer a series of questions specific to each work item and store the responses directly in a custom field of the work item.

https://marketplace.visualstudio.com/items?itemName=AgileCoachingSolutionsFeatureAddon.feature-livrable-extension&targetId=9ed4a67e-0bf4-48ce-9276-8750bc557cd4&utm_source=vstsproduct&utm_medium=ExtHubManageList


## Features

- Load predefined questions from Azure DevOps Extension Data Service.
- Allow (Developers/IT Sec etc) to provide answers and optional links related to each  question.
- Store responses in a custom field associated with the work item.
- Ensure persistence and consistency of responses across sessions.

## Requirements

- Azure DevOps account
- Access to Azure DevOps Services for installing extensions
- Node.js and npm (for building the project)

## Setup and Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/your-username/azure-devops-extension-questionnaire.git
   cd azure-devops-extension-questionnaire
   ```

2. **Install Dependencies:**

   Run the following command to install all necessary dependencies:

   ```bash
   npm install
   npm install -g tfx-cli
   ```

3. **Build the Project:**

   Compile the extension using:

   ```bash
   npm run compile
   ```

   

4. **Package the Extension:**

   Use `tfx-cli` to package the extension, ensuring you have it installed previously:

   ```bash
   tfx extension publish --manifest-globs azure-devops-extension.json --overrides-file configs/dev.json --token xxxxxxxxxx
   ```

5. **Publish the Extension:**

   Login to Azure DevOps using `tfx` and publish the packaged extension. Ensure you replace placeholders with your publisher ID and other necessary identification.

   ```bash
   tfx extension publish --publisher your-publisher-id
   ```

6. **Add the Extension to Your Organization:**

   Once published, add the extension to your organization in Azure DevOps through the Marketplace.

## Setup

- Setup a new custom field "AnswersField" (Will hold the questionary answers) & "RoleAssignmentsField" (Will hold the role & responsabilities for the RACI table) in a process, this field can be hidden from the work item type.
      <img src="/src/ScreenShot/image-1.png" alt="Process" />
      ![Export](/dist/ScreenShot/image-1.png)
 In the below sample I have also create 3 new page which will use the ad-hoc contribution.
       <img src="/src/ScreenShot/image-2.png" alt="New custom Field" />
       ![Export](/dist/ScreenShot/image-2.png)
 For each page, select the corresponding contribution:
 - "Deliverable Control is the result of the Questionary for the Feature"
 - "Progress Indicator is a widget to show a pregress bar"
        <img src="/src/ScreenShot/image-3.png" alt="Contributions" />
        ![Export](/dist/ScreenShot/image-3.png)
 - "RACI Control" is used to assign roles and responsabilities to each deliveries
        <img src="/src/ScreenShot/image-4.png" alt="RACI" /> 
        ![Export](/dist/ScreenShot/image-4.png)
 - Define the deliverables through project/setting/extension
        <img src="/src/ScreenShot/2025-02-27 14_09_27-Clipboard.png" alt="Deliverables" /> 
        ![Export](/dist/ScreenShot/2025-02-27 14_09_27-Clipboard.png)
 - Define the Questionary through project/setting/extension
        <img src="/src/ScreenShot/Set you questionary.png" alt="Questionary" /> 
        ![Export](/dist/ScreenShot/Set you questionary.png)
 - Define the Questionary and assign deliverables
        <img src="/src/ScreenShot/AssignDeliverables.png" alt="Questionary & Assign deliverables" /> 
        ![Export](/dist/ScreenShot/AssignDeliverables.png)

- Set the contribution "work item setup"
        <img src="/src/ScreenShot/image-6.png" alt="Workitem Setup" /> 
        ![Export](/dist/ScreenShot/image-6.png)
- Role & responsability
        <img src="/src/ScreenShot/image-7.png" alt="Role Setup" /> 
        ![Export](/dist/ScreenShot/image-7.png)
- Navigate to a work item Epic in your Azure DevOps project.
        <img src="/src/ScreenShot/image-8.png" alt="Navigate to an Epic" /> 
        ![Export](/dist/ScreenShot/image-8.png)
- use the "create hierarchy" button
        <img src="/src/ScreenShot/image-9.png" alt="Create hierarchy" /> 
        ![Export](/dist/ScreenShot/image-9.png)
- Click on the new child "Feature" work Item
- Use the integrated tab to fill out the questionnaire.
        <img src="/src/ScreenShot/image-10.png" alt="Fill out questionary" /> 
        ![Export](/dist/ScreenShot/image-10.png)
- Fill out the questionnaire as you need.
        <img src="/src/ScreenShot/Risk Profile Matrice.png" alt="Fill out questionary" /> 
        ![Export](/dist/ScreenShot/Risk Profile Matrice.png)
- Once you have previously set the roles, you can start to set the RACI
        <img src="/src/ScreenShot/image-13.png" alt="Set the Roles" /> 
        ![Export](/dist/ScreenShot/image-13.png)
- The RACI TABLE will show the settings throught a RACI table
        <img src="/src/ScreenShot/image-14.png" alt="See the Raci table" /> 
        ![Export](/dist/ScreenShot/image-14.png)

## Development

- Make use of `azure-devops-extension-sdk` to interface with the Azure DevOps environment.
- Utilize `react` and `azure-devops-ui` components for the UI development.

## Debug

- In order to have your extension looking for a "local" resource, insure to have the entry
   - 
  "baseUri": "https://localhost:3000",
  "id": "feature-livrable-extension",
  "public": false,
  "version": "1.0.173",
 - This will indicate AzureDevOps extension where the resources are located.

- npm run compile:dev
- At least an initial deployment is required:
   - tfx extension publish --manifest-globs azure-devops-extension.json --overrides-file configs/dev.json  --token xxxx
- npx webpack serve
- From AzureDev you resources are now local to your development environment
- https://localhost:3000/dist/TestConfigurationPage/TestConfigurationPage.html

## Run (classic)

- npm run compile
- Update the Extension version
- tfx extension publish --manifest-globs azure-devops-extension.json --overrides-file configs/dev.json  --token xxxx

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request or report any issues you encounter.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Azure DevOps SDK](https://learn.microsoft.com/en-us/javascript/api/overview/azure/devops)
- [React](https://reactjs.org/)
- [Azure DevOps UI](https://github.com/microsoft/azure-devops-ui)

## Some ScreenShot

