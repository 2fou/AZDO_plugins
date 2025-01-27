# Azure DevOps Extension: Questionnaire Form

This project is an Azure DevOps extension that provides a questionnaire form for work items. It enables users to answer a series of questions specific to each work item and store the responses directly in a custom field of the work item.

https://marketplace.visualstudio.com/items?itemName=AgileCoachingSolutionsFeatureAddon.feature-livrable-extension&targetId=9ed4a67e-0bf4-48ce-9276-8750bc557cd4&utm_source=vstsproduct&utm_medium=ExtHubManageList


## Features

- Load predefined questions from Azure DevOps Extension Data Service.
- Allow users to provide answers and optional links related to each question.
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

## Usage

- Setup a new custom field "AnswersField" in a process, this field can be hidden from the work item type.
- Add the Livrable Control to the workitem type
- Add questions to the project through project/setting/extension
- Navigate to a work item in your Azure DevOps project.
- Use the integrated tab to fill out the questionnaire.
- Save your responses, and they will persist in the associated work item custom field.

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

## Some screenshoot

