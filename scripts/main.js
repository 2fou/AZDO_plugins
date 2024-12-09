import { IWorkItemFormService, WorkItemTrackingServiceIds } from "azure-devops-extension-api/WorkItemTracking";
import * as SDK from "azure-devops-extension-sdk";

document.addEventListener('DOMContentLoaded', () => {
    if (SDK) {
        console.log("SDK is defined. Trying to initialize...");
        SDK.init();

        SDK.ready().then(async () => {
            try {
                console.log("SDK is ready");

                const question1 = document.getElementById('question1');
                const question2 = document.getElementById('question2');
                const commonDeliverable = document.getElementById('commonDeliverable');
                const additionalDeliverable = document.getElementById('additionalDeliverable');
                const saveButton = document.getElementById('saveButton');

                // Function to update visibility of deliverables
                const updateDeliverablesVisibility = () => {
                    commonDeliverable.style.display = (question1.checked || question2.checked) ? 'block' : 'none';
                    additionalDeliverable.style.display = question2.checked ? 'block' : 'none';
                };

                // Function to update fields in Azure DevOps
                const updateFields = async (workItemFormService) => {
                    try {
                        const data = {
                            question1Checked: question1.checked,
                            question2Checked: question2.checked,
                            securityPlanUrl: document.getElementById('securityPlanUrl').value,
                            legalValidationUrl: document.getElementById('legalValidationUrl').value
                        };

                        let updatedFields = {};
                        if (data.question1Checked || data.question2Checked) {
                            updatedFields['Custom.SecurityPlanUrl'] = data.securityPlanUrl;
                        }
                        if (data.question2Checked) {
                            updatedFields['Custom.LegalValidationUrl'] = data.legalValidationUrl;
                        }

                        await workItemFormService.setFieldValues(updatedFields);
                        console.log('Fields updated successfully');
                    } catch (error) {
                        console.error('Error updating fields: ', error);
                    }
                };

                question1.addEventListener('change', updateDeliverablesVisibility);
                question2.addEventListener('change', updateDeliverablesVisibility);

                saveButton.addEventListener('click', async () => {
                    try {
                        const workItemFormService = await SDK.getService(IWorkItemFormService, WorkItemTrackingServiceIds.WorkItemFormService);
                        await updateFields(workItemFormService);
                    } catch (error) {
                        console.error('Error getting work item form service:', error);
                    }
                });
            } catch (error) {
                console.error('Error initializing services:', error);
            }
        });
    } else {
        console.log('SDK is not defined');
    }
});