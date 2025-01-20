import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { useState, useEffect } from 'react';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import { Button } from 'azure-devops-ui/Button';
import { IWorkItemFormNavigationService } from "azure-devops-extension-api/WorkItemTracking";
import { showRootComponent } from '../Common/Common';


const CreateWorkItemsComponent: React.FC = () => {
    const [config, setConfig] = useState<{
        featureTitle: string;
        storyTitle: string;
        taskTitles: string[];
    } | null>(null);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        const initializeComponent = async () => {
            try {
                console.log('Initializing SDK...');
                await SDK.init();
                console.log('SDK initialized.');

                const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
                const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
                const savedConfig = await dataManager.getValue<any>('workItemConfig');
                setConfig(savedConfig);
            } catch (err) {
                console.error('Failed to load configuration:', err);
                setError('Failed to load configuration.');
            }
        };
        initializeComponent();
    }, []);

    const handleCreateWorkItems = async () => {
        if (!config) {
            setError('No configuration found.');
            return;
        }

        try {
                // Get the Work Item Form Navigation Service
                const navigationService = await SDK.getService<IWorkItemFormNavigationService>("ms.vss-work-web.work-item-form-navigation");
                // Open the editor for a new work item of type "Feature"
                navigationService.openNewWorkItem("Feature", {
        "System.Title": config.featureTitle,
        "System.Description": "Description of the new feature"
    });

    navigationService.openNewWorkItem("User Story", {
        "System.Title": config.storyTitle,
        "System.Description": "Description of the new story"
    });

          

            // Create Tasks
            for (const taskTitle of config.taskTitles) {
      
                navigationService.openNewWorkItem("Task", {
                    "System.Title": taskTitle,
                    "System.Description": "Description of the new task"
                });
            }

            alert('Work items created successfully!');
        } catch (err) {
            console.error('Failed to create work items:', err);
            setError('Failed to create work items.');
        }
    };

    return (
        <div>
            <h2>Create Work Items</h2>
            {error && <div className="error">{error}</div>}
            <Button text="Create Work Items" onClick={handleCreateWorkItems} disabled={!config} />
        </div>
    );
};

showRootComponent(<CreateWorkItemsComponent />, 'extension-root');
export default CreateWorkItemsComponent;
