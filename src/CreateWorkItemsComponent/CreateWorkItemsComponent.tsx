import * as SDK from 'azure-devops-extension-sdk';
import { getClient, IProjectPageService, CommonServiceIds, IExtensionDataService } from 'azure-devops-extension-api';
import { WorkItemTrackingRestClient, WorkItem, IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { CoreRestClient } from 'azure-devops-extension-api/Core';
import { JsonPatchOperation, Operation } from 'azure-devops-node-api/interfaces/common/VSSInterfaces';

interface StoryConfig {
    title: string;
    tasks: string[];
}

interface WorkItemTypesConfig {
    epic: string;
    feature: string;
    story: string;
    task: string;
}

interface Config {
    featureTitle: string;
    stories: StoryConfig[];
    workItemTypes: WorkItemTypesConfig;
}

const DEFAULT_CONFIG: Config = {
    featureTitle: 'Default Feature',
    stories: [],
    workItemTypes: {
        epic: 'Epic',
        feature: 'Feature',
        story: 'Product Backlog Item',
        task: 'Task'
    }
};

SDK.init();

// Before saving, perform checks to see if the form is dirty
const checkIfFormIsDirty = async (workItemService: IWorkItemFormService) => {
    const isDirty = await workItemService.isDirty();
    console.log('Is form dirty:', isDirty);

    if (!isDirty) {
        console.log('No changes detected, skipping save.');
        return false;
    }

    return true;
};

const createWorkItem = async (orgUrl: string, config: { type: string; title: string; parentId?: number; }): Promise<WorkItem> => {
    const client = getClient(WorkItemTrackingRestClient);
    const patchDoc: JsonPatchOperation[] = [
        { op: Operation.Add, path: '/fields/System.Title', value: config.title }
    ];

    if (config.parentId) {
        patchDoc.push({
            op: Operation.Add,
            path: '/relations/-',
            value: {
                rel: 'System.LinkTypes.Hierarchy-Reverse',
                url: `${orgUrl}_apis/wit/workItems/${config.parentId}`
            }
        });
    }

    const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
    const project = await projectService.getProject();
    if (!project) throw new Error('Project is not available.');

    return client.createWorkItem(patchDoc, project.id, config.type);
};

const createWorkItemWithChildren = async (
    orgUrl: string,
    parentType: string,
    parentId: number,
    childrenConfigs: StoryConfig[],
    childType: string,
    grandchildType: string | null = null
) => {
    for (const childConfig of childrenConfigs) {
        const childItem = await createWorkItem(orgUrl, { type: childType, title: childConfig.title, parentId });
        console.log(`${childType} created:`, childItem.id);

        if (grandchildType) {
            for (const grandchildTitle of childConfig.tasks) {
                const grandchildItem = await createWorkItem(orgUrl, { type: grandchildType, title: grandchildTitle, parentId: childItem.id });
                console.log(`${grandchildType} created:`, grandchildItem.id);
            }
        }
    }
};

SDK.register("createHierarchy", () => {
    return {
        execute: async (context: any) => {
            try {
                console.log('SDK is going to init!');
                await SDK.ready();
                console.log('SDK ready!');
                const locationService = await SDK.getService<any>(CommonServiceIds.LocationService);
                const orgUrl = await locationService.getResourceAreaLocation(CoreRestClient.RESOURCE_AREA_ID);
                console.log('Origine URL:', orgUrl);
                const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);

                const dataManager = await extDataService.getExtensionDataManager(
                    SDK.getExtensionContext().id,
                    await SDK.getAccessToken()
                );

                const config = await dataManager.getValue<Config>('workItemConfig') || DEFAULT_CONFIG;
                console.log('Config:', config);
                const workItemService = await SDK.getService<IWorkItemFormService>(
                    WorkItemTrackingServiceIds.WorkItemFormService
                );

                const currentWorkItemType = await workItemService.getFieldValue('System.WorkItemType', { returnOriginalValue: true }) as string;
                const currentItemId = await workItemService.getId();
                console.log('Current item type %s id %d', currentWorkItemType, currentItemId);

switch (currentWorkItemType) {
    case config.workItemTypes.epic: {
        console.log('Current item is an Epic.');

        // Create a Feature under the Epic
        const featureItem = await createWorkItem(orgUrl, {
            type: config.workItemTypes.feature,
            title: config.featureTitle,
            parentId: currentItemId
        });

        console.log(`Feature "${config.featureTitle}" created:`, featureItem.id);

        // For each story, create a story under the feature and tasks under each story
        for (const storyConfig of config.stories) {
            const storyItem = await createWorkItem(orgUrl, {
                type: config.workItemTypes.story,
                title: storyConfig.title,
                parentId: featureItem.id
            });
            console.log(`Story "${storyConfig.title}" created:`, storyItem.id);

            for (const taskTitle of storyConfig.tasks) {
                const taskItem = await createWorkItem(orgUrl, {
                    type: config.workItemTypes.task,
                    title: taskTitle,
                    parentId: storyItem.id
                });
                console.log(`Task "${taskTitle}" created:`, taskItem.id);
            }
        }
        break;
    }

    case config.workItemTypes.feature: {
        console.log('Current item is a Feature.');
        await createWorkItemWithChildren(orgUrl, currentWorkItemType, currentItemId, config.stories, config.workItemTypes.story, config.workItemTypes.task);
        break;
    }

    case config.workItemTypes.story: {
        console.log('Current item is a Story.');
        await createWorkItemWithChildren(orgUrl, currentWorkItemType, currentItemId, config.stories, config.workItemTypes.task);
        break;
    }

    default:
        alert('Creation is only allowed from Epics, Features, or User Stories.');
        break;
}
                // Check if the form is dirty before saving
                const shouldSave = await checkIfFormIsDirty(workItemService);
                if (shouldSave) {
                    await workItemService.save();
                    await workItemService.refresh();
                    console.log('Form saved and refreshed.');
                    alert('Hierarchy processed successfully!');
                } else {
                    console.log('No changes to save.');
                }
            } catch (err) {
                console.error('Execution error:', err);
                alert('Failed to create work items. Check console for details.');
            }
        }
    };
});