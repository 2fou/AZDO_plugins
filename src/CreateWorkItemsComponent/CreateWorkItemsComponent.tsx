import * as SDK from 'azure-devops-extension-sdk';
import {
  getClient,
  IProjectPageService,
  CommonServiceIds,
  IExtensionDataService
} from 'azure-devops-extension-api';
import {
  WorkItemTrackingRestClient,
  WorkItem,
  IWorkItemFormService,
  WorkItemTrackingServiceIds
} from 'azure-devops-extension-api/WorkItemTracking';
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

const checkIfFormIsDirty = async (workItemService: IWorkItemFormService) => {
  const isDirty = await workItemService.isDirty();
  console.log('Is form dirty:', isDirty);

  if (!isDirty) {
    console.log('No changes detected, skipping save.');
    return false;
  }

  return true;
};

const createWorkItem = async (
  orgUrl: string,
  config: { type: string; title: string; parentId?: number }
): Promise<WorkItem> => {
  const client = getClient(WorkItemTrackingRestClient);
  const patchDoc: JsonPatchOperation[] = [
    {
      op: Operation.Add,
      path: '/fields/System.Title',
      value: config.title
    }
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

  const projectService = await SDK.getService<IProjectPageService>(
    CommonServiceIds.ProjectPageService
  );
  const project = await projectService.getProject();
  if (!project) throw new Error('Project is not available.');

  return client.createWorkItem(patchDoc, project.id, config.type);
};

const createWorkItemWithChildren = async (
  orgUrl: string,
  parentType: string,
  parentId: number,
  childrenConfigs: StoryConfig[],
  directChildType: string,
  grandchildType?: string
) => {
  if (childrenConfigs.length === 0) {
    console.log(
      `No children configurations found for ${parentType}. Skipping child creation.`
    );
    return;
  }

  for (const childConfig of childrenConfigs) {
    // If no "grandchildType", create tasks directly under the parent
    if (!grandchildType) {
      for (const taskTitle of childConfig.tasks) {
        const taskItem = await createWorkItem(orgUrl, {
          type: directChildType,
          title: taskTitle,
          parentId
        });
        console.log(`${directChildType} created:`, taskItem.id);
      }
    } else {
      // Create the child item, then tasks under it
      const childItem = await createWorkItem(orgUrl, {
        type: directChildType,
        title: childConfig.title,
        parentId
      });
      console.log(`${directChildType} created:`, childItem.id);

      for (const grandchildTitle of childConfig.tasks) {
        const grandchildItem = await createWorkItem(orgUrl, {
          type: grandchildType,
          title: grandchildTitle,
          parentId: childItem.id
        });
        console.log(`${grandchildType} created:`, grandchildItem.id);
      }
    }
  }
};

SDK.register('createHierarchy', () => {
  return {
    execute: async (context: any) => {
      try {
        console.log('SDK is going to init!');
        await SDK.ready();
        console.log('SDK ready!');

        const locationService = await SDK.getService<any>(
          CommonServiceIds.LocationService
        );
        const orgUrl = await locationService.getResourceAreaLocation(
          CoreRestClient.RESOURCE_AREA_ID
        );
        console.log('Origin URL:', orgUrl);

        // Retrieve the saved extension config (including custom WIT names)
        const extDataService = await SDK.getService<IExtensionDataService>(
          CommonServiceIds.ExtensionDataService
        );
        const dataManager = await extDataService.getExtensionDataManager(
          SDK.getExtensionContext().id,
          await SDK.getAccessToken()
        );

        const config =
          (await dataManager.getValue<Config>('workItemConfig')) ||
          DEFAULT_CONFIG;
        console.log('Config:', config);

        const workItemService = await SDK.getService<IWorkItemFormService>(
          WorkItemTrackingServiceIds.WorkItemFormService
        );

        const currentWorkItemType = (await workItemService.getFieldValue(
          'System.WorkItemType',
          { returnOriginalValue: true }
        )) as string;
        const currentItemId = await workItemService.getId();
        console.log(
          'Current item type %s id %d',
          currentWorkItemType,
          currentItemId
        );

        // Use the user-defined WIT from config.workItemTypes
        switch (currentWorkItemType) {
          case config.workItemTypes.epic: {
            console.log('Current item is an Epic.');

            const featureItem = await createWorkItem(orgUrl, {
              type: config.workItemTypes.feature,
              title: config.featureTitle,
              parentId: currentItemId
            });
            console.log(
              `Feature "${config.featureTitle}" created:`,
              featureItem.id
            );

            // Create stories and tasks under the new Feature
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
            await createWorkItemWithChildren(
              orgUrl,
              currentWorkItemType,
              currentItemId,
              config.stories,
              config.workItemTypes.story,
              config.workItemTypes.task
            );
            break;
          }

          case config.workItemTypes.story: {
            console.log('Current item is a Story.');
            // Directly create tasks under the current Story
            for (const taskTitle of config.stories.flatMap(
              (story) => story.tasks
            )) {
              const taskItem = await createWorkItem(orgUrl, {
                type: config.workItemTypes.task,
                title: taskTitle,
                parentId: currentItemId
              });
              console.log(`Task "${taskTitle}" created:`, taskItem.id);
            }
            break;
          }

          default:
            alert('Creation is only allowed from Epics, Features, or User Stories.');
            return;
        }

        // Finally, save & refresh if needed
        const shouldSave = await checkIfFormIsDirty(workItemService);
        if (shouldSave) {
          await workItemService.save();
          await workItemService.refresh();
          console.log('Form saved and refreshed.');
        } else {
          console.log('No changes to save.');
        }
        // Always show the success message, regardless of whether the form needed saving
        alert('All items have been created successfully!');
      } catch (err) {
        console.error('Execution error:', err);
        alert('Failed to create work items. Check console for details.');
      }
    }
  };
});