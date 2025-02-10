import * as SDK from 'azure-devops-extension-sdk';
import { getClient, IProjectPageService, CommonServiceIds, IExtensionDataService} from 'azure-devops-extension-api';
import { WorkItemTrackingRestClient, WorkItem, IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { CoreRestClient } from 'azure-devops-extension-api/Core';
import { JsonPatchOperation, Operation } from 'azure-devops-node-api/interfaces/common/VSSInterfaces';



SDK.init();


interface Config {
  featureTitle: string;
  stories: Array<{
    title: string;
    tasks: string[];
  }>;
}

const DEFAULT_CONFIG: Config = {
  featureTitle: 'Default Feature',
  stories: []
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
        // Retrieve configuration using the data manager
        const dataManager = await extDataService.getExtensionDataManager(
          SDK.getExtensionContext().id,
          await SDK.getAccessToken()
        );
        
        const config = await dataManager.getValue<Config>('workItemConfig') || DEFAULT_CONFIG;
        console.log('Config:', config);
        const workItemService = await SDK.getService<IWorkItemFormService>(
          WorkItemTrackingServiceIds.WorkItemFormService
        );
        const epicId = await workItemService.getId();
        console.log('Epic Id:', epicId);
        const feature = await createWorkItem(orgUrl, { type: 'Feature', title: config.featureTitle, parentId: epicId });
        console.log('Feature created:', feature.id);

        for (const story of config.stories) {
          const storyItem = await createWorkItem(orgUrl, { type: 'Product Backlog Item', title: story.title, parentId: feature.id });
          console.log('Story created:', storyItem.id);
          for (const taskTitle of story.tasks) {
            const taskItem = await createWorkItem(orgUrl, { type: 'Task', title: taskTitle, parentId: storyItem.id });
            console.log('Task created:', taskItem.id);
          }
        }

        alert('Hierarchy created successfully!');
      } catch (err) {
        console.error('Execution error:', err);
        alert('Failed to create work items. Check console for details.');
      }
    }
  };
});