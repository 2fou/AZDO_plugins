import React, { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { JsonPatchOperation, Operation } from 'azure-devops-node-api/interfaces/common/VSSInterfaces';
import { showRootComponent } from '../Common/Common';
import { IExtensionDataService, IProjectPageService, CommonServiceIds, getClient } from 'azure-devops-extension-api';
import { WorkItemTrackingRestClient, WorkItem, IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { CoreRestClient } from 'azure-devops-extension-api/Core';

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

const CreateWorkItemsComponent: React.FC = () => {
  const [config, setConfig] = useState<Config | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isCreating, setIsCreating] = useState(false);
  const [orgUrl, setOrgUrl] = useState<string | null>(null);
  const [isInitialized, setIsInitialized] = useState(false);

  useEffect(() => {
    const initializeSDK = async () => {
      console.log("Initializing SDK...");
      try {
        await SDK.init();
        await SDK.ready();
        console.log("SDK initialization complete.");
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        console.log('Extension DataService:', extDataService);
        const locationService = await SDK.getService<any>(CommonServiceIds.LocationService);
        const dataManager = await extDataService.getExtensionDataManager(
          SDK.getExtensionContext().id,
          await SDK.getAccessToken()
        );

        const savedConfig = await dataManager.getValue<Config>('workItemConfig') || DEFAULT_CONFIG;
        console.log('Saved config:', savedConfig);
        const retrievedOrgUrl = await locationService.getResourceAreaLocation(CoreRestClient.RESOURCE_AREA_ID);
        console.log("Configuration and organization URL retrieved.", retrievedOrgUrl);
        setConfig(savedConfig);
        setOrgUrl(retrievedOrgUrl);
        setIsInitialized(true);
        console.log("SDK initialization complete with config:", savedConfig);

      } catch (err) {
        console.error("Initialization error:", err);
        setError('Failed to load configuration. Using default settings.');
        setConfig(DEFAULT_CONFIG);
        setIsInitialized(true);
      }
    };

    initializeSDK();
  }, []);

  useEffect(() => {
    if (isInitialized) {
      SDK.register("createHierarchy", handleCreate);
      console.log("Event handler registered after initialization");
    }
  }, [isInitialized]);

  const createWorkItem = async (
    type: string,
    title: string,
    parentId?: number
  ): Promise<WorkItem> => {
    if (!orgUrl) throw new Error("Organization URL is not available.");

    const client = getClient(WorkItemTrackingRestClient);
    const patchDoc: JsonPatchOperation[] = [{
      op: Operation.Add,
      path: '/fields/System.Title',
      value: title
    }];

    if (parentId) {
      patchDoc.push({
        op: Operation.Add,
        path: '/relations/-',
        value: {
          rel: 'System.LinkTypes.Hierarchy-Reverse',
          url: `${orgUrl}_apis/wit/workItems/${parentId}`
        }
      });
    }

    const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
    const project = await projectService.getProject();

    return client.createWorkItem(
      patchDoc,
      project!.id!,
      type,
      false,
      false
    );
  };

  const handleCreate = async () => {
    console.log("Is config available...", config);
    if (!isInitialized || !config) {
      alert('System not ready. Please wait for initialization to complete.');
      return;
    }

    setIsCreating(true);
    try {
      const workItemService = await SDK.getService<IWorkItemFormService>(
        WorkItemTrackingServiceIds.WorkItemFormService
      );
      const epicId = await workItemService.getId();
      const feature = await createWorkItem('Feature', config.featureTitle, epicId);
      console.log("Feature created", feature.id);
      for (const story of config.stories) {
        const storyItem = await createWorkItem('Product Backlog Item', story.title, feature.id);
        console.log("Story created", storyItem.id);
        for (const taskTitle of story.tasks) {
          const taskItem = await createWorkItem('Task', taskTitle, storyItem.id);
          console.log("Task created", taskItem.id);
        }
      }

      SDK.notifyLoadSucceeded();
      alert('Hierarchy created successfully!');
    } catch (err) {
      console.error("Creation error:", err);
      setError('Failed to create work items. Check console for details.');
    } finally {
      setIsCreating(false);
    }
  };

  return (
    <div className="flex-column">
      {error && <div className="error-message">{error}</div>}
      <Button
        text={isCreating ? "Creating..." : "Create Work Items"}
        onClick={handleCreate}
        disabled={!isInitialized || isCreating}
      />
    </div>
  );
};

showRootComponent(<CreateWorkItemsComponent />, 'CreateWorkItemsComponent-root');
export default CreateWorkItemsComponent;
