import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { IProjectPageService, IProjectInfo, CommonServiceIds } from "azure-devops-extension-api";
import { getClient } from "azure-devops-extension-api";
import { WorkItemTrackingRestClient, WorkItem } from "azure-devops-extension-api/WorkItemTracking";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { DropdownSelection } from "azure-devops-ui/Utilities/DropdownSelection";

interface IWorkItemPickerProps {
    onWorkItemSelected: (workItemId: number) => void;
}

export const WorkItemPicker: React.FC<IWorkItemPickerProps> = (props) => {
    const [workItems, setWorkItems] = React.useState<WorkItem[]>([]);
    const [project, setProject] = React.useState<IProjectInfo | undefined>();
    const selection = new DropdownSelection();

    React.useEffect(() => {
        const fetchProject = async () => {
            await SDK.init();
            const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            const projectInfo = await projectService.getProject();
            setProject(projectInfo);
        };

        fetchProject();
    }, []);

    React.useEffect(() => {
        const fetchWorkItems = async () => {
            if (project) {
                const client = getClient(WorkItemTrackingRestClient);
                const wiqlQuery = {
                    query: `SELECT [System.Id], [System.Title], [System.State] FROM WorkItems WHERE [System.TeamProject] = @project  ORDER BY [System.Id]`
                };
                const workItemQueryResult = await client.queryByWiql(wiqlQuery, project.id);
                const workItemRefs = workItemQueryResult.workItems;
                if (workItemRefs && workItemRefs.length > 0) {
                    const workItems = await client.getWorkItems(workItemRefs.map(wi => wi.id));
                    setWorkItems(workItems);
                }
            }
        };

        fetchWorkItems();
    }, [project]);

    return (
        <div>
            <Dropdown
                placeholder="Select a Work Item"
                items={workItems.map(item => ({
                    id: item.id.toString(),
                    text: `${item.id} - ${item.fields["System.Title"]}`
                }))}
                selection={selection}
                onSelect={(event, item) => {
                    if (item) {
                        props.onWorkItemSelected(parseInt(item.id));
                    }
                }}
            />
        </div>
    );
};