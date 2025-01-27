import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService, getClient, CommonServiceIds } from 'azure-devops-extension-api';
import { WorkItemTrackingRestClient, WorkItem } from 'azure-devops-extension-api/WorkItemTracking';
import { AnswerDetail, showRootComponent, decodeHtmlEntities } from '../Common/Common';
import * as Dashboard from "azure-devops-extension-api/Dashboard";

interface WorkItemProgress {
    id: number;
    title: string;
    progress: number;
    completed: number;
    total: number;
}

interface IState {
    title: string;
    loading: boolean;
    error: string | null;
    workItemsProgress: WorkItemProgress[];
}

interface IWidgetSettings {
    customSettings: {
        data: string;
    };
}

class QueryAndFieldDashboardWidget extends React.Component<{}, IState> implements Dashboard.IConfigurableWidget {
    constructor(props: {}) {
        super(props);
        this.state = {
            title: '',
            loading: true,
            error: null,
            workItemsProgress: []
        };
    }

    componentDidMount() {
        console.log("Component mounted, initializing SDK...");
        SDK.init().then(() => {
            SDK.register('query-and-field-dashboard-widget', this);
            console.log("SDK initialized and widget registered.");
        });
    }

    render(): JSX.Element {
        const { title, loading, error, workItemsProgress } = this.state;
        if (loading) {
            return <div>Loading work item progress...</div>;
        }

        if (error) {
            return <div>Error: {error}</div>;
        }

        return (
            <div>
                <h3>{title || "Work Item Progress"}</h3>
                {workItemsProgress.map(({ id, title, progress, completed, total }) => (
                    <div key={id}>
                        <h4>{title}</h4>
                        <div style={{ width: '100%', backgroundColor: '#e0e0e0', borderRadius: '8px', marginBottom: '10px' }}>
                            <div
                                style={{
                                    width: `${progress}%`,
                                    backgroundColor: '#76c7c0',
                                    height: '24px',
                                    borderRadius: '8px'
                                }}
                            />
                        </div>
                        <div>{`Work Item ID: ${id}, Progress: ${progress.toFixed(2)}%, Completed ${completed} out of ${total}`}</div>
                    </div>
                ))}
            </div>
        );
    }

    async preload(): Promise<Dashboard.WidgetStatus> {
        console.log("Preloading widget...");
        return Dashboard.WidgetStatusHelper.Success();
    }

    async load(widgetSettings: Dashboard.WidgetSettings): Promise<Dashboard.WidgetStatus> {
        console.log("Loading widget with settings:", widgetSettings);
        try {
            await this.initializeAndFetchData(widgetSettings);
            console.log("Widget loaded successfully.");
            return Dashboard.WidgetStatusHelper.Success();
        } catch (e) {
            console.error("Loading widget failed:", e);
            return Dashboard.WidgetStatusHelper.Failure((e as any).toString());
        }
    }

    async reload(widgetSettings: Dashboard.WidgetSettings): Promise<Dashboard.WidgetStatus> {
        console.log("Reloading widget with settings:", widgetSettings);
        return this.load(widgetSettings);
    }

    private async initializeAndFetchData(widgetSettings: Dashboard.WidgetSettings) {
        console.log("Initializing data fetch...");
        this.setState({
            title: widgetSettings.name || 'Default Title',
        });

        const customData: IWidgetSettings = JSON.parse(widgetSettings.customSettings.data || '{}');
        console.log("Custom settings data:", customData);

        try {
            const projectInfo = await this.fetchProjectInfo();
            if (projectInfo) {
                console.log("Project info fetched:", projectInfo);
                await this.fetchAndProcessWorkItems(projectInfo.id);
            } else {
                console.warn("No project info found.");
            }
        } catch (err) {
            console.error("Error during data fetch:", err);
            this.setState({ error: "Failed to load data." });
        } finally {
            this.setState({ loading: false });
        }
    }

    private async fetchProjectInfo() {
        try {
            const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            const project = await projectService.getProject();
            console.log("Fetched project info:", project);
            return project;
        } catch (error) {
            console.error("Error fetching project info:", error);
            throw error;
        }
    }

    private async fetchAndProcessWorkItems(projectId: string) {
        try {
            const client = getClient(WorkItemTrackingRestClient);
            console.log("Fetching work items...");
            const wiqlQuery = {
                query: `SELECT [System.Id], [System.Title], [Custom.AnswersField] FROM WorkItems WHERE [System.TeamProject] = @project AND [Custom.AnswersField] Is Not Empty ORDER BY [System.Id]`,
                parameters: { project: projectId }
            };            
            const queryResult = await client.queryByWiql(wiqlQuery, projectId);
            const workItemRefs = queryResult.workItems;
            console.log("Number of work item references:", workItemRefs.length);

            if (workItemRefs.length > 0) {
                const workItems = await client.getWorkItems(workItemRefs.map(wi => wi.id), undefined, ['System.Title', 'Custom.AnswersField']);
                console.log(`Fetched ${workItems.length} work items.`);
                const progressData = workItems.map(this.mapWorkItemToProgress);
                this.setState({ workItemsProgress: progressData });
            } else {
                console.warn("No work item references were found.");
            }
        } catch (error) {
            console.error("Error fetching and processing work items:", error);
        }
    }

private mapWorkItemToProgress(workItem: WorkItem): WorkItemProgress {
    try {
        const fieldData = workItem.fields['Custom.AnswersField'];
        if (fieldData) {
            const decodedValue = decodeHtmlEntities(fieldData as string);
            const answers: { [key: string]: AnswerDetail } = JSON.parse(decodedValue);

            const totalEntries = Object.keys(answers).length;
            let completedEntriesCount = 0;

            // Iterate over each AnswerDetail
            for (const answer of Object.values(answers)) {
                if (answer.entries && answer.entries.length > 0) {
                    // If there are entries, check if all are non-empty
                    if (answer.entries.every(entry => Boolean(entry.value))) {
                        completedEntriesCount += 1;
                    }
                }
            }

            const progress = totalEntries > 0 ? (completedEntriesCount / totalEntries) * 100 : 0;

            console.log(`Work item ID ${workItem.id} mapped with progress: ${progress.toFixed(2)}%`);
            return {
                id: workItem.id,
                title: workItem.fields['System.Title'] || "Untitled",
                progress,
                completed: completedEntriesCount,
                total: totalEntries
            };
        } else {
            console.warn(`Work item ID ${workItem.id} does not have the 'Custom.AnswersField' field.`);
        }
    } catch (parseError) {
        console.error("Error parsing field data:", parseError);
    }

    return {
        id: workItem.id,
        title: workItem.fields['System.Title'] || "Untitled",
        progress: 0,
        completed: 0,
        total: 0
    };
}
}

showRootComponent(<QueryAndFieldDashboardWidget />, "query-and-field-dashboard-root");