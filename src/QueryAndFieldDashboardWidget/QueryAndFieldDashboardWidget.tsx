import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService, getClient, CommonServiceIds } from 'azure-devops-extension-api';
import { WorkItemTrackingRestClient, WorkItem } from 'azure-devops-extension-api/WorkItemTracking';
import { showRootComponent, decodeHtmlEntities } from '../Common/Common';
import * as Dashboard from "azure-devops-extension-api/Dashboard";

interface WorkItemProgress {
    id: number;
    title: string;
    progress: number;
    completed: number;
    total: number;
    status: string; // Added for filtering purposes
}

interface IState {
    title: string;
    loading: boolean;
    error: string | null;
    workItemsProgress: WorkItemProgress[];
    filterStatus: string | null;
    availableStatuses: string[]; // New state for holding dynamic statuses
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
            workItemsProgress: [],
            filterStatus: null,
            availableStatuses: [] // Initialize an empty array for statuses
        };
        // Bind methods to ensure the proper `this` context
        this.mapWorkItemToProgress = this.mapWorkItemToProgress.bind(this);
    }

    componentDidMount() {
        console.log("Component mounted, initializing SDK...");
        SDK.init().then(() => {
            SDK.register('query-and-field-dashboard-widget', this);
            console.log("SDK initialized and widget registered.");
        });
    }

    render(): JSX.Element {
        const { title, loading, error, workItemsProgress, filterStatus, availableStatuses } = this.state;

        if (loading) {
            return <div>Loading work item progress...</div>;
        }

        if (error) {
            return <div>Error: {error}</div>;
        }

        return (
            <div>
                <h3>{title || "Work Item Progress"}</h3>
                <label htmlFor="statusFilter">Filter by Status: </label>
                <select
                    id="statusFilter"
                    value={filterStatus || ""}
                    onChange={(e) => this.setState({ filterStatus: e.target.value })}
                >
                    <option value="">All Statuses</option>
                    {availableStatuses.map((status) => (
                        <option key={status} value={status}>{status}</option>
                    ))}
                </select>
                {workItemsProgress.length === 0 ? (
                    <div>No work items to display.</div>
                ) : (
                    workItemsProgress
                        .filter(item => !filterStatus || item.status === filterStatus)
                        .map(({ id, title, progress, completed, total }) => (
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
                        ))
                )}
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

        try {
            const projectInfo = await this.fetchProjectInfo();
            if (projectInfo) {
                console.log("Project info fetched:", projectInfo);
                await this.fetchAvailableStatuses(projectInfo.id); // Fetch statuses first
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

    private async fetchAvailableStatuses(projectId: string) {
        try {
            const client = getClient(WorkItemTrackingRestClient);
            const workItemTypes = await client.getWorkItemTypes(projectId);

            const allStates: Set<string> = new Set();

            for (const type of workItemTypes) {
                const states = await client.getWorkItemTypeStates(projectId, type.name);
                states.forEach(state => allStates.add(state.name));
            }

            this.setState({ availableStatuses: Array.from(allStates) });
            console.log("Fetched statuses:", Array.from(allStates));
        } catch (error) {
            console.error("Error fetching available statuses:", error);
        }
    }

    private async fetchAndProcessWorkItems(projectId: string) {
        try {
            const client = getClient(WorkItemTrackingRestClient);
            console.log("Fetching work items...");
            const wiqlQuery = {
                query: `SELECT [System.Id], [System.Title], [Custom.AnswersField], [System.State] FROM WorkItems WHERE [System.TeamProject] = @project AND [Custom.AnswersField] Is Not Empty ORDER BY [System.Id]`,
                parameters: { project: projectId }
            };
            const queryResult = await client.queryByWiql(wiqlQuery, projectId);
            const workItemRefs = queryResult.workItems;
            console.log("Number of work item references:", workItemRefs.length);

            if (workItemRefs.length > 0) {
                const workItems = await client.getWorkItems(workItemRefs.map(wi => wi.id), undefined, ['System.Title', 'Custom.AnswersField', 'System.State']);
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

   private mapWorkItemToProgress = (workItem: WorkItem): WorkItemProgress => {
    try {
        const fieldData = workItem.fields['Custom.AnswersField'];
        const status = workItem.fields['System.State'];

        if (fieldData) {
            const { completedEntriesCount, totalQuestionCount } = this.calculateProgress(fieldData);
            const progress = totalQuestionCount > 0 ? (completedEntriesCount / totalQuestionCount) * 100 : 0;

            console.log(`Work item ID ${workItem.id} mapped with progress: ${progress.toFixed(2)}%`);
            return {
                id: workItem.id,
                title: workItem.fields['System.Title'] || "Untitled",
                progress,
                completed: completedEntriesCount,
                total: totalQuestionCount,
                status
            };
        } else {
            console.warn(`Work item ID ${workItem.id} does not have the 'Custom.AnswersField' field.`);
        }
    } catch (error) {
        console.error("Error mapping work item to progress:", error);
    }

    return {
        id: workItem.id,
        title: workItem.fields['System.Title'] || "Untitled",
        progress: 0,
        completed: 0,
        total: 0,
        status: status || "Unknown"
    };
};
    private calculateProgress(fieldData: any): { completedEntriesCount: number, totalQuestionCount: number } {
        const decodedValue = decodeHtmlEntities(fieldData as string);
        const answers = JSON.parse(decodedValue);

        if (!answers?.data) {
            throw new Error("Decoded answers are undefined or incorrectly structured.");
        }

        let completedEntriesCount = 0;
        let totalQuestionCount = 0;

        // Iterate over each question in the "data" object
        for (const answerKey in answers.data) {
            const answer = answers.data[answerKey];

            if (answer.checked) {
                // Count the number of '1's in the binary representations of `uniqueResult` and `totalWeight`
                completedEntriesCount = (answer.uniqueResult || 0).toString(2).split('1').length - 1;
                totalQuestionCount = (answer.totalWeight || 0).toString(2).split('1').length - 1;

                break; // Since only the checked entry determines the completeness
            }
        }

        return { completedEntriesCount, totalQuestionCount };
    }
}

showRootComponent(<QueryAndFieldDashboardWidget />, "query-and-field-dashboard-root");